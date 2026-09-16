"""設定ファイル(YAML)の読み込みと検証。"""
from __future__ import annotations

import os
from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import Dict, List

import yaml

from scanner import DEFAULT_MTIME_TOLERANCE_SEC, HASH_MODE_ALWAYS, HASH_MODES


SUPPORTED_FORMATS = {"excel", "html", "both"}

# HTML レポートの 1 つの表に出す最大行数。
# 設定ファイルに書かれていない場合の既定値 (既に配布済みの config.yaml でも
# 巨大なレポートが開けなくなる問題を避けられるようにするため)。
# 実測: 20,000 行はロード 1.1 秒で操作可能、50,000 行はブラウザが 30 秒以上無応答。
DEFAULT_MAX_TABLE_ROWS = 20_000
SUPPORTED_HASH_ALGOS = {"sha256"}


@dataclass(frozen=True)
class Location:
    name: str
    path: Path


@dataclass(frozen=True)
class OutputConfig:
    format: str
    output_dir: Path
    # HTML の 1 表あたりの最大行数。0 は無制限。Excel には適用しない。
    max_table_rows: int = DEFAULT_MAX_TABLE_ROWS
    # 残す過去レポートの実行回数。0 は削除しない (既定)。
    # 削除は元に戻せないので、明示的に指定したときだけ動かす。
    keep_reports: int = 0


@dataclass(frozen=True)
class PerformanceConfig:
    parallel_workers: int
    hash_algorithm: str
    hash_mode: str = HASH_MODE_ALWAYS
    mtime_tolerance_sec: float = DEFAULT_MTIME_TOLERANCE_SEC


@dataclass(frozen=True)
class MatchingConfig:
    """拠点間で「同じファイル」とみなす条件。"""

    normalize_unicode: bool = True
    case_sensitive: bool = True


@dataclass(frozen=True)
class Comparison:
    """1 組の比較対象。

    `name` が空文字なら、設定ファイルにトップレベルの `locations:` だけが
    書かれている従来の形式 (名前の無い 1 組) を表す。
    """

    name: str
    locations: List[Location]


@dataclass(frozen=True)
class Config:
    comparisons: List[Comparison]
    # 実行対象として選ばれている比較の名前 (従来形式では空文字)
    selected: str = ""
    exclude_patterns: List[str] = field(default_factory=list)
    matching: MatchingConfig = field(default_factory=MatchingConfig)
    output: OutputConfig = field(default_factory=lambda: OutputConfig("excel", Path("./reports")))
    performance: PerformanceConfig = field(
        default_factory=lambda: PerformanceConfig(4, "sha256")
    )

    @property
    def locations(self) -> List[Location]:
        """選ばれている比較の拠点一覧。

        比較対象が 1 組だけだった頃の呼び出し側をそのまま動かすための入口。
        """
        for c in self.comparisons:
            if c.name == self.selected:
                return c.locations
        return self.comparisons[0].locations

    @property
    def comparison_names(self) -> List[str]:
        """定義されている比較の名前。従来形式では空リスト。"""
        return [c.name for c in self.comparisons if c.name]


class ConfigError(ValueError):
    pass


def apply_overrides(
    config: Config,
    *,
    output_format: str | None = None,
    output_dir: str | None = None,
    hash_mode: str | None = None,
) -> Config:
    """コマンドラインからの上書きを設定に反映する。

    CI から実行するときに config.yaml を書き換えずに出力先や形式を変えられるようにする。
    None の項目は設定ファイルの値をそのまま使う。
    `output_dir` は設定ファイル内の相対パスと違い、CWD からの相対として解決する
    (コマンドラインに書いたパスは打った場所から見た相対と考えるのが自然なため)。
    """
    if output_format is not None and output_format not in SUPPORTED_FORMATS:
        raise ConfigError(
            f"--format は {sorted(SUPPORTED_FORMATS)} のいずれかを指定してください "
            f"(現在: {output_format})"
        )
    if hash_mode is not None and hash_mode not in HASH_MODES:
        raise ConfigError(
            f"--hash-mode は {sorted(HASH_MODES)} のいずれかを指定してください "
            f"(現在: {hash_mode})"
        )

    output = config.output
    if output_format is not None or output_dir is not None:
        output = replace(
            output,
            format=output_format if output_format is not None else output.format,
            output_dir=(
                Path(output_dir).resolve() if output_dir is not None else output.output_dir
            ),
        )
    performance = config.performance
    if hash_mode is not None:
        performance = replace(performance, hash_mode=hash_mode)
    return replace(config, output=output, performance=performance)


def _positive_int(value, key: str) -> int:
    """0 以上の整数として検証する (0 は「無制限 / 無効」を意味する)。"""
    if isinstance(value, bool) or not isinstance(value, int):
        raise ConfigError(f"{key} は整数で指定してください (現在: {value!r})")
    if value < 0:
        raise ConfigError(f"{key} は0以上を指定してください")
    return value


def _resolve(p: Path, base: Path) -> Path:
    """相対パスは base からの相対として解決し、絶対パスはそのまま返す。"""
    return p if p.is_absolute() else (base / p).resolve()


# 比較の名前に使えない文字。レポートのファイル名に含めるため、
# パス区切りや Windows のファイル名で使えない文字を弾く。
_INVALID_NAME_CHARS = set('\\/:*?"<>|')


def _parse_locations(items, base_dir: Path, where: str) -> List[Location]:
    """拠点の一覧を検証して組み立てる。`where` はエラーメッセージ用の位置。"""
    if not items or not isinstance(items, list):
        raise ConfigError(f"{where} は1件以上指定してください")

    locations: List[Location] = []
    seen_names: set[str] = set()
    for i, item in enumerate(items):
        if not isinstance(item, dict):
            raise ConfigError(f"{where}[{i}] が不正です")
        name = item.get("name")
        path = item.get("path")
        if not name or not path:
            raise ConfigError(f"{where}[{i}] は name と path が必須です")
        if name in seen_names:
            raise ConfigError(f"拠点名が重複しています: {name} ({where})")
        # バックスラッシュ表記は YAML パース時に潰れて壊れることが多く、
        # またコピー機能でのパス組み立てを単純化するためフォワードスラッシュに統一する。
        path_str = str(path)
        if "\\" in path_str:
            raise ConfigError(
                f"{where}[{i}].path にバックスラッシュが含まれています: {path_str!r}\n"
                f"  → フォワードスラッシュで記述してください "
                f"(例: '//server-a/share/docs')"
            )
        seen_names.add(name)
        locations.append(Location(name=str(name), path=_resolve(Path(path_str), base_dir)))

    if len(locations) < 2:
        raise ConfigError(f"{where} は2件以上必要です")

    # パス重複検出: 同じ実体パスを複数拠点として登録すると常に一致してしまうため弾く。
    # 比較をまたいで同じパスを使うのは問題ないので、検査は 1 組の中だけで行う。
    # os.path.normcase は Windows では大小無視・スラッシュ統一、UNIX では no-op。
    seen_paths: Dict[str, str] = {}
    for loc in locations:
        key = os.path.normcase(os.path.normpath(str(loc.path)))
        if key in seen_paths:
            raise ConfigError(
                f"拠点 '{loc.name}' のパスが拠点 '{seen_paths[key]}' と重複しています: "
                f"{loc.path} ({where})"
            )
        seen_paths[key] = loc.name
    return locations


def _parse_comparisons(raw: dict, base_dir: Path) -> List[Comparison]:
    """`comparisons:` または従来の `locations:` から比較の一覧を組み立てる。"""
    comparisons_raw = raw.get("comparisons")
    if comparisons_raw is None:
        # 従来形式: トップレベルの locations だけ。名前の無い 1 組として扱う。
        return [Comparison(name="", locations=_parse_locations(
            raw.get("locations"), base_dir, "locations"))]

    if raw.get("locations"):
        raise ConfigError(
            "comparisons と locations は同時に指定できません。\n"
            "  → 比較を複数定義する場合は locations を comparisons の中に移してください"
        )
    if not isinstance(comparisons_raw, dict) or not comparisons_raw:
        raise ConfigError(
            "comparisons は「名前: 拠点の一覧」の形式で1件以上指定してください"
        )

    comparisons: List[Comparison] = []
    for name, items in comparisons_raw.items():
        label = str(name)
        if not label.strip():
            raise ConfigError("comparisons の名前が空です")
        bad = sorted(_INVALID_NAME_CHARS & set(label))
        if bad:
            raise ConfigError(
                f"比較の名前に使えない文字が含まれています: {label!r} ({''.join(bad)})\n"
                f"  → レポートのファイル名に使うため、"
                f"{''.join(sorted(_INVALID_NAME_CHARS))} は使えません"
            )
        comparisons.append(Comparison(
            name=label,
            locations=_parse_locations(items, base_dir, f"comparisons.{label}"),
        ))
    return comparisons


def _select_comparison(comparisons: List[Comparison], requested: str | None) -> str:
    """実行対象の比較を決める。見つからなければ候補を添えてエラーにする。"""
    names = [c.name for c in comparisons if c.name]
    if requested:
        if requested not in names:
            raise ConfigError(
                f"比較 '{requested}' は設定にありません。\n"
                f"  → 定義されているのは: {', '.join(names) if names else '(名前なし)'}"
            )
        return requested
    if len(comparisons) == 1:
        return comparisons[0].name
    raise ConfigError(
        "比較対象を指定してください (--comparison NAME)。\n"
        f"  → 定義されているのは: {', '.join(names)}"
    )


def load_config(
    path: str | Path,
    comparison: str | None = None,
    *,
    require_selection: bool = True,
) -> Config:
    config_path = Path(path).resolve()
    if not config_path.is_file():
        raise ConfigError(f"設定ファイルが見つかりません: {config_path}")

    base_dir = config_path.parent

    with config_path.open("r", encoding="utf-8") as f:
        raw = yaml.safe_load(f) or {}

    comparisons = _parse_comparisons(raw, base_dir)
    # 一覧表示のときは選択を求めない (名前を知るために名前が要る状態を避ける)
    selected = (
        _select_comparison(comparisons, comparison) if require_selection
        else (comparison or "")
    )

    exclude_patterns = raw.get("exclude_patterns") or []
    if not isinstance(exclude_patterns, list):
        raise ConfigError("exclude_patterns はリスト形式で指定してください")
    exclude_patterns = [str(p) for p in exclude_patterns]

    output_raw = raw.get("output") or {}
    fmt = str(output_raw.get("format", "excel")).lower()
    if fmt not in SUPPORTED_FORMATS:
        raise ConfigError(
            f"output.format は {sorted(SUPPORTED_FORMATS)} のいずれかを指定してください (現在: {fmt})"
        )
    output_dir = _resolve(Path(str(output_raw.get("output_dir", "./reports"))), base_dir)
    raw_limit = _positive_int(
        output_raw.get("max_table_rows", DEFAULT_MAX_TABLE_ROWS), "output.max_table_rows"
    )
    keep_reports = _positive_int(
        output_raw.get("keep_reports", 0), "output.keep_reports"
    )
    output = OutputConfig(
        format=fmt,
        output_dir=output_dir,
        max_table_rows=raw_limit,
        keep_reports=keep_reports,
    )

    perf_raw = raw.get("performance") or {}
    workers = int(perf_raw.get("parallel_workers", 4))
    if workers < 1:
        raise ConfigError("performance.parallel_workers は1以上を指定してください")
    algo = str(perf_raw.get("hash_algorithm", "sha256")).lower()
    if algo not in SUPPORTED_HASH_ALGOS:
        raise ConfigError(
            f"performance.hash_algorithm は {sorted(SUPPORTED_HASH_ALGOS)} のみ対応 (現在: {algo})"
        )
    hash_mode = str(perf_raw.get("hash_mode", HASH_MODE_ALWAYS)).lower()
    if hash_mode not in HASH_MODES:
        raise ConfigError(
            f"performance.hash_mode は {sorted(HASH_MODES)} のいずれかを指定してください "
            f"(現在: {hash_mode})"
        )
    try:
        tolerance = float(perf_raw.get("mtime_tolerance_sec", DEFAULT_MTIME_TOLERANCE_SEC))
    except (TypeError, ValueError):
        raise ConfigError("performance.mtime_tolerance_sec は数値で指定してください")
    if tolerance < 0:
        raise ConfigError("performance.mtime_tolerance_sec は0以上を指定してください")
    performance = PerformanceConfig(
        parallel_workers=workers,
        hash_algorithm=algo,
        hash_mode=hash_mode,
        mtime_tolerance_sec=tolerance,
    )

    matching_raw = raw.get("matching") or {}
    for bool_key in ("normalize_unicode", "case_sensitive"):
        if bool_key in matching_raw and not isinstance(matching_raw[bool_key], bool):
            raise ConfigError(f"matching.{bool_key} は true / false で指定してください")
    matching = MatchingConfig(
        normalize_unicode=bool(matching_raw.get("normalize_unicode", True)),
        case_sensitive=bool(matching_raw.get("case_sensitive", True)),
    )

    return Config(
        comparisons=comparisons,
        selected=selected,
        exclude_patterns=exclude_patterns,
        matching=matching,
        output=output,
        performance=performance,
    )
