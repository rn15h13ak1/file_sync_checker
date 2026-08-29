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
class Config:
    locations: List[Location]
    exclude_patterns: List[str] = field(default_factory=list)
    matching: MatchingConfig = field(default_factory=MatchingConfig)
    output: OutputConfig = field(default_factory=lambda: OutputConfig("excel", Path("./reports")))
    performance: PerformanceConfig = field(
        default_factory=lambda: PerformanceConfig(4, "sha256")
    )


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


def load_config(path: str | Path) -> Config:
    config_path = Path(path).resolve()
    if not config_path.is_file():
        raise ConfigError(f"設定ファイルが見つかりません: {config_path}")

    base_dir = config_path.parent

    with config_path.open("r", encoding="utf-8") as f:
        raw = yaml.safe_load(f) or {}

    locations_raw = raw.get("locations")
    if not locations_raw or not isinstance(locations_raw, list):
        raise ConfigError("locations は1件以上指定してください")

    locations: List[Location] = []
    seen_names: set[str] = set()
    for i, item in enumerate(locations_raw):
        if not isinstance(item, dict):
            raise ConfigError(f"locations[{i}] が不正です")
        name = item.get("name")
        path = item.get("path")
        if not name or not path:
            raise ConfigError(f"locations[{i}] は name と path が必須です")
        if name in seen_names:
            raise ConfigError(f"拠点名が重複しています: {name}")
        # バックスラッシュ表記は YAML パース時に潰れて壊れることが多く、
        # またコピー機能でのパス組み立てを単純化するためフォワードスラッシュに統一する。
        path_str = str(path)
        if "\\" in path_str:
            raise ConfigError(
                f"locations[{i}].path にバックスラッシュが含まれています: {path_str!r}\n"
                f"  → フォワードスラッシュで記述してください "
                f"(例: '//server-a/share/docs')"
            )
        seen_names.add(name)
        locations.append(Location(name=str(name), path=_resolve(Path(path_str), base_dir)))

    if len(locations) < 2:
        raise ConfigError("locations は2件以上必要です")

    # パス重複検出: 同じ実体パスを複数拠点として登録すると常に一致してしまうため弾く。
    # os.path.normcase は Windows では大小無視・スラッシュ統一、UNIX では no-op。
    seen_paths: Dict[str, str] = {}
    for loc in locations:
        key = os.path.normcase(os.path.normpath(str(loc.path)))
        if key in seen_paths:
            raise ConfigError(
                f"拠点 '{loc.name}' のパスが拠点 '{seen_paths[key]}' と重複しています: {loc.path}"
            )
        seen_paths[key] = loc.name

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
        locations=locations,
        exclude_patterns=exclude_patterns,
        matching=matching,
        output=output,
        performance=performance,
    )
