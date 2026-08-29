"""ファイル同期チェックツール - エントリポイント。"""
from __future__ import annotations

import argparse
import os
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path

from config import (
    SUPPORTED_FORMATS,
    Config,
    ConfigError,
    apply_overrides,
    load_config,
)
from comparator import compare
from reporter import ReportContext, ReportSettings, write_excel, write_html
from scanner import HASH_MODES, ScanCancelled, scan_locations
from utils import ensure_dir, human_bytes, setup_logging, timestamp_slug


SCRIPT_DIR = Path(__file__).resolve().parent

# 終了コード
EXIT_OK = 0
EXIT_DIFF = 1           # 差分あり (スキャン自体は完走)
EXIT_CONFIG_ERROR = 2
EXIT_SCAN_ERROR = 3     # 読み取り失敗あり = スキャンが不完全
EXIT_UNEXPECTED = 4     # 予期しない例外 (レポート生成失敗など)
EXIT_INTERRUPTED = 130


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="3拠点（以上）のファイル同期状況をチェックして Excel/HTML レポートを生成します。"
    )
    parser.add_argument(
        "-c", "--config", default=None,
        help="設定ファイルのパス (default: ./config.yaml またはスクリプトと同じディレクトリの config.yaml)",
    )
    parser.add_argument("-v", "--verbose", action="store_true", help="詳細ログを出力")
    parser.add_argument(
        "--no-progress", action="store_true",
        help="進捗バーを無効化 (端末以外に出力している場合は指定しなくても無効)",
    )
    # 以下は設定ファイルの値を上書きする (CI から config.yaml を書き換えずに使うため)
    parser.add_argument(
        "--format", dest="output_format", choices=sorted(SUPPORTED_FORMATS), default=None,
        help="出力形式を上書き (設定: output.format)",
    )
    parser.add_argument(
        "-o", "--output-dir", default=None,
        help="レポート出力先を上書き (設定: output.output_dir。CWD からの相対)",
    )
    parser.add_argument(
        "--hash-mode", choices=sorted(HASH_MODES), default=None,
        help="ハッシュ計算範囲を上書き (設定: performance.hash_mode)",
    )
    parser.add_argument(
        "--retry", type=int, default=0, metavar="N",
        help="読み取りに失敗したファイルを N 回まで再試行する "
             "(既定: 0 = 再試行しない)。ネットワークの瞬断や一時的なロック向け",
    )
    args = parser.parse_args()
    if args.retry < 0:
        parser.error("--retry は0以上を指定してください")
    return args


def _resolve_default_config() -> Path:
    """`-c` 省略時の config.yaml を CWD → スクリプト同梱の順で探す。"""
    cwd_candidate = Path.cwd() / "config.yaml"
    if cwd_candidate.is_file():
        return cwd_candidate
    script_candidate = SCRIPT_DIR / "config.yaml"
    if script_candidate.is_file():
        return script_candidate
    # どちらも無い場合は CWD を返す (load_config 側でエラーメッセージにする)
    return cwd_candidate


def _write_latest_alias(src: Path, dst: Path) -> Path:
    """最新レポートの安定名を差し替える。

    直接上書きすると、書き込み中に開かれたレポートが途中までの内容になる。
    同じディレクトリに一時ファイルを作ってから `os.replace` で差し替えることで、
    閲覧側からは常に完全な旧版か新版のどちらかに見える (差し替えはアトミック)。
    """
    tmp = dst.with_name(dst.name + ".tmp")
    try:
        shutil.copy2(src, tmp)
        os.replace(tmp, dst)
    except BaseException:
        tmp.unlink(missing_ok=True)
        raise
    return dst


# このツールが出力したタイムスタンプ付きレポートだけを表す。
# 削除対象を誤らないよう、桁数まで含めて厳密に一致させる。
# 固定名の sync-check.html (最新への安定リンク) はこれに一致しないので消えない。
_REPORT_NAME_RE = re.compile(r"^sync-check-(\d{8}-\d{6})\.(?:html|xlsx)$")


def _prune_old_reports(out_dir: Path, keep: int, log) -> None:
    """古いレポートを削除し、直近 `keep` 回分だけ残す。

    定期実行では出力ディレクトリにレポートが際限なく溜まる
    (大きい共有だと HTML 22MB + Excel 3MB で年間 9GB 規模)。

    削除は元に戻せないため、次の条件をすべて満たすファイルだけを対象にする:
      - このツールの命名規則 `sync-check-YYYYMMDD-HHMMSS.{html,xlsx}` に完全一致
      - 出力ディレクトリ直下の通常ファイル
    「実行回数」で数えるので、1 回の実行が html と xlsx を出していれば
    それらは 1 回分として扱う。削除に失敗しても実行自体は成功させる
    (レポートは既に書けており、後片付けの失敗で終了コードを変えない)。
    """
    if keep <= 0:
        return

    runs: dict = {}
    for path in out_dir.iterdir():
        if not path.is_file():
            continue
        m = _REPORT_NAME_RE.match(path.name)
        if m:
            runs.setdefault(m.group(1), []).append(path)

    # タイムスタンプの新しい順に keep 回分を残す (ファイル名がそのまま時系列)
    for slug in sorted(runs, reverse=True)[keep:]:
        for path in runs[slug]:
            try:
                path.unlink()
                log.info("  古いレポートを削除: %s", path.name)
            except OSError as e:
                log.warning("  レポートの削除に失敗: %s (%s)", path.name, e)


def run(
    config: Config, config_path: Path, *, show_progress: bool, log, retry: int = 0
) -> int:
    started_at = datetime.now()
    log.info("スキャン開始: %d 拠点", len(config.locations))

    for loc in config.locations:
        log.info("  [%s] %s", loc.name, loc.path)

    def _stat_done(sr) -> None:
        size = sum(f.size for f in sr.stats.values())
        log.info(
            "  [%s] 列挙完了: ファイル数=%d, サイズ=%s, エラー=%d",
            sr.location_name, len(sr.stats), human_bytes(size), len(sr.errors),
        )

    scans = scan_locations(
        [(loc.name, loc.path) for loc in config.locations],
        exclude_patterns=config.exclude_patterns,
        parallel_workers=config.performance.parallel_workers,
        hash_algorithm=config.performance.hash_algorithm,
        hash_mode=config.performance.hash_mode,
        mtime_tolerance_sec=config.performance.mtime_tolerance_sec,
        normalize_unicode=config.matching.normalize_unicode,
        case_sensitive=config.matching.case_sensitive,
        show_progress=show_progress,
        retry=retry,
        on_stat_done=_stat_done,
    )
    for s in scans:
        if s.skipped_hashes:
            log.info(
                "  [%s] ハッシュ省略: %d 件 (hash_mode=%s)",
                s.location_name, s.skipped_hashes, config.performance.hash_mode,
            )

    log.info("差分検出中...")
    comparison = compare(scans)
    finished_at = datetime.now()

    out_dir = ensure_dir(config.output.output_dir)
    slug = timestamp_slug(started_at)
    ctx = ReportContext(
        started_at=started_at,
        finished_at=finished_at,
        config_path=config_path,
        scans=scans,
        comparison=comparison,
        # 設定ファイルのパスだけでは、そのレポートがどの保証を意味するのか
        # 読み手に分からない。実際に使われた条件をレポートに残す。
        settings=ReportSettings(
            hash_mode=config.performance.hash_mode,
            hash_algorithm=config.performance.hash_algorithm,
            mtime_tolerance_sec=config.performance.mtime_tolerance_sec,
            exclude_patterns=config.exclude_patterns,
            normalize_unicode=config.matching.normalize_unicode,
            case_sensitive=config.matching.case_sensitive,
            retry=retry,
            max_table_rows=config.output.max_table_rows,
        ),
    )

    written = []
    fmt = config.output.format
    if fmt in ("excel", "both"):
        out_xlsx = out_dir / f"sync-check-{slug}.xlsx"
        write_excel(ctx, out_xlsx)
        written.append(out_xlsx)
    if fmt in ("html", "both"):
        out_html = out_dir / f"sync-check-{slug}.html"
        write_html(ctx, out_html)
        written.append(out_html)
        # 安定リンク用: タイムスタンプ無しの最新レポートを上書きで生成する
        # (ブックマークや自動化スクリプトから常に最新を参照できるようにするため)
        written.append(_write_latest_alias(out_html, out_dir / "sync-check.html"))

    # 新しいレポートを書いた後に、古い分を片付ける
    _prune_old_reports(out_dir, config.output.keep_reports, log)

    # コンソールサマリー
    elapsed = (finished_at - started_at).total_seconds()
    print()
    print("=" * 60)
    print(" スキャン結果サマリー")
    print("=" * 60)
    for s in scans:
        size = sum(f.size for f in s.files.values())
        print(f"  {s.location_name:<10} files={len(s.files):>6}  size={human_bytes(size):>10}"
              f"  errors={len(s.errors)}")
    print("-" * 60)
    print(f"  ハッシュ不一致         : {len(comparison.hash_mismatches)}")
    print(f"  ファイル欠落           : {len(comparison.missing_files)}")
    print(f"  余分なファイル         : {len(comparison.extra_files)}")
    print(f"  エラー対象ファイル     : {len(comparison.errored_files)}")
    print(f"  フォルダ構造差分       : {len(comparison.dir_diffs)}")
    print("-" * 60)
    print(f"  所要時間               : {elapsed:.2f} 秒")
    for p in written:
        print(f"  出力                   : {p}")
    print("=" * 60)

    total_diffs = (
        len(comparison.hash_mismatches)
        + len(comparison.missing_files)
        + len(comparison.extra_files)
        + len(comparison.errored_files)
        + len(comparison.dir_diffs)
    )
    total_errors = sum(len(s.errors) for s in scans)
    # 読み取りエラーは「差分あり」と区別する。エラーがあるとスキャン自体が不完全で、
    # 差分が 0 件でも「一致していた」とは言えないため、自動化側で別扱いできるようにする。
    if total_errors > 0:
        return EXIT_SCAN_ERROR
    if total_diffs > 0:
        return EXIT_DIFF
    return EXIT_OK


def main() -> int:
    args = parse_args()
    log = setup_logging(verbose=args.verbose)
    config_path = Path(args.config) if args.config else _resolve_default_config()
    try:
        config = apply_overrides(
            load_config(config_path),
            output_format=args.output_format,
            output_dir=args.output_dir,
            hash_mode=args.hash_mode,
        )
    except ConfigError as e:
        log.error("設定エラー: %s", e)
        return EXIT_CONFIG_ERROR
    except Exception as e:  # YAML パースエラー等
        log.error("設定ファイル読み込み失敗: %s", e)
        return EXIT_CONFIG_ERROR

    # 進捗バーは端末に出しているときだけ表示する。
    # cron などで stderr をファイルに落としていると、tqdm の更新が
    # そのままログに書き込まれる (実測: 10 秒のスキャンで stderr の 97% が
    # 進捗バー由来の制御文字。10 分なら 1 実行あたり 400KB 程度)。
    show_progress = not args.no_progress and sys.stderr.isatty()

    try:
        return run(
            config,
            config_path,
            show_progress=show_progress,
            log=log,
            retry=args.retry,
        )
    except (KeyboardInterrupt, ScanCancelled):
        log.warning("中断されました")
        return EXIT_INTERRUPTED
    except Exception:
        # 素通りさせるとトレースバックのまま終了コード 1 になり、
        # 「差分あり」と区別できない。原因は追えるようログには残す。
        log.exception("予期しないエラーで中断しました")
        return EXIT_UNEXPECTED


if __name__ == "__main__":
    sys.exit(main())
