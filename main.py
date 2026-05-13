"""ファイル同期チェックツール - エントリポイント。"""
from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

from config import Config, ConfigError, load_config
from comparator import compare
from reporter import ReportContext, write_excel, write_html
from scanner import scan_location
from utils import ensure_dir, human_bytes, setup_logging, timestamp_slug


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="3拠点（以上）のファイル同期状況をチェックして Excel/HTML レポートを生成します。"
    )
    parser.add_argument(
        "-c", "--config", default="config.yaml",
        help="設定ファイルのパス (default: config.yaml)",
    )
    parser.add_argument("-v", "--verbose", action="store_true", help="詳細ログを出力")
    parser.add_argument("--no-progress", action="store_true", help="進捗バーを無効化")
    return parser.parse_args()


def run(config: Config, config_path: Path, *, show_progress: bool, log) -> int:
    started_at = datetime.now()
    log.info("スキャン開始: %d 拠点", len(config.locations))

    scans = []
    for loc in config.locations:
        log.info("  [%s] %s", loc.name, loc.path)
        scan = scan_location(
            name=loc.name,
            root=loc.path,
            exclude_patterns=config.exclude_patterns,
            parallel_workers=config.performance.parallel_workers,
            hash_algorithm=config.performance.hash_algorithm,
            show_progress=show_progress,
        )
        size = sum(f.size for f in scan.files.values())
        log.info(
            "  [%s] ファイル数=%d, サイズ=%s, エラー=%d",
            loc.name, len(scan.files), human_bytes(size), len(scan.errors),
        )
        scans.append(scan)

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
    print(f"  フォルダ構造差分       : {len(comparison.dir_diffs)}")
    print("-" * 60)
    print(f"  所要時間               : {elapsed:.2f} 秒")
    for p in written:
        print(f"  出力                   : {p}")
    print("=" * 60)

    # 差分・エラーがあれば exit 1
    total_diffs = (
        len(comparison.hash_mismatches)
        + len(comparison.missing_files)
        + len(comparison.extra_files)
        + len(comparison.dir_diffs)
    )
    total_errors = sum(len(s.errors) for s in scans)
    if total_diffs > 0 or total_errors > 0:
        return 1
    return 0


def main() -> int:
    args = parse_args()
    log = setup_logging(verbose=args.verbose)
    config_path = Path(args.config)
    try:
        config = load_config(config_path)
    except ConfigError as e:
        log.error("設定エラー: %s", e)
        return 2
    except Exception as e:  # YAML パースエラー等
        log.error("設定ファイル読み込み失敗: %s", e)
        return 2

    try:
        return run(
            config,
            config_path,
            show_progress=not args.no_progress,
            log=log,
        )
    except KeyboardInterrupt:
        log.warning("中断されました")
        return 130


if __name__ == "__main__":
    sys.exit(main())
