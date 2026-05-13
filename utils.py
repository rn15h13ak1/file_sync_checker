"""ロギングと共通ユーティリティ。"""
from __future__ import annotations

import logging
import sys
from datetime import datetime
from pathlib import Path


def setup_logging(verbose: bool = False) -> logging.Logger:
    logger = logging.getLogger("file_sync_checker")
    if logger.handlers:
        return logger
    handler = logging.StreamHandler(sys.stderr)
    handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S"))
    logger.addHandler(handler)
    logger.setLevel(logging.DEBUG if verbose else logging.INFO)
    return logger


def timestamp_slug(dt: datetime | None = None) -> str:
    """レポートファイル名向けのタイムスタンプ (YYYYMMDD-HHMMSS)。"""
    return (dt or datetime.now()).strftime("%Y%m%d-%H%M%S")


def human_bytes(n: int) -> str:
    if n < 1024:
        return f"{n} B"
    units = ["KB", "MB", "GB", "TB"]
    size = float(n)
    for u in units:
        size /= 1024
        if size < 1024:
            return f"{size:.2f} {u}"
    return f"{size:.2f} PB"


def ensure_dir(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path
