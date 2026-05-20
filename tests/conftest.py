"""共通フィクスチャ / ヘルパー。"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Dict, Optional

import pytest

from comparator import FileRow
from scanner import FileEntry, ScanResult


def make_entry(hash_: str = "h", size: int = 1, mtime: Optional[datetime] = None) -> FileEntry:
    return FileEntry(size=size, mtime=mtime or datetime(2026, 1, 1, 0, 0, 0), hash=hash_)


def make_scan(
    name: str,
    files: Optional[Dict[str, FileEntry]] = None,
    dirs: Optional[Dict] = None,
    file_errors: Optional[Dict[str, str]] = None,
    errors: Optional[list] = None,
) -> ScanResult:
    return ScanResult(
        location_name=name,
        root=Path("/tmp/dummy"),
        files=files or {},
        dirs=dirs or {},
        errors=errors or [],
        file_errors=file_errors or {},
    )


@pytest.fixture
def tmp_tree(tmp_path: Path):
    """3拠点フォルダのスケルトンを返すヘルパー。"""
    locs = {n: tmp_path / n for n in ("A", "B", "C")}
    for p in locs.values():
        p.mkdir()
    return locs
