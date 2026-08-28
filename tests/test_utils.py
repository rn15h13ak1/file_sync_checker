"""utils: 表示整形とロギング初期化。"""
from __future__ import annotations

import logging
from datetime import datetime
from pathlib import Path

import pytest

from utils import ensure_dir, human_bytes, setup_logging, timestamp_slug


class TestHumanBytes:
    @pytest.mark.parametrize(
        "n, expected",
        [
            (0, "0 B"),
            (1023, "1023 B"),          # KB に切り替わる直前
            (1024, "1.00 KB"),
            (1024 ** 2 - 1, "1024.00 KB"),
            (1024 ** 2, "1.00 MB"),
            (1024 ** 3, "1.00 GB"),
            (1024 ** 4, "1.00 TB"),
            (1024 ** 5, "1.00 PB"),    # 単位表の外 (最後のフォールバック)
            (1024 ** 6, "1024.00 PB"),
        ],
    )
    def test_unit_boundaries(self, n: int, expected: str):
        assert human_bytes(n) == expected


class TestTimestampSlug:
    def test_formats_for_filenames(self):
        assert timestamp_slug(datetime(2026, 5, 20, 9, 7, 3)) == "20260520-090703"

    def test_defaults_to_now(self):
        slug = timestamp_slug()
        assert len(slug) == 15 and slug[8] == "-"


class TestEnsureDir:
    def test_creates_nested_directories(self, tmp_path: Path):
        target = tmp_path / "a" / "b" / "c"
        assert ensure_dir(target) == target
        assert target.is_dir()

    def test_is_idempotent(self, tmp_path: Path):
        target = tmp_path / "reports"
        ensure_dir(target)
        assert ensure_dir(target).is_dir()  # 既存でも失敗しない


class TestSetupLogging:
    def test_returns_same_logger_without_duplicating_handlers(self):
        """複数回呼んでもハンドラが増えない (ログが二重に出ない)。"""
        first = setup_logging()
        count = len(first.handlers)
        second = setup_logging(verbose=True)
        assert second is first
        assert len(second.handlers) == count

    def test_verbose_sets_debug_level(self):
        # 既にハンドラが付いている場合は早期 return するため、専用のロガーで確認する
        logger = logging.getLogger("file_sync_checker")
        saved_handlers, saved_level = logger.handlers[:], logger.level
        logger.handlers.clear()
        try:
            assert setup_logging(verbose=True).level == logging.DEBUG
            logger.handlers.clear()
            assert setup_logging(verbose=False).level == logging.INFO
        finally:
            logger.handlers[:] = saved_handlers
            logger.setLevel(saved_level)
