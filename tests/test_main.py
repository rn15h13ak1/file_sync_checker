"""main.run の出力ファイル名整合性テスト。"""
from __future__ import annotations

import logging
from pathlib import Path

import yaml

from config import load_config
from main import run


def _make_locations(tmp_path: Path) -> tuple[Path, Path]:
    a = tmp_path / "locA"
    b = tmp_path / "locB"
    a.mkdir()
    b.mkdir()
    (a / "shared.txt").write_text("data", encoding="utf-8")
    (b / "shared.txt").write_text("data", encoding="utf-8")
    return a, b


def _write_config(tmp_path: Path, fmt: str) -> Path:
    a, b = _make_locations(tmp_path)
    cfg = tmp_path / "config.yaml"
    cfg.write_text(yaml.safe_dump({
        "locations": [
            {"name": "A", "path": str(a)},
            {"name": "B", "path": str(b)},
        ],
        "output": {"format": fmt, "output_dir": str(tmp_path / "reports")},
        "performance": {"parallel_workers": 1, "hash_algorithm": "sha256"},
    }, allow_unicode=True))
    return cfg


def _silent_logger() -> logging.Logger:
    log = logging.getLogger("test_main")
    if not log.handlers:
        log.addHandler(logging.NullHandler())
    log.setLevel(logging.CRITICAL)
    return log


class TestHtmlAliasOutput:
    def test_html_format_emits_both_timestamped_and_latest(self, tmp_path, capsys):
        cfg_path = _write_config(tmp_path, "html")
        config = load_config(cfg_path)
        run(config, cfg_path, show_progress=False, log=_silent_logger())
        capsys.readouterr()  # コンソール出力を捨てる

        out_dir = tmp_path / "reports"
        names = sorted(f.name for f in out_dir.iterdir())
        # 安定名 (タイムスタンプ無し) と タイムスタンプ付き 両方存在
        assert "sync-check.html" in names
        assert any(
            n.startswith("sync-check-") and n.endswith(".html") and n != "sync-check.html"
            for n in names
        )

    def test_latest_html_has_same_content_as_timestamped(self, tmp_path, capsys):
        cfg_path = _write_config(tmp_path, "html")
        config = load_config(cfg_path)
        run(config, cfg_path, show_progress=False, log=_silent_logger())
        capsys.readouterr()

        out_dir = tmp_path / "reports"
        latest = (out_dir / "sync-check.html").read_bytes()
        timestamped = next(
            f for f in out_dir.iterdir()
            if f.name.startswith("sync-check-") and f.name != "sync-check.html"
        ).read_bytes()
        assert latest == timestamped

    def test_excel_only_format_does_not_create_latest_html(self, tmp_path, capsys):
        cfg_path = _write_config(tmp_path, "excel")
        config = load_config(cfg_path)
        run(config, cfg_path, show_progress=False, log=_silent_logger())
        capsys.readouterr()

        out_dir = tmp_path / "reports"
        names = [f.name for f in out_dir.iterdir()]
        assert "sync-check.html" not in names
        # 安定名 Excel は意図的に作らない (HTML だけが alias 対象)
        assert "sync-check.xlsx" not in names
        # タイムスタンプ付き xlsx は存在
        assert any(n.startswith("sync-check-") and n.endswith(".xlsx") for n in names)

    def test_both_format_creates_html_alias_only(self, tmp_path, capsys):
        cfg_path = _write_config(tmp_path, "both")
        config = load_config(cfg_path)
        run(config, cfg_path, show_progress=False, log=_silent_logger())
        capsys.readouterr()

        out_dir = tmp_path / "reports"
        names = [f.name for f in out_dir.iterdir()]
        assert "sync-check.html" in names      # HTML alias は出る
        assert "sync-check.xlsx" not in names  # Excel alias は出ない

    def test_re_running_overwrites_latest_html(self, tmp_path, capsys):
        """2回実行すると sync-check.html は最新の内容で上書きされる。"""
        cfg_path = _write_config(tmp_path, "html")
        config = load_config(cfg_path)
        out_dir = tmp_path / "reports"

        run(config, cfg_path, show_progress=False, log=_silent_logger())
        capsys.readouterr()
        first_content = (out_dir / "sync-check.html").read_bytes()
        first_mtime = (out_dir / "sync-check.html").stat().st_mtime

        # 内容を変えて再実行
        (tmp_path / "locA" / "new.txt").write_text("only in A", encoding="utf-8")

        run(config, cfg_path, show_progress=False, log=_silent_logger())
        capsys.readouterr()
        second_content = (out_dir / "sync-check.html").read_bytes()
        second_mtime = (out_dir / "sync-check.html").stat().st_mtime

        assert second_content != first_content
        assert second_mtime >= first_mtime
