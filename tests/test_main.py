"""main.run の出力ファイル名整合性テスト。"""
from __future__ import annotations

import logging
import os
import sys
from pathlib import Path

import pytest

from .conftest import skip_or_fail
import yaml

from config import apply_overrides, load_config
from main import (
    EXIT_CONFIG_ERROR,
    EXIT_DIFF,
    EXIT_INTERRUPTED,
    EXIT_OK,
    EXIT_SCAN_ERROR,
    EXIT_UNEXPECTED,
    run,
)


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


class TestExitCodes:
    def test_all_match_returns_zero(self, tmp_path, capsys):
        cfg_path = _write_config(tmp_path, "html")
        rc = run(load_config(cfg_path), cfg_path, show_progress=False, log=_silent_logger())
        capsys.readouterr()
        assert rc == EXIT_OK

    def test_difference_returns_one(self, tmp_path, capsys):
        cfg_path = _write_config(tmp_path, "html")
        (tmp_path / "locA" / "extra.txt").write_text("only in A", encoding="utf-8")
        rc = run(load_config(cfg_path), cfg_path, show_progress=False, log=_silent_logger())
        capsys.readouterr()
        assert rc == EXIT_DIFF

    def test_read_error_returns_three(self, tmp_path, capsys):
        """読み取りエラーは差分と区別する。スキャンが不完全なため。"""
        if sys.platform == "win32":
            skip_or_fail("chmod-based unreadable test is POSIX only")
        if os.geteuid() == 0:
            skip_or_fail("running as root bypasses permission denial")
        cfg_path = _write_config(tmp_path, "html")
        locked = tmp_path / "locA" / "locked.txt"
        locked.write_text("secret", encoding="utf-8")
        (tmp_path / "locB" / "locked.txt").write_text("secret", encoding="utf-8")
        locked.chmod(0o000)
        try:
            rc = run(load_config(cfg_path), cfg_path, show_progress=False,
                     log=_silent_logger())
            capsys.readouterr()
            assert rc == EXIT_SCAN_ERROR
        finally:
            locked.chmod(0o600)


class TestUnexpectedError:
    def test_unexpected_exception_gets_its_own_exit_code(
        self, tmp_path, monkeypatch, capsys
    ):
        """予期しない例外を「差分あり」と同じ 1 で返さない。

        素通りさせるとトレースバックのまま終了コード 1 になり、
        CI からは差分が出ただけに見えてしまう。
        """
        import main as main_mod

        cfg_path = _write_config(tmp_path, "html")
        monkeypatch.setattr(
            sys, "argv", ["main.py", "-c", str(cfg_path), "--no-progress"]
        )

        def boom(*a, **k):
            raise RuntimeError("レポート生成に失敗")

        monkeypatch.setattr(main_mod, "write_html", boom)
        assert main_mod.main() == EXIT_UNEXPECTED
        capsys.readouterr()

    def test_normal_run_is_unaffected(self, tmp_path, monkeypatch, capsys):
        import main as main_mod

        cfg_path = _write_config(tmp_path, "html")
        monkeypatch.setattr(
            sys, "argv", ["main.py", "-c", str(cfg_path), "--no-progress"]
        )
        assert main_mod.main() == EXIT_OK
        capsys.readouterr()


class TestMainErrorPaths:
    """main() の異常系。CLI テストは別プロセスで動きカバレッジに乗らないため、
    ここでは main() を直接呼んで確認する。"""

    def _argv(self, monkeypatch, *args: str) -> None:
        monkeypatch.setattr(sys, "argv", ["main.py", *args])

    def test_missing_config_returns_config_error(self, tmp_path, monkeypatch, capsys):
        import main as main_mod

        self._argv(monkeypatch, "-c", str(tmp_path / "nope.yaml"), "--no-progress")
        assert main_mod.main() == EXIT_CONFIG_ERROR
        capsys.readouterr()

    def test_broken_yaml_returns_config_error(self, tmp_path, monkeypatch, capsys):
        """ConfigError 以外 (YAML パースエラー) も設定エラーとして扱う。"""
        import main as main_mod

        bad = tmp_path / "bad.yaml"
        bad.write_text("locations: [unclosed\n", encoding="utf-8")
        self._argv(monkeypatch, "-c", str(bad), "--no-progress")
        assert main_mod.main() == EXIT_CONFIG_ERROR
        capsys.readouterr()

    def test_keyboard_interrupt_returns_130(self, tmp_path, monkeypatch, capsys):
        import main as main_mod

        cfg_path = _write_config(tmp_path, "html")
        self._argv(monkeypatch, "-c", str(cfg_path), "--no-progress")

        def interrupted(*a, **k):
            raise KeyboardInterrupt

        monkeypatch.setattr(main_mod, "scan_locations", interrupted)
        assert main_mod.main() == EXIT_INTERRUPTED
        capsys.readouterr()

    def test_scan_cancelled_returns_130(self, tmp_path, monkeypatch, capsys):
        import main as main_mod
        from scanner import ScanCancelled

        cfg_path = _write_config(tmp_path, "html")
        self._argv(monkeypatch, "-c", str(cfg_path), "--no-progress")

        def cancelled(*a, **k):
            raise ScanCancelled()

        monkeypatch.setattr(main_mod, "scan_locations", cancelled)
        assert main_mod.main() == EXIT_INTERRUPTED
        capsys.readouterr()

    def test_default_config_lookup_prefers_cwd(self, tmp_path, monkeypatch, capsys):
        """-c 省略時は CWD の config.yaml を使う。"""
        import main as main_mod

        _write_config(tmp_path, "html")  # tmp_path/config.yaml を作る
        monkeypatch.chdir(tmp_path)
        self._argv(monkeypatch, "--no-progress")
        assert main_mod.main() == EXIT_OK
        capsys.readouterr()
        assert (tmp_path / "reports" / "sync-check.html").is_file()

    def test_default_config_falls_back_to_script_dir(self, tmp_path, monkeypatch, capsys):
        """CWD に無ければスクリプト同梱の config.yaml を探す。"""
        import main as main_mod

        empty = tmp_path / "empty"
        empty.mkdir()
        monkeypatch.chdir(empty)
        monkeypatch.setattr(main_mod, "SCRIPT_DIR", tmp_path)
        _write_config(tmp_path, "html")
        self._argv(monkeypatch, "--no-progress")
        assert main_mod.main() == EXIT_OK
        capsys.readouterr()

    def test_default_config_lookup_falls_back_to_cwd_when_absent(
        self, tmp_path, monkeypatch, capsys
    ):
        """どちらにも無ければ CWD のパスを返し、設定エラーとして報告する。"""
        import main as main_mod

        empty = tmp_path / "empty"
        empty.mkdir()
        monkeypatch.chdir(empty)
        monkeypatch.setattr(main_mod, "SCRIPT_DIR", tmp_path / "also_empty")
        self._argv(monkeypatch, "--no-progress")
        assert main_mod.main() == EXIT_CONFIG_ERROR
        capsys.readouterr()

    def test_skipped_hash_count_is_logged(self, tmp_path, caplog):
        """hash_mode=smart でハッシュを省略した件数をログに出す。"""
        cfg_path = _write_config(tmp_path, "html")
        config = apply_overrides(load_config(cfg_path), hash_mode="smart")
        log = logging.getLogger("test_main_skipped")
        with caplog.at_level(logging.INFO, logger="test_main_skipped"):
            run(config, cfg_path, show_progress=False, log=log)
        assert any("ハッシュ省略" in r.message for r in caplog.records), caplog.text


class TestArgParsing:
    def test_negative_retry_is_rejected(self, monkeypatch, capsys):
        """--retry に負値を渡したら起動前に弾く。"""
        import main as main_mod

        monkeypatch.setattr(sys, "argv", ["main.py", "--retry", "-1"])
        with pytest.raises(SystemExit) as exc:
            main_mod.parse_args()
        assert exc.value.code == 2
        assert "--retry" in capsys.readouterr().err

    def test_retry_defaults_to_zero(self, monkeypatch):
        import main as main_mod

        monkeypatch.setattr(sys, "argv", ["main.py"])
        assert main_mod.parse_args().retry == 0


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

    def test_latest_alias_is_replaced_atomically(self, tmp_path, capsys):
        """差し替えは os.replace で行い、一時ファイルを残さない。

        直接上書きすると、閲覧中のブラウザが途中までの HTML を読む可能性がある。
        """
        cfg_path = _write_config(tmp_path, "html")
        config = load_config(cfg_path)
        out_dir = tmp_path / "reports"

        run(config, cfg_path, show_progress=False, log=_silent_logger())
        capsys.readouterr()
        run(config, cfg_path, show_progress=False, log=_silent_logger())
        capsys.readouterr()

        names = [f.name for f in out_dir.iterdir()]
        assert "sync-check.html" in names
        assert not any(n.endswith(".tmp") for n in names), names

    def test_latest_alias_survives_copy_failure(self, tmp_path, capsys, monkeypatch):
        """差し替えに失敗しても、既存の sync-check.html は壊さない。"""
        cfg_path = _write_config(tmp_path, "html")
        config = load_config(cfg_path)
        out_dir = tmp_path / "reports"

        run(config, cfg_path, show_progress=False, log=_silent_logger())
        capsys.readouterr()
        original = (out_dir / "sync-check.html").read_bytes()

        import main as main_mod
        monkeypatch.setattr(
            main_mod.shutil, "copy2",
            lambda *a, **k: (_ for _ in ()).throw(OSError("disk full")),
        )
        with pytest.raises(OSError):
            run(config, cfg_path, show_progress=False, log=_silent_logger())
        capsys.readouterr()

        assert (out_dir / "sync-check.html").read_bytes() == original
        assert not any(f.name.endswith(".tmp") for f in out_dir.iterdir())

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
