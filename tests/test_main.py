"""main.run の出力ファイル名整合性テスト。"""
from __future__ import annotations

import logging
from dataclasses import replace
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


class TestPruneOldReports:
    """古いレポートの削除。削除は元に戻せないので、対象の厳密さを重点的に確認する。"""

    def _make_reports(self, out_dir: Path, slugs: list) -> None:
        out_dir.mkdir(parents=True, exist_ok=True)
        for slug in slugs:
            (out_dir / f"sync-check-{slug}.html").write_text("h", encoding="utf-8")
            (out_dir / f"sync-check-{slug}.xlsx").write_text("x", encoding="utf-8")

    def test_keeps_the_newest_runs(self, tmp_path):
        from main import _prune_old_reports

        out = tmp_path / "reports"
        slugs = ["20260101-090000", "20260102-090000", "20260103-090000",
                 "20260104-090000", "20260105-090000"]
        self._make_reports(out, slugs)

        _prune_old_reports(out, keep=2, log=_silent_logger())

        remaining = sorted(p.name for p in out.iterdir())
        assert remaining == [
            "sync-check-20260104-090000.html", "sync-check-20260104-090000.xlsx",
            "sync-check-20260105-090000.html", "sync-check-20260105-090000.xlsx",
        ]

    def test_counts_runs_not_files(self, tmp_path):
        """1回の実行が html と xlsx を出していれば、それらで1回分。"""
        from main import _prune_old_reports

        out = tmp_path / "reports"
        self._make_reports(out, ["20260101-090000", "20260102-090000"])
        _prune_old_reports(out, keep=1, log=_silent_logger())
        assert len(list(out.iterdir())) == 2   # 新しい方の html と xlsx

    def test_never_deletes_the_stable_alias(self, tmp_path):
        """固定名 sync-check.html はブックマーク先なので消さない。"""
        from main import _prune_old_reports

        out = tmp_path / "reports"
        self._make_reports(out, ["20260101-090000", "20260102-090000"])
        (out / "sync-check.html").write_text("latest", encoding="utf-8")

        _prune_old_reports(out, keep=1, log=_silent_logger())
        assert (out / "sync-check.html").is_file()

    def test_never_deletes_unrelated_files(self, tmp_path):
        """このツールの命名規則に一致しないファイルには触れない。"""
        from main import _prune_old_reports

        out = tmp_path / "reports"
        self._make_reports(out, ["20260101-090000", "20260102-090000"])
        others = [
            "メモ.txt",
            "sync-check.xlsx",              # 固定名
            "sync-check-2026.html",         # 桁数が違う
            "sync-check-20260101.html",     # 時刻が無い
            "sync-check-20260101-090000.csv",   # 対象外の拡張子
            "old-sync-check-20260101-090000.html",  # 前置きがある
            "sync-check-20260101-090000.html.bak",
        ]
        for name in others:
            (out / name).write_text("keep", encoding="utf-8")
        (out / "サブフォルダ").mkdir()

        _prune_old_reports(out, keep=1, log=_silent_logger())

        for name in others:
            assert (out / name).is_file(), f"消してはいけない: {name}"
        assert (out / "サブフォルダ").is_dir()

    def test_zero_keeps_everything(self, tmp_path):
        """既定 (0) では削除しない。既に配布済みの環境で勝手に消さないため。"""
        from main import _prune_old_reports

        out = tmp_path / "reports"
        self._make_reports(out, ["20260101-090000", "20260102-090000"])
        _prune_old_reports(out, keep=0, log=_silent_logger())
        assert len(list(out.iterdir())) == 4

    def test_fewer_reports_than_keep_is_fine(self, tmp_path):
        from main import _prune_old_reports

        out = tmp_path / "reports"
        self._make_reports(out, ["20260101-090000"])
        _prune_old_reports(out, keep=10, log=_silent_logger())
        assert len(list(out.iterdir())) == 2

    def test_deletion_failure_does_not_stop_the_run(self, tmp_path, monkeypatch):
        """削除に失敗しても実行自体は成功させる (レポートは既に書けている)。"""
        from main import _prune_old_reports

        out = tmp_path / "reports"
        self._make_reports(out, ["20260101-090000", "20260102-090000"])
        monkeypatch.setattr(
            Path, "unlink",
            lambda self, **k: (_ for _ in ()).throw(OSError("busy")),
        )
        _prune_old_reports(out, keep=1, log=_silent_logger())   # 例外を投げない
        assert len(list(out.iterdir())) == 4

    def test_end_to_end_through_run(self, tmp_path, capsys):
        """run() から実際に古いレポートが消える。"""
        cfg_path = _write_config(tmp_path, "both")
        out = tmp_path / "reports"
        self._make_reports(out, ["20250101-090000", "20250102-090000"])

        config = load_config(cfg_path)
        config = replace(config, output=replace(config.output, keep_reports=1))
        run(config, cfg_path, show_progress=False, log=_silent_logger())
        capsys.readouterr()

        names = sorted(p.name for p in out.iterdir())
        # 今回の実行分 (html/xlsx) と固定名だけが残る
        assert "sync-check.html" in names
        assert not any(n.startswith("sync-check-2025") for n in names), names


class TestProgressBarSuppression:
    """端末でなければ進捗バーを出さない。

    cron などで stderr をファイルに落としていると、tqdm の更新がそのまま
    ログに書き込まれる (実測: 10 秒のスキャンで stderr の 97% が進捗バー由来)。
    """

    def _capture_show_progress(self, tmp_path, monkeypatch, *, isatty: bool, argv: list):
        import main as main_mod

        cfg_path = _write_config(tmp_path, "html")
        monkeypatch.setattr(sys, "argv", ["main.py", "-c", str(cfg_path), *argv])

        class FakeStderr:
            def isatty(self):
                return isatty

            def write(self, *a):
                pass

            def flush(self):
                pass

        monkeypatch.setattr(sys, "stderr", FakeStderr())
        captured = {}

        def fake_run(config, config_path, *, show_progress, log, retry=0):
            captured["show_progress"] = show_progress
            return EXIT_OK

        monkeypatch.setattr(main_mod, "run", fake_run)
        assert main_mod.main() == EXIT_OK
        return captured["show_progress"]

    def test_disabled_when_stderr_is_not_a_tty(self, tmp_path, monkeypatch):
        assert self._capture_show_progress(
            tmp_path, monkeypatch, isatty=False, argv=[]
        ) is False

    def test_enabled_on_a_terminal(self, tmp_path, monkeypatch):
        assert self._capture_show_progress(
            tmp_path, monkeypatch, isatty=True, argv=[]
        ) is True

    def test_no_progress_wins_on_a_terminal(self, tmp_path, monkeypatch):
        """端末でも --no-progress を付ければ出さない。"""
        assert self._capture_show_progress(
            tmp_path, monkeypatch, isatty=True, argv=["--no-progress"]
        ) is False


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


class TestComparisonReports:
    """複数の比較を定義したときのレポート名と削除範囲。"""

    def _write_multi_config(self, tmp_path: Path) -> Path:
        for d in ("a", "b", "c", "d"):
            loc = tmp_path / d
            loc.mkdir()
            (loc / "shared.txt").write_text("data", encoding="utf-8")
        cfg = tmp_path / "config.yaml"
        cfg.write_text(yaml.safe_dump({
            "comparisons": {
                "契約書": [{"name": "本社", "path": str(tmp_path / "a")},
                           {"name": "大阪", "path": str(tmp_path / "b")}],
                "設計": [{"name": "本社", "path": str(tmp_path / "c")},
                         {"name": "大阪", "path": str(tmp_path / "d")}],
            },
            "output": {"format": "both", "output_dir": str(tmp_path / "reports")},
        }, allow_unicode=True), encoding="utf-8")
        return cfg

    def test_report_name_includes_the_comparison(self, tmp_path, capsys):
        cfg = self._write_multi_config(tmp_path)
        run(load_config(cfg, comparison="契約書"), cfg,
            show_progress=False, log=_silent_logger())
        capsys.readouterr()

        names = sorted(p.name for p in (tmp_path / "reports").iterdir())
        assert any(n.startswith("sync-check-契約書-") for n in names), names
        # 固定名も比較ごとに分かれる
        assert "sync-check-契約書.html" in names

    def test_comparisons_do_not_overwrite_each_other(self, tmp_path, capsys):
        cfg = self._write_multi_config(tmp_path)
        for name in ("契約書", "設計"):
            run(load_config(cfg, comparison=name), cfg,
                show_progress=False, log=_silent_logger())
            capsys.readouterr()

        names = {p.name for p in (tmp_path / "reports").iterdir()}
        assert "sync-check-契約書.html" in names
        assert "sync-check-設計.html" in names

    def test_pruning_is_per_comparison(self, tmp_path, capsys):
        """比較ごとに keep 回分を残す (互いのレポートを消し合わない)。"""
        from main import _prune_old_reports

        out = tmp_path / "reports"
        out.mkdir()
        for name in ("契約書", "設計"):
            for slug in ("20260101-090000", "20260102-090000", "20260103-090000"):
                (out / f"sync-check-{name}-{slug}.html").write_text("x", encoding="utf-8")

        _prune_old_reports(out, keep=1, log=_silent_logger(), comparison="契約書")

        remaining = sorted(p.name for p in out.iterdir())
        # 契約書は直近1回分だけ、設計は手つかず
        assert remaining == [
            "sync-check-契約書-20260103-090000.html",
            "sync-check-設計-20260101-090000.html",
            "sync-check-設計-20260102-090000.html",
            "sync-check-設計-20260103-090000.html",
        ]

    def test_pruning_unnamed_ignores_named_reports(self, tmp_path):
        """従来形式の実行は、名前付きレポートに触れない。"""
        from main import _prune_old_reports

        out = tmp_path / "reports"
        out.mkdir()
        for slug in ("20260101-090000", "20260102-090000"):
            (out / f"sync-check-{slug}.html").write_text("x", encoding="utf-8")
        (out / "sync-check-契約書-20260101-090000.html").write_text("x", encoding="utf-8")

        _prune_old_reports(out, keep=1, log=_silent_logger(), comparison="")

        remaining = sorted(p.name for p in out.iterdir())
        assert remaining == [
            "sync-check-20260102-090000.html",
            "sync-check-契約書-20260101-090000.html",
        ]

    def test_comparison_is_recorded_in_the_report(self, tmp_path, capsys):
        cfg = self._write_multi_config(tmp_path)
        run(load_config(cfg, comparison="設計"), cfg,
            show_progress=False, log=_silent_logger())
        capsys.readouterr()

        html = (tmp_path / "reports" / "sync-check-設計.html").read_text(encoding="utf-8")
        conditions = html.split("<h3>実行条件</h3>")[1].split("</table>")[0]
        assert "設計" in conditions

    def test_list_comparisons_exits_without_scanning(self, tmp_path, monkeypatch, capsys):
        import main as main_mod

        cfg = self._write_multi_config(tmp_path)
        monkeypatch.setattr(sys, "argv",
                            ["main.py", "-c", str(cfg), "--list-comparisons"])
        assert main_mod.main() == EXIT_OK
        out = capsys.readouterr().out
        assert "契約書" in out and "設計" in out
        assert not (tmp_path / "reports").exists()


class TestSingleComparisonSameShape:
    """比較が 1 組でも、複数と同じ `comparisons:` の書き方で動くこと。

    組の数で設定の書き方が変わらないことを固定する。
    """

    def _write(self, tmp_path: Path, names: list) -> Path:
        comparisons = {}
        for i, name in enumerate(names):
            a, b = tmp_path / f"{name}A", tmp_path / f"{name}B"
            for d in (a, b):
                d.mkdir()
                (d / "shared.txt").write_text("data", encoding="utf-8")
            comparisons[name] = [{"name": "本社", "path": str(a)},
                                 {"name": "大阪", "path": str(b)}]
        cfg = tmp_path / "config.yaml"
        cfg.write_text(yaml.safe_dump({
            "comparisons": comparisons,
            "output": {"format": "html", "output_dir": str(tmp_path / "reports")},
        }, allow_unicode=True), encoding="utf-8")
        return cfg

    def test_single_entry_needs_no_comparison_argument(self, tmp_path, monkeypatch, capsys):
        """1 組だけなら --comparison を付けずに実行できる。"""
        import main as main_mod

        cfg = self._write(tmp_path, ["契約書"])
        monkeypatch.setattr(sys, "argv", ["main.py", "-c", str(cfg), "--no-progress"])
        assert main_mod.main() == EXIT_OK
        capsys.readouterr()
        assert (tmp_path / "reports" / "sync-check-契約書.html").is_file()

    def test_adding_a_second_entry_only_changes_the_invocation(
        self, tmp_path, monkeypatch, capsys
    ):
        """2 組目を足しても設定の書き方は同じ。変わるのは実行時の指定だけ。"""
        import main as main_mod

        cfg = self._write(tmp_path, ["契約書", "設計"])
        # 名前を指定すれば動く
        monkeypatch.setattr(sys, "argv",
                            ["main.py", "-c", str(cfg), "--no-progress",
                             "--comparison", "契約書"])
        assert main_mod.main() == EXIT_OK
        capsys.readouterr()
        # 省略すると、どれを実行するか決められないので設定エラー
        monkeypatch.setattr(sys, "argv", ["main.py", "-c", str(cfg), "--no-progress"])
        assert main_mod.main() == EXIT_CONFIG_ERROR
        capsys.readouterr()

    def test_single_entry_reports_carry_the_name(self, tmp_path, capsys):
        """1 組でもレポート名に比較名が入る (複数と同じ扱い)。"""
        cfg = self._write(tmp_path, ["契約書"])
        run(load_config(cfg), cfg, show_progress=False, log=_silent_logger())
        capsys.readouterr()
        names = [p.name for p in (tmp_path / "reports").iterdir()]
        assert all("契約書" in n for n in names), names
