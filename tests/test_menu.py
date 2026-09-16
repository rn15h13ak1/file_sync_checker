"""menu: 対話メニューの引数組み立て・入力処理・設定表示。

対話ループ本体 (`main`) は副作用が大きいため対象外にしている。
"""
from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pytest
import yaml

import menu


# ============================================================
# 本体に渡す引数の組み立て
# ============================================================
class TestBuildArgs:
    def test_empty_means_use_the_config_file(self):
        """何も選ばなければ引数を足さない (設定ファイルの値が使われる)。"""
        assert menu.build_args() == []

    def test_config_path_is_passed_first(self):
        args = menu.build_args(config_path="my.yaml")
        assert args[:2] == ["--config", "my.yaml"]

    def test_hash_mode_and_format(self):
        args = menu.build_args(hash_mode="smart", output_format="html")
        assert args == ["--hash-mode", "smart", "--format", "html"]

    def test_retry_zero_is_omitted(self):
        """再試行 0 は既定なので渡さない。"""
        assert menu.build_args(retry="0") == []

    def test_retry_is_passed_when_positive(self):
        assert menu.build_args(retry="3") == ["--retry", "3"]

    def test_built_args_are_accepted_by_the_tool(self):
        """組み立てた引数を本体が受け付けること (オプション名の取り違え検出)。"""
        import main as main_mod

        args = menu.build_args(
            config_path="c.yaml", hash_mode="smart", output_format="both", retry="2"
        )
        sys_argv = sys.argv
        try:
            sys.argv = ["main.py", *args]
            parsed = main_mod.parse_args()
        finally:
            sys.argv = sys_argv
        assert parsed.config == "c.yaml"
        assert parsed.hash_mode == "smart"
        assert parsed.output_format == "both"
        assert parsed.retry == 2


# ============================================================
# 入力
# ============================================================
class TestPrintMenu:
    def _answers(self, monkeypatch, values):
        it = iter(values)
        monkeypatch.setattr("builtins.input", lambda *a: next(it))

    def test_returns_selected_number(self, monkeypatch, capsys):
        self._answers(monkeypatch, ["2"])
        assert menu.print_menu("t", ["a", "b"]) == 2
        capsys.readouterr()

    def test_zero_returns_back(self, monkeypatch, capsys):
        self._answers(monkeypatch, ["0"])
        assert menu.print_menu("t", ["a"]) == 0
        capsys.readouterr()

    def test_empty_uses_default(self, monkeypatch, capsys):
        self._answers(monkeypatch, [""])
        assert menu.print_menu("t", ["a", "b"], default=2) == 2
        capsys.readouterr()

    def test_out_of_range_reprompts(self, monkeypatch, capsys):
        self._answers(monkeypatch, ["9", "1"])
        assert menu.print_menu("t", ["a"]) == 1
        assert "無効な入力" in capsys.readouterr().out

    def test_non_numeric_reprompts(self, monkeypatch, capsys):
        self._answers(monkeypatch, ["x", "1"])
        assert menu.print_menu("t", ["a"]) == 1
        capsys.readouterr()


class TestInputText:
    def _answers(self, monkeypatch, values):
        it = iter(values)
        monkeypatch.setattr("builtins.input", lambda *a: next(it))

    def test_returns_entered_value(self, monkeypatch, capsys):
        self._answers(monkeypatch, ["  5 "])
        assert menu.input_text("n") == "5"
        capsys.readouterr()

    def test_empty_returns_default(self, monkeypatch, capsys):
        self._answers(monkeypatch, [""])
        assert menu.input_text("n", default="2") == "2"
        capsys.readouterr()

    def test_empty_without_default_returns_none(self, monkeypatch, capsys):
        self._answers(monkeypatch, [""])
        assert menu.input_text("n") is None
        capsys.readouterr()

    def test_reprompts_until_valid(self, monkeypatch, capsys):
        self._answers(monkeypatch, ["-1", "abc", "3"])
        assert menu.input_text("n", validate=menu.is_retry_count, hint="整数") == "3"
        capsys.readouterr()


class TestIsRetryCount:
    @pytest.mark.parametrize("value, expected", [
        ("0", True), ("3", True), ("10", True),
        ("-1", False), ("1.5", False), ("abc", False), ("", False),
    ])
    def test_values(self, value, expected):
        assert menu.is_retry_count(value) is expected


# ============================================================
# 前回値の記憶
# ============================================================
class TestHistory:
    def test_roundtrip(self, tmp_path, monkeypatch):
        monkeypatch.setattr(menu, "HISTORY_PATH", tmp_path / "h.json")
        menu.save_history({"mode": "run", "retry": "2"})
        assert menu.load_history() == {"mode": "run", "retry": "2"}

    def test_missing_file_returns_empty(self, tmp_path, monkeypatch):
        monkeypatch.setattr(menu, "HISTORY_PATH", tmp_path / "none.json")
        assert menu.load_history() == {}

    def test_broken_file_returns_empty(self, tmp_path, monkeypatch):
        """壊れていても既定値で続行する (メニューが起動しないと困る)。"""
        broken = tmp_path / "h.json"
        broken.write_text("{ not json", encoding="utf-8")
        monkeypatch.setattr(menu, "HISTORY_PATH", broken)
        assert menu.load_history() == {}

    def test_non_dict_returns_empty(self, tmp_path, monkeypatch):
        path = tmp_path / "h.json"
        path.write_text("[1, 2]", encoding="utf-8")
        monkeypatch.setattr(menu, "HISTORY_PATH", path)
        assert menu.load_history() == {}

    def test_save_failure_is_ignored(self, tmp_path, monkeypatch):
        """保存できなくても実行は妨げない。"""
        monkeypatch.setattr(menu, "HISTORY_PATH", tmp_path / "nodir" / "h.json")
        menu.save_history({"mode": "run"})   # 例外を投げない


# ============================================================
# 設定の表示
# ============================================================
def _write_config(tmp_path: Path, **overrides) -> Path:
    a, b = tmp_path / "A", tmp_path / "B"
    a.mkdir(exist_ok=True)
    b.mkdir(exist_ok=True)
    data = {
        "locations": [{"name": "本社", "path": str(a)}, {"name": "大阪", "path": str(b)}],
        "exclude_patterns": ["*.tmp", "作業中/*"],
        "output": {"format": "html", "output_dir": str(tmp_path / "reports")},
    }
    data.update(overrides)
    cfg = tmp_path / "config.yaml"
    cfg.write_text(yaml.safe_dump(data, allow_unicode=True), encoding="utf-8")
    return cfg


class TestDescribeTargets:
    def test_lists_the_locations(self, tmp_path):
        cfg = _write_config(tmp_path)
        assert menu.describe_targets(str(cfg)) == "本社 / 大阪 (2 拠点)"

    def test_missing_config_explains_why(self, tmp_path):
        got = menu.describe_targets(str(tmp_path / "none.yaml"))
        assert got.startswith("※")
        assert "設定ファイルが見つかりません" in got

    def test_broken_config_explains_why(self, tmp_path):
        bad = tmp_path / "bad.yaml"
        bad.write_text("locations: [unclosed\n", encoding="utf-8")
        assert menu.describe_targets(str(bad)).startswith("※")


class TestPrintConfig:
    def test_shows_locations_and_patterns(self, tmp_path, capsys):
        menu.print_config(str(_write_config(tmp_path)))
        out = capsys.readouterr().out
        assert "本社" in out and "大阪" in out
        # 除外パターンは照合対象の違いまで示す
        assert "*.tmp  (ファイル名と照合)" in out
        assert "作業中/*  (相対パス全体と照合)" in out

    def test_shows_reason_when_unreadable(self, tmp_path, capsys):
        menu.print_config(str(tmp_path / "none.yaml"))
        assert "※" in capsys.readouterr().out

    def test_no_exclude_patterns(self, tmp_path, capsys):
        menu.print_config(str(_write_config(tmp_path, exclude_patterns=[])))
        assert "(なし)" in capsys.readouterr().out


# ============================================================
# レポートを開く
# ============================================================
class TestLatestReport:
    def test_returns_the_newest_report(self, tmp_path):
        cfg = _write_config(tmp_path)
        reports = tmp_path / "reports"
        reports.mkdir()
        import os
        old = reports / "sync-check-20260101-090000.html"
        new = reports / "sync-check-20260102-090000.html"
        old.write_text("o", encoding="utf-8")
        new.write_text("n", encoding="utf-8")
        os.utime(old, (1000, 1000))
        os.utime(new, (2000, 2000))
        assert menu.latest_report(str(cfg)) == new

    def test_ignores_unrelated_files(self, tmp_path):
        cfg = _write_config(tmp_path)
        reports = tmp_path / "reports"
        reports.mkdir()
        (reports / "メモ.txt").write_text("x", encoding="utf-8")
        assert menu.latest_report(str(cfg)) is None

    def test_missing_output_dir(self, tmp_path):
        assert menu.latest_report(str(_write_config(tmp_path))) is None

    def test_unreadable_config(self, tmp_path):
        assert menu.latest_report(str(tmp_path / "none.yaml")) is None


# ============================================================
# 終了コードの説明
# ============================================================
class TestExitMeanings:
    def test_covers_every_exit_code_of_the_tool(self):
        """本体が返しうる終了コードすべてに説明があること。"""
        import main as main_mod

        codes = {getattr(main_mod, n) for n in dir(main_mod) if n.startswith("EXIT_")}
        missing = codes - set(menu.EXIT_MEANINGS)
        assert not missing, f"説明の無い終了コード: {missing}"


# ============================================================
# 起動
# ============================================================
class TestLaunch:
    def test_help_runs(self):
        r = subprocess.run(
            [sys.executable, str(Path(menu.__file__)), "--help"],
            capture_output=True, text=True, timeout=60,
        )
        assert r.returncode == 0
        assert "--config" in r.stdout

    def test_eof_exits_cleanly(self, tmp_path):
        """入力が尽きたら (パイプ実行など) トレースバックを出さずに終わる。"""
        r = subprocess.run(
            [sys.executable, str(Path(menu.__file__)),
             "--config", str(_write_config(tmp_path))],
            input="", capture_output=True, text=True, timeout=60,
        )
        assert r.returncode == menu.EXIT_EOF
        assert "Traceback" not in r.stderr
