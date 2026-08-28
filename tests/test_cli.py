"""CLI を実際に起動して終了コードと引数の配線を確認する。

他のテストは run() や load_config() を直接呼ぶため、argparse の定義
(choices・dest 名) と main() への受け渡しが検証されていなかった。
ここだけは subprocess で本物のプロセスとして起動する。
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import yaml

REPO_ROOT = Path(__file__).resolve().parent.parent
MAIN = REPO_ROOT / "main.py"

EXIT_OK = 0
EXIT_DIFF = 1
EXIT_CONFIG_ERROR = 2
EXIT_USAGE = 2  # argparse が不正な引数で返す値


def _run(*args: str, cwd: Path | None = None) -> subprocess.CompletedProcess:
    return subprocess.run(
        [sys.executable, str(MAIN), *args],
        capture_output=True, text=True, cwd=cwd, timeout=120,
    )


def _make_config(tmp_path: Path, *, differ: bool = False, fmt: str = "html") -> Path:
    a, b = tmp_path / "locA", tmp_path / "locB"
    a.mkdir()
    b.mkdir()
    (a / "shared.txt").write_text("data", encoding="utf-8")
    (b / "shared.txt").write_text("data", encoding="utf-8")
    if differ:
        (a / "only_in_a.txt").write_text("x", encoding="utf-8")
    cfg = tmp_path / "config.yaml"
    cfg.write_text(yaml.safe_dump({
        "locations": [{"name": "A", "path": str(a)}, {"name": "B", "path": str(b)}],
        "output": {"format": fmt, "output_dir": str(tmp_path / "reports")},
        "performance": {"parallel_workers": 1, "hash_algorithm": "sha256"},
    }, allow_unicode=True))
    return cfg


class TestExitCodes:
    def test_match_returns_zero(self, tmp_path: Path):
        r = _run("-c", str(_make_config(tmp_path)), "--no-progress")
        assert r.returncode == EXIT_OK, r.stderr
        assert "スキャン結果サマリー" in r.stdout

    def test_difference_returns_one(self, tmp_path: Path):
        r = _run("-c", str(_make_config(tmp_path, differ=True)), "--no-progress")
        assert r.returncode == EXIT_DIFF, r.stderr

    def test_missing_config_returns_two(self, tmp_path: Path):
        r = _run("-c", str(tmp_path / "no_such.yaml"), "--no-progress")
        assert r.returncode == EXIT_CONFIG_ERROR
        assert "設定ファイルが見つかりません" in r.stderr


class TestArgumentWiring:
    def test_help_lists_override_options(self):
        r = _run("--help")
        assert r.returncode == EXIT_OK
        for opt in ("--format", "--output-dir", "--hash-mode", "--no-progress"):
            assert opt in r.stdout, opt

    def test_format_override_changes_output(self, tmp_path: Path):
        """設定は html だが --format excel で xlsx だけが出る。"""
        cfg = _make_config(tmp_path, fmt="html")
        r = _run("-c", str(cfg), "--no-progress", "--format", "excel")
        assert r.returncode == EXIT_OK, r.stderr
        names = [p.name for p in (tmp_path / "reports").iterdir()]
        assert any(n.endswith(".xlsx") for n in names), names
        assert not any(n.endswith(".html") for n in names), names

    def test_output_dir_override_is_relative_to_cwd(self, tmp_path: Path):
        """-o は設定ファイルではなく実行時のカレントから解決する。"""
        cfg = _make_config(tmp_path)
        workdir = tmp_path / "work"
        workdir.mkdir()
        r = _run("-c", str(cfg), "--no-progress", "-o", "./out", cwd=workdir)
        assert r.returncode == EXIT_OK, r.stderr
        assert (workdir / "out" / "sync-check.html").is_file()

    def test_hash_mode_override_is_applied(self, tmp_path: Path):
        """--hash-mode smart でハッシュ省略のログが出る。"""
        cfg = _make_config(tmp_path)
        r = _run("-c", str(cfg), "--no-progress", "-v", "--hash-mode", "smart")
        assert r.returncode == EXIT_OK, r.stderr
        assert "ハッシュ省略" in r.stderr

    def test_invalid_format_is_rejected_by_argparse(self, tmp_path: Path):
        r = _run("-c", str(_make_config(tmp_path)), "--format", "pdf")
        assert r.returncode == EXIT_USAGE
        assert "--format" in r.stderr

    def test_invalid_hash_mode_is_rejected_by_argparse(self, tmp_path: Path):
        r = _run("-c", str(_make_config(tmp_path)), "--hash-mode", "trust_me")
        assert r.returncode == EXIT_USAGE
        assert "--hash-mode" in r.stderr


class TestDefaultConfigLookup:
    def test_config_yaml_in_cwd_is_used_without_c(self, tmp_path: Path):
        """-c 省略時は CWD の config.yaml を使う。"""
        cfg = _make_config(tmp_path)
        cfg.rename(tmp_path / "config.yaml")  # 既にその名前だが意図を明示
        r = _run("--no-progress", cwd=tmp_path)
        assert r.returncode == EXIT_OK, r.stderr
        assert (tmp_path / "reports" / "sync-check.html").is_file()
