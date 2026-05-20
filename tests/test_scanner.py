"""scanner: 除外パターン・シンボリックリンク・エラー記録の検証。"""
from __future__ import annotations

import os
import sys
from pathlib import Path

import pytest

from scanner import scan_location


def _scan(root: Path, exclude=None, workers=1):
    return scan_location(
        name="t",
        root=root,
        exclude_patterns=exclude or [],
        parallel_workers=workers,
        hash_algorithm="sha256",
        show_progress=False,
    )


class TestBasicScan:
    def test_empty_directory(self, tmp_path: Path):
        result = _scan(tmp_path)
        assert result.files == {}
        assert result.dirs == {}
        assert result.errors == []

    def test_single_file(self, tmp_path: Path):
        (tmp_path / "a.txt").write_text("hello", encoding="utf-8")
        result = _scan(tmp_path)
        assert list(result.files.keys()) == ["a.txt"]
        entry = result.files["a.txt"]
        assert entry.size == 5
        # sha256("hello")
        assert entry.hash == "2cf24dba5fb0a30e26e83b2ac5b9e29e1b161e5c1fa7425e73043362938b9824"

    def test_subdirectory_recursion(self, tmp_path: Path):
        (tmp_path / "sub").mkdir()
        (tmp_path / "sub" / "nested.txt").write_text("x")
        result = _scan(tmp_path)
        assert "sub/nested.txt" in result.files
        assert "sub" in result.dirs

    def test_relpath_uses_forward_slash(self, tmp_path: Path):
        (tmp_path / "a" / "b").mkdir(parents=True)
        (tmp_path / "a" / "b" / "deep.txt").write_text("x")
        result = _scan(tmp_path)
        assert "a/b/deep.txt" in result.files

    def test_nonexistent_root_returns_error(self, tmp_path: Path):
        result = _scan(tmp_path / "no_such_dir")
        assert result.files == {}
        assert len(result.errors) == 1
        assert "does not exist" in result.errors[0].message

    def test_root_is_file_returns_error(self, tmp_path: Path):
        f = tmp_path / "single.txt"
        f.write_text("x")
        result = _scan(f)
        assert len(result.errors) == 1
        assert "not a directory" in result.errors[0].message


class TestExcludePatterns:
    def test_filename_pattern(self, tmp_path: Path):
        (tmp_path / "keep.txt").write_text("k")
        (tmp_path / "~$lock.xlsx").write_text("l")
        (tmp_path / "skip.tmp").write_text("t")
        result = _scan(tmp_path, exclude=["~$*", "*.tmp"])
        assert set(result.files.keys()) == {"keep.txt"}

    def test_excluded_directory_not_descended(self, tmp_path: Path):
        (tmp_path / ".git").mkdir()
        (tmp_path / ".git" / "config").write_text("x")
        (tmp_path / ".git" / "objects").mkdir()
        (tmp_path / ".git" / "objects" / "deep").write_text("y")
        (tmp_path / "src" / "code.py").parent.mkdir()
        (tmp_path / "src" / "code.py").write_text("z")
        result = _scan(tmp_path, exclude=[".git"])
        # .git 配下のファイルが1件も含まれない
        assert all(not k.startswith(".git") for k in result.files), result.files
        assert "src/code.py" in result.files
        # ディレクトリ側も .git は記録されない
        assert ".git" not in result.dirs

    def test_empty_excludes_keeps_everything(self, tmp_path: Path):
        (tmp_path / "a").write_text("1")
        (tmp_path / "b").write_text("2")
        result = _scan(tmp_path, exclude=[])
        assert set(result.files.keys()) == {"a", "b"}


class TestSymlinks:
    def test_symlinked_directory_not_followed(self, tmp_path: Path):
        target = tmp_path / "target"
        target.mkdir()
        (target / "in_target.txt").write_text("t")
        try:
            (tmp_path / "link").symlink_to(target)
        except (OSError, NotImplementedError):
            pytest.skip("symlinks not supported on this platform")
        result = _scan(tmp_path)
        # target 直下は含まれるが、link 経由では辿らない
        assert "target/in_target.txt" in result.files
        assert "link/in_target.txt" not in result.files


class TestErrorRecording:
    def test_unreadable_file_recorded_in_file_errors(self, tmp_path: Path):
        if sys.platform == "win32":
            pytest.skip("chmod-based unreadable test is POSIX only")
        if os.geteuid() == 0:  # type: ignore[attr-defined]
            pytest.skip("running as root bypasses permission denial")
        f = tmp_path / "locked.txt"
        f.write_text("secret")
        f.chmod(0o000)
        try:
            result = _scan(tmp_path)
            # ファイル自体は file_errors に記録され、files には入らない
            assert "locked.txt" in result.file_errors
            assert "locked.txt" not in result.files
            # errors リストにも対応する ScanError がある
            assert any(e.relpath == "locked.txt" for e in result.errors)
        finally:
            f.chmod(0o600)  # クリーンアップのため戻す


class TestParallelEquivalence:
    """並列モードでもシリアルと同じ結果になる。"""

    def test_serial_and_parallel_match(self, tmp_path: Path):
        for i in range(20):
            (tmp_path / f"file_{i:02d}.txt").write_text(f"content-{i}")
        serial = _scan(tmp_path, workers=1)
        parallel = _scan(tmp_path, workers=4)
        # files dict は順序保証しないので key 集合 + hash で照合
        assert set(serial.files.keys()) == set(parallel.files.keys())
        for k in serial.files:
            assert serial.files[k].hash == parallel.files[k].hash
            assert serial.files[k].size == parallel.files[k].size
