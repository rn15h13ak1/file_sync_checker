"""scanner: 除外パターン・シンボリックリンク・エラー記録の検証。"""
from __future__ import annotations

import os
import sys
import threading
import time
from pathlib import Path

import pytest

from scanner import (
    HASH_CHUNK_SIZE,
    HASH_MODE_ALWAYS,
    HASH_MODE_SMART,
    ScanCancelled,
    _hash_file,
    _managed_pool,
    plan_hash_targets,
    scan_location,
    scan_locations,
    stat_location,
)


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


# ============================================================
# フェーズ 1 / フェーズ 2 の分離
# ============================================================
def _write(path: Path, content: str, mtime: float | None = None) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")
    if mtime is not None:
        os.utime(path, (mtime, mtime))
    return path


class TestStatLocation:
    def test_collects_metadata_without_reading_content(self, tmp_path: Path):
        _write(tmp_path / "a.txt", "hello")
        result = stat_location("t", tmp_path, exclude_patterns=[])
        assert result.stats["a.txt"].size == 5
        assert result.stats["a.txt"].mtime_ns > 0
        # フェーズ 1 ではハッシュを持たない (FileStat にフィールドが無い)
        assert not hasattr(result.stats["a.txt"], "hash")

    def test_excluded_directory_not_descended(self, tmp_path: Path):
        _write(tmp_path / ".git" / "objects" / "deep", "y")
        _write(tmp_path / "src" / "code.py", "z")
        result = stat_location("t", tmp_path, exclude_patterns=[".git"])
        assert set(result.stats) == {"src/code.py"}
        assert ".git" not in result.dirs

    def test_nonexistent_root(self, tmp_path: Path):
        result = stat_location("t", tmp_path / "nope", exclude_patterns=[])
        assert result.stats == {}
        assert "does not exist" in result.errors[0].message


class TestPlanHashTargets:
    def _stats(self, tmp_path: Path, spec: dict) -> list:
        """spec: {拠点名: {relpath: (content, mtime)}} からフェーズ1結果を作る。"""
        results = []
        for loc, files in spec.items():
            root = tmp_path / loc
            root.mkdir()
            for rel, (content, mtime) in files.items():
                _write(root / rel, content, mtime)
            results.append(stat_location(loc, root, exclude_patterns=[]))
        return results

    def test_always_mode_hashes_everything(self, tmp_path: Path):
        stats = self._stats(tmp_path, {
            "A": {"same.txt": ("x", 1000), "diff.txt": ("xx", 1000)},
            "B": {"same.txt": ("x", 1000), "diff.txt": ("yyy", 1000)},
        })
        targets = plan_hash_targets(stats, hash_mode=HASH_MODE_ALWAYS)
        assert targets == {"same.txt", "diff.txt"}

    def test_smart_skips_identical_size_and_mtime(self, tmp_path: Path):
        stats = self._stats(tmp_path, {
            "A": {"same.txt": ("x", 1000)},
            "B": {"same.txt": ("x", 1000)},
        })
        assert plan_hash_targets(stats, hash_mode=HASH_MODE_SMART) == set()

    def test_smart_skips_when_size_differs(self, tmp_path: Path):
        """サイズが違えば内容も違うことが確定するため、ハッシュは結論を変えない。"""
        stats = self._stats(tmp_path, {
            "A": {"f.txt": ("x", 1000)},
            "B": {"f.txt": ("xxxx", 1000)},
        })
        assert plan_hash_targets(stats, hash_mode=HASH_MODE_SMART) == set()

    def test_smart_hashes_when_mtime_differs(self, tmp_path: Path):
        """サイズが同じで更新日時が違う = 内容が同じか判断できないのでハッシュする。"""
        stats = self._stats(tmp_path, {
            "A": {"f.txt": ("x", 1000)},
            "B": {"f.txt": ("y", 99999)},
        })
        assert plan_hash_targets(stats, hash_mode=HASH_MODE_SMART) == {"f.txt"}

    def test_smart_tolerates_mtime_within_tolerance(self, tmp_path: Path):
        """ファイルシステム間の更新日時の丸め (最大2秒) を同一とみなす。"""
        stats = self._stats(tmp_path, {
            "A": {"f.txt": ("x", 1000.0)},
            "B": {"f.txt": ("x", 1001.5)},
        })
        assert plan_hash_targets(stats, hash_mode=HASH_MODE_SMART) == set()
        # 許容誤差を 0 にすれば対象になる
        assert plan_hash_targets(
            stats, hash_mode=HASH_MODE_SMART, mtime_tolerance_sec=0
        ) == {"f.txt"}

    def test_smart_hashes_single_location_files(self, tmp_path: Path):
        """1拠点にしか無いファイルは比較不要だが、レポート表示のためハッシュする。"""
        stats = self._stats(tmp_path, {
            "A": {"only_a.txt": ("x", 1000)},
            "B": {},
        })
        assert plan_hash_targets(stats, hash_mode=HASH_MODE_SMART) == {"only_a.txt"}


class TestScanLocations:
    def test_smart_mode_leaves_hash_none_for_skipped(self, tmp_path: Path):
        for loc in ("A", "B"):
            _write(tmp_path / loc / "same.txt", "x", 1000)
            _write(tmp_path / loc / "changed.txt", "x", 1000 if loc == "A" else 50000)
        (tmp_path / "B" / "changed.txt").write_text("z", encoding="utf-8")
        os.utime(tmp_path / "B" / "changed.txt", (50000, 50000))

        scans = scan_locations(
            [("A", tmp_path / "A"), ("B", tmp_path / "B")],
            exclude_patterns=[],
            parallel_workers=2,
            hash_algorithm="sha256",
            hash_mode=HASH_MODE_SMART,
            show_progress=False,
        )
        a, b = scans
        # 一致ファイルはハッシュ未計算
        assert a.files["same.txt"].hash is None
        assert b.files["same.txt"].hash is None
        assert a.skipped_hashes == 1
        # 更新日時が違うファイルはハッシュ済み
        assert a.files["changed.txt"].hash is not None
        assert a.files["changed.txt"].hash != b.files["changed.txt"].hash

    def test_always_mode_matches_per_location_scan(self, tmp_path: Path):
        """always では従来の scan_location と同じ結果になる。"""
        for loc in ("A", "B"):
            _write(tmp_path / loc / "f.txt", f"content-{loc}")
        scans = scan_locations(
            [("A", tmp_path / "A"), ("B", tmp_path / "B")],
            exclude_patterns=[],
            parallel_workers=2,
            hash_algorithm="sha256",
            show_progress=False,
        )
        for scan, loc in zip(scans, ("A", "B")):
            single = _scan(tmp_path / loc)
            assert scan.files["f.txt"].hash == single.files["f.txt"].hash

    def test_reports_stat_progress_per_location(self, tmp_path: Path):
        for loc in ("A", "B", "C"):
            _write(tmp_path / loc / "f.txt", "x")
        seen = []
        scan_locations(
            [(n, tmp_path / n) for n in ("A", "B", "C")],
            exclude_patterns=[],
            parallel_workers=2,
            hash_algorithm="sha256",
            show_progress=False,
            on_stat_done=lambda sr: seen.append(sr.location_name),
        )
        # 並列実行なので順序は問わない
        assert sorted(seen) == ["A", "B", "C"]

    def test_result_order_matches_input_order(self, tmp_path: Path):
        """フェーズ1を並列化しても拠点の並び順は設定順を保つ。"""
        for loc in ("A", "B", "C"):
            _write(tmp_path / loc / "f.txt", "x")
        scans = scan_locations(
            [(n, tmp_path / n) for n in ("A", "B", "C")],
            exclude_patterns=[],
            parallel_workers=2,
            hash_algorithm="sha256",
            show_progress=False,
        )
        assert [s.location_name for s in scans] == ["A", "B", "C"]

    def test_hash_file_stops_on_cancel(self, tmp_path: Path):
        """中断フラグが立っていれば大きいファイルの途中でも止まる。"""
        big = tmp_path / "big.bin"
        big.write_bytes(b"x" * (HASH_CHUNK_SIZE * 3))
        cancel = threading.Event()
        cancel.set()
        with pytest.raises(ScanCancelled):
            _hash_file(big, "sha256", cancel)

    def test_pool_does_not_wait_for_queued_work_on_interrupt(self):
        """中断時にキュー済みタスクの完了を待たない。

        `with ThreadPoolExecutor(...)` は shutdown(wait=True) するため、
        Ctrl+C を押してもキュー全部が終わるまで戻らなかった。
        ここでは 20 タスク × 0.5 秒 (直列なら 10 秒) を 1 ワーカーに積み、
        例外送出後すぐ戻ることを確認する。
        """
        cancel = threading.Event()
        started = threading.Event()

        def slow():
            started.set()
            for _ in range(50):
                if cancel.is_set():
                    raise ScanCancelled()
                time.sleep(0.01)

        begin = time.monotonic()
        with pytest.raises(RuntimeError):
            with _managed_pool(1, cancel) as ex:
                for _ in range(20):
                    ex.submit(slow)
                started.wait(timeout=2)
                raise RuntimeError("interrupted")
        elapsed = time.monotonic() - begin
        assert cancel.is_set()
        assert elapsed < 2.0, f"中断に {elapsed:.1f} 秒かかった (キューを待っている)"

    def test_unreadable_file_recorded_as_error(self, tmp_path: Path):
        if sys.platform == "win32":
            pytest.skip("chmod-based unreadable test is POSIX only")
        if os.geteuid() == 0:  # type: ignore[attr-defined]
            pytest.skip("running as root bypasses permission denial")
        for loc in ("A", "B"):
            _write(tmp_path / loc / "locked.txt", "secret")
        # 片方だけ読めなくする (サイズは同じなので smart でもハッシュ対象になるよう mtime をずらす)
        os.utime(tmp_path / "B" / "locked.txt", (50000, 50000))
        (tmp_path / "B" / "locked.txt").chmod(0o000)
        try:
            scans = scan_locations(
                [("A", tmp_path / "A"), ("B", tmp_path / "B")],
                exclude_patterns=[],
                parallel_workers=2,
                hash_algorithm="sha256",
                hash_mode=HASH_MODE_SMART,
                show_progress=False,
            )
            b = scans[1]
            assert "locked.txt" in b.file_errors
            assert "locked.txt" not in b.files
        finally:
            (tmp_path / "B" / "locked.txt").chmod(0o600)
