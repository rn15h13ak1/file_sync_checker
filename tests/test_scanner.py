"""scanner: 除外パターン・シンボリックリンク・エラー記録の検証。"""
from __future__ import annotations

import os
import sys
import threading
import time
import unicodedata
from pathlib import Path

import pytest

from .conftest import skip_or_fail

import scanner

from scanner import (
    HASH_CHUNK_SIZE,
    ExcludeMatcher,
    _hash_task,
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

    def test_path_pattern_excludes_by_relative_path(self, tmp_path: Path):
        """`/` を含むパターンは相対パス全体と照合する。

        以前はファイル名としか照合せず、パスを書いても 1 件も除外されないのに
        エラーも警告も出なかった。
        """
        (tmp_path / "作業中").mkdir()
        (tmp_path / "納品").mkdir()
        (tmp_path / "作業中" / "資料.docx").write_text("x")
        (tmp_path / "納品" / "資料.docx").write_text("y")

        assert set(_scan(tmp_path, exclude=["作業中/*"]).files) == {"納品/資料.docx"}
        assert set(_scan(tmp_path, exclude=["納品/*"]).files) == {"作業中/資料.docx"}
        assert set(_scan(tmp_path, exclude=["*/資料.docx"]).files) == set()

    def test_deep_path_pattern(self, tmp_path: Path):
        (tmp_path / "a" / "一時" / "b").mkdir(parents=True)
        (tmp_path / "a" / "一時" / "b" / "x.txt").write_text("x")
        (tmp_path / "a" / "keep.txt").write_text("k")
        result = _scan(tmp_path, exclude=["a/一時/*"])
        assert set(result.files) == {"a/keep.txt"}

    def test_name_pattern_still_matches_at_any_depth(self, tmp_path: Path):
        """`/` を含まないパターンは従来どおりファイル名照合 (深さを問わない)。"""
        (tmp_path / "sub" / "deep").mkdir(parents=True)
        (tmp_path / "sub" / "deep" / "a.tmp").write_text("x")
        (tmp_path / "b.tmp").write_text("y")
        (tmp_path / "keep.txt").write_text("k")
        assert set(_scan(tmp_path, exclude=["*.tmp"]).files) == {"keep.txt"}

    def test_path_pattern_can_exclude_a_directory_subtree(self, tmp_path: Path):
        """ディレクトリにパスパターンが当たれば配下ごと降りない。"""
        (tmp_path / "old" / "2019").mkdir(parents=True)
        (tmp_path / "old" / "2019" / "x.txt").write_text("x")
        (tmp_path / "new.txt").write_text("n")
        result = _scan(tmp_path, exclude=["old/2019"])
        assert set(result.files) == {"new.txt"}
        assert "old/2019" not in result.dirs

    def test_empty_excludes_keeps_everything(self, tmp_path: Path):
        (tmp_path / "a").write_text("1")
        (tmp_path / "b").write_text("2")
        result = _scan(tmp_path, exclude=[])
        assert set(result.files.keys()) == {"a", "b"}


class TestWalkProgress:
    """列挙フェーズの進捗報告。

    大きい共有では列挙だけで数分かかることがあり、その間なにも表示されないと
    止まっているのか進んでいるのか分からない。
    """

    def test_reports_every_entry_examined(self, tmp_path: Path):
        for i in range(30):
            _write(tmp_path / f"f{i}.txt", "x")
        (tmp_path / "sub").mkdir()
        _write(tmp_path / "sub" / "deep.txt", "y")

        seen = []
        stat_location(
            "A", tmp_path, exclude_patterns=[], on_progress=seen.append
        )
        # ファイル30 + ディレクトリ1 + sub 配下1 = 32
        assert sum(seen) == 32

    def test_reports_incrementally_for_large_trees(self, tmp_path: Path, monkeypatch):
        """完了時にまとめてではなく、途中でも報告する。"""
        monkeypatch.setattr(scanner, "_WALK_PROGRESS_INTERVAL", 10)
        for d in range(5):
            for i in range(10):
                _write(tmp_path / f"d{d}" / f"f{i}.txt", "x")

        seen = []
        result = stat_location(
            "A", tmp_path, exclude_patterns=[], on_progress=seen.append
        )
        assert len(result.stats) == 50
        assert len(seen) > 1, "1回にまとめて報告されている"
        assert sum(seen) == 55      # ファイル50 + ディレクトリ5

    def test_excluded_entries_are_still_counted(self, tmp_path: Path):
        """除外したエントリも走査はしているので進捗に含める。"""
        _write(tmp_path / "keep.txt", "x")
        _write(tmp_path / "skip.tmp", "y")
        seen = []
        stat_location(
            "A", tmp_path, exclude_patterns=["*.tmp"], on_progress=seen.append
        )
        assert sum(seen) == 2

    def test_no_callback_is_fine(self, tmp_path: Path):
        _write(tmp_path / "f.txt", "x")
        assert len(stat_location("A", tmp_path, exclude_patterns=[]).stats) == 1

    def test_scan_locations_without_progress_does_not_create_a_bar(
        self, tmp_path: Path, monkeypatch
    ):
        """show_progress=False なら tqdm を作らない。"""
        created = []
        real_tqdm = scanner.tqdm

        def spy(*a, **k):
            created.append(k.get("desc"))
            return real_tqdm(*a, **k)

        monkeypatch.setattr(scanner, "tqdm", spy)
        for loc in ("A", "B"):
            _write(tmp_path / loc / "f.txt", "x")
        scan_locations(
            [("A", tmp_path / "A"), ("B", tmp_path / "B")],
            exclude_patterns=[], parallel_workers=1,
            hash_algorithm="sha256", show_progress=False,
        )
        assert created == []

    def test_scan_locations_with_progress_creates_the_walk_bar(
        self, tmp_path: Path, monkeypatch
    ):
        created = []
        real_tqdm = scanner.tqdm

        def spy(*a, **k):
            created.append(k.get("desc"))
            return real_tqdm(*a, **k)

        monkeypatch.setattr(scanner, "tqdm", spy)
        for loc in ("A", "B"):
            _write(tmp_path / loc / "f.txt", "x")
        scan_locations(
            [("A", tmp_path / "A"), ("B", tmp_path / "B")],
            exclude_patterns=[], parallel_workers=1,
            hash_algorithm="sha256", show_progress=True,
        )
        assert "列挙中" in created


class TestExcludeMatcher:
    """事前コンパイルした除外判定が fnmatch と同じ結果になること。

    エントリ×パターンの回数だけ fnmatch を呼ぶと内部で normcase が
    2 回走り、列挙フェーズの 4 割を占めていたため 1 本の正規表現にまとめた。
    速度のために判定がずれては意味がないので、fnmatch と突き合わせる。
    """

    CASES = [
        (["*.tmp"], "a.tmp", "sub/a.tmp"),
        (["~$*"], "~$book.xlsx", "d/~$book.xlsx"),
        ([".DS_Store"], ".DS_Store", "x/.DS_Store"),
        (["*.tmp", "*.bak"], "a.bak", "a.bak"),
        (["作業中/*"], "資料.docx", "作業中/資料.docx"),
        (["*/一時/*"], "x.txt", "a/一時/x.txt"),
        # glob として特別な意味を持たない文字が正規表現として解釈されないこと
        (["a+b.txt"], "a+b.txt", "a+b.txt"),
        (["report(1).docx"], "report(1).docx", "report(1).docx"),
    ]

    @pytest.mark.parametrize("patterns, name, relpath", CASES)
    def test_matches_agree_with_fnmatch(self, patterns, name, relpath):
        import fnmatch as fn

        expected = any(
            fn.fnmatch(relpath if "/" in p else name, p) for p in patterns
        )
        assert ExcludeMatcher(patterns).matches(name, relpath) is expected

    def test_non_matching_names_are_kept(self):
        m = ExcludeMatcher(["*.tmp", "作業中/*"])
        assert m.matches("keep.txt", "納品/keep.txt") is False
        assert m.matches("keep.txt", "作業中/keep.txt") is True
        assert m.matches("x.tmp", "納品/x.tmp") is True

    def test_empty_patterns_match_nothing(self):
        m = ExcludeMatcher([])
        assert not m
        assert m.matches("anything.txt", "a/anything.txt") is False

    def test_patterns_do_not_bleed_into_each_other(self):
        """複数パターンを 1 本の正規表現にまとめても、部分一致で誤爆しない。"""
        m = ExcludeMatcher(["*.tmp", "*.bak"])
        assert m.matches("a.tmpx", "a.tmpx") is False
        assert m.matches("tmp", "tmp") is False


class TestSymlinks:
    def test_symlinked_directory_not_followed(self, tmp_path: Path):
        target = tmp_path / "target"
        target.mkdir()
        (target / "in_target.txt").write_text("t")
        try:
            (tmp_path / "link").symlink_to(target)
        except (OSError, NotImplementedError):
            skip_or_fail("symlinks not supported on this platform")
        result = _scan(tmp_path)
        # target 直下は含まれるが、link 経由では辿らない
        assert "target/in_target.txt" in result.files
        assert "link/in_target.txt" not in result.files


class TestErrorRecording:
    def test_unreadable_file_recorded_in_file_errors(self, tmp_path: Path):
        if sys.platform == "win32":
            skip_or_fail("chmod-based unreadable test is POSIX only")
        if os.geteuid() == 0:  # type: ignore[attr-defined]
            skip_or_fail("running as root bypasses permission denial")
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

    def test_unreadable_directory_recorded_and_scan_continues(self, tmp_path: Path):
        """権限の無いディレクトリは記録して、他の拠点/フォルダの走査は続ける。"""
        if sys.platform == "win32":
            skip_or_fail("chmod-based unreadable test is POSIX only")
        if os.geteuid() == 0:  # type: ignore[attr-defined]
            skip_or_fail("running as root bypasses permission denial")
        _write(tmp_path / "readable" / "ok.txt", "x")
        secret = tmp_path / "secret"
        _write(secret / "hidden.txt", "y")
        secret.chmod(0o000)
        try:
            result = stat_location("t", tmp_path, exclude_patterns=[])
            # 読めるファイルは拾えている
            assert "readable/ok.txt" in result.stats
            # 読めないディレクトリの中身は入らず、エラーとして残る
            assert not any(k.startswith("secret/") for k in result.stats)
            assert any("walk error" in e.message for e in result.errors)
            # ディレクトリ自体は列挙できているので dirs には載る
            assert "secret" in result.dirs
        finally:
            secret.chmod(0o700)

    def test_broken_symlink_recorded_as_file_error(self, tmp_path: Path):
        """リンク切れは stat で失敗する。走査を止めずファイル単位のエラーにする。"""
        try:
            (tmp_path / "dangling.txt").symlink_to(tmp_path / "no_such_target")
        except (OSError, NotImplementedError):
            skip_or_fail("symlinks not supported on this platform")
        _write(tmp_path / "ok.txt", "x")

        result = stat_location("t", tmp_path, exclude_patterns=[])
        assert "ok.txt" in result.stats
        assert "dangling.txt" not in result.stats
        assert "dangling.txt" in result.file_errors
        assert "FileNotFoundError" in result.file_errors["dangling.txt"]
        assert any(e.relpath == "dangling.txt" for e in result.errors)


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

    def test_mtime_comparison_is_timezone_independent(self, tmp_path: Path):
        """更新日時の一致判定はタイムゾーンに影響されない。

        表示用の `mtime` はローカル時刻の datetime だが、判定には
        エポックからの `mtime_ns` を使う。拠点ごとにタイムゾーン設定の違う
        マシン経由でマウントしても判定結果が変わらないようにするため。
        """
        stats = self._stats(tmp_path, {
            "A": {"f.txt": ("x", 1_700_000_000)},
            "B": {"f.txt": ("x", 1_700_000_000)},
        })
        a, b = stats
        # 同じ瞬間なので mtime_ns は一致する
        assert a.stats["f.txt"].mtime_ns == b.stats["f.txt"].mtime_ns
        assert plan_hash_targets(stats, hash_mode=HASH_MODE_SMART) == set()

        # 表示用 datetime は naive (ローカル時刻)。判定には使わない
        assert a.stats["f.txt"].mtime.tzinfo is None

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

    def test_nfd_and_nfc_names_match_as_one_file(self, tmp_path: Path):
        """macOS (NFD) と Windows (NFC) で同じ日本語ファイル名が別物にならない。

        濁点付きの日本語ファイル名は macOS では分解形で返るため、
        正規化しないと「欠落」+「余分」として二重計上される。
        """
        name = "議事録_ガバナンス部会.docx"
        _write(tmp_path / "A" / unicodedata.normalize("NFC", name), "same")
        _write(tmp_path / "B" / unicodedata.normalize("NFD", name), "same")

        scans = scan_locations(
            [("A", tmp_path / "A"), ("B", tmp_path / "B")],
            exclude_patterns=[], parallel_workers=2,
            hash_algorithm="sha256", show_progress=False,
        )
        a, b = scans
        assert list(a.files) == list(b.files), "照合キーが一致していない"
        # 実ファイルを開けている (ハッシュが取れている = 元のファイル名を保持できている)
        key = list(a.files)[0]
        assert a.files[key].hash == b.files[key].hash
        assert unicodedata.is_normalized("NFC", key)
        # 元のファイル名は拠点ごとに保持される
        assert unicodedata.is_normalized("NFD", b.real_relpaths[key])

    def test_normalization_can_be_disabled(self, tmp_path: Path):
        name = "議事録_ガバナンス部会.docx"
        _write(tmp_path / "A" / unicodedata.normalize("NFC", name), "same")
        _write(tmp_path / "B" / unicodedata.normalize("NFD", name), "same")
        scans = scan_locations(
            [("A", tmp_path / "A"), ("B", tmp_path / "B")],
            exclude_patterns=[], parallel_workers=2,
            hash_algorithm="sha256", show_progress=False,
            normalize_unicode=False,
        )
        assert list(scans[0].files) != list(scans[1].files)

    def test_case_insensitive_matching(self, tmp_path: Path):
        _write(tmp_path / "A" / "Report.DOCX", "same")
        _write(tmp_path / "B" / "report.docx", "same")
        scans = scan_locations(
            [("A", tmp_path / "A"), ("B", tmp_path / "B")],
            exclude_patterns=[], parallel_workers=2,
            hash_algorithm="sha256", show_progress=False,
            case_sensitive=False,
        )
        a, b = scans
        assert list(a.files) == list(b.files)
        key = list(a.files)[0]
        # 実ファイル名は拠点ごとの元の大小文字を保つ
        assert a.real_relpaths[key] == "Report.DOCX"
        assert b.real_relpaths[key] == "report.docx"
        assert a.files[key].hash == b.files[key].hash

    def test_case_sensitive_by_default(self, tmp_path: Path):
        _write(tmp_path / "A" / "Report.docx", "same")
        _write(tmp_path / "B" / "report.docx", "same")
        scans = scan_locations(
            [("A", tmp_path / "A"), ("B", tmp_path / "B")],
            exclude_patterns=[], parallel_workers=2,
            hash_algorithm="sha256", show_progress=False,
        )
        assert list(scans[0].files) != list(scans[1].files)

    def test_key_collision_within_location_is_reported(self, tmp_path: Path):
        """同一拠点内で照合キーが衝突したら、握り潰さずエラーに残す。

        大小文字を区別するファイルシステム (Linux の ext4 等) でのみ再現できる。
        macOS の APFS は既定で区別しないため 1 ファイルにまとまり、衝突が起きない。
        """
        root = tmp_path / "A"
        _write(root / "Report.docx", "one")
        _write(root / "report.docx", "two")
        if len(list(root.iterdir())) < 2:
            skip_or_fail("大小文字を区別しないファイルシステムでは再現できない")

        result = stat_location("A", root, exclude_patterns=[], case_sensitive=False)
        assert len(result.stats) == 1, "衝突したら先勝ちで1件だけ採用する"
        assert any("照合キーが重複" in e.message for e in result.errors)

    def test_key_collision_branch(self, tmp_path: Path, monkeypatch):
        """衝突分岐そのものの検証。

        大小文字を区別しないファイルシステム (macOS) では実ファイルで
        衝突を作れないため、ディレクトリ列挙を差し替えて確認する。
        """
        class FakeStat:
            st_size, st_mtime, st_mtime_ns = 10, 1000.0, 1000 * 10**9

        class FakeEntry:
            def __init__(self, name):
                self.name = name
                self.path = str(tmp_path / name)

            def is_dir(self):
                return False

            def is_symlink(self):
                return False

            def stat(self):
                return FakeStat()

        class FakeScandir:
            def __init__(self, entries):
                self.entries = entries

            def __enter__(self):
                return iter(self.entries)

            def __exit__(self, *a):
                return False

        monkeypatch.setattr(
            os, "scandir",
            lambda _p: FakeScandir([FakeEntry("Report.docx"), FakeEntry("report.docx")]),
        )
        result = stat_location("A", tmp_path, exclude_patterns=[], case_sensitive=False)
        assert list(result.real_relpaths.values()) == ["Report.docx"], "先勝ち"
        assert any("照合キーが重複" in e.message for e in result.errors)

    def test_hash_spans_multiple_chunks(self, tmp_path: Path):
        """チャンク境界 (1MiB) を跨ぐファイルでもハッシュが正しい。

        分割読みの積み上げを間違えても小さいファイルでは気付けない。
        """
        import hashlib

        payload = bytes(range(256)) * (HASH_CHUNK_SIZE * 2 // 256 + 500)
        assert len(payload) > HASH_CHUNK_SIZE * 2, "チャンクを跨いでいない"
        (tmp_path / "big.bin").write_bytes(payload)

        result = _scan(tmp_path)
        assert result.files["big.bin"].hash == hashlib.sha256(payload).hexdigest()
        assert result.files["big.bin"].size == len(payload)

class TestRetry:
    """読み取り失敗の再試行。

    ネットワーク共有では瞬断や一時的なロックですぐ直る失敗が起こる。
    1件でも失敗すると終了コード 3 (スキャン不完全) になるため、
    定期実行では実害のない失敗で警告が上がりがちだった。
    """

    def _flaky(self, monkeypatch, fail_times: int):
        calls = {"n": 0}

        def fake_hash(path, algorithm, cancel=None):
            calls["n"] += 1
            if calls["n"] <= fail_times:
                raise OSError(11, "Resource temporarily unavailable")
            return "recovered"

        monkeypatch.setattr(scanner, "_hash_file", fake_hash)
        return calls

    def test_transient_failure_recovers(self, monkeypatch, tmp_path: Path):
        calls = self._flaky(monkeypatch, fail_times=2)
        _, _, digest, err = _hash_task(
            0, "f.txt", tmp_path / "f.txt", "sha256", None,
            retry=3, retry_wait_sec=0.001,
        )
        assert digest == "recovered"
        assert err is None
        assert calls["n"] == 3

    def test_gives_up_after_the_configured_attempts(self, monkeypatch, tmp_path: Path):
        calls = self._flaky(monkeypatch, fail_times=99)
        _, _, digest, err = _hash_task(
            0, "f.txt", tmp_path / "f.txt", "sha256", None,
            retry=2, retry_wait_sec=0.001,
        )
        assert digest is None
        assert calls["n"] == 3            # 初回 + 再試行2回
        assert "3 回試行" in err          # 何回試したかを記録する

    def test_no_retry_by_default(self, monkeypatch, tmp_path: Path):
        """既定は現状維持 (1回で諦める)。"""
        calls = self._flaky(monkeypatch, fail_times=99)
        _, _, digest, err = _hash_task(0, "f.txt", tmp_path / "f.txt", "sha256")
        assert digest is None
        assert calls["n"] == 1
        assert "回試行" not in err        # 再試行していないので回数は書かない

    def test_cancel_during_retry_wait(self, monkeypatch, tmp_path: Path):
        """再試行の待ち時間中でも中断できる。"""
        self._flaky(monkeypatch, fail_times=99)
        cancel = threading.Event()
        cancel.set()
        with pytest.raises(ScanCancelled):
            _hash_task(
                0, "f.txt", tmp_path / "f.txt", "sha256", cancel,
                retry=3, retry_wait_sec=0.001,
            )

    def test_scan_locations_passes_retry_through(self, monkeypatch, tmp_path: Path):
        calls = self._flaky(monkeypatch, fail_times=1)
        _write(tmp_path / "A" / "f.txt", "x")
        _write(tmp_path / "B" / "f.txt", "x")
        monkeypatch.setattr(scanner, "DEFAULT_RETRY_WAIT_SEC", 0.001)
        scans = scan_locations(
            [("A", tmp_path / "A"), ("B", tmp_path / "B")],
            exclude_patterns=[], parallel_workers=1,
            hash_algorithm="sha256", show_progress=False, retry=2,
        )
        # 1回目の失敗を再試行で吸収し、エラーが残らない
        assert all(not s.file_errors for s in scans), [s.file_errors for s in scans]


class TestCancellation:
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

    def test_single_location_is_scanned_serially(self, tmp_path: Path):
        """拠点が1つならスレッドプールを立てずに直列で列挙する。"""
        _write(tmp_path / "A" / "f.txt", "x")
        seen = []
        scans = scan_locations(
            [("A", tmp_path / "A")],
            exclude_patterns=[], parallel_workers=2,
            hash_algorithm="sha256", show_progress=False,
            on_stat_done=lambda sr: seen.append(sr.location_name),
        )
        assert seen == ["A"]
        assert scans[0].files["f.txt"].hash is not None

    def test_progress_bar_path_produces_same_result(self, tmp_path: Path):
        """進捗バー表示あり (既定) でも結果は変わらない。"""
        for loc in ("A", "B"):
            for i in range(3):
                _write(tmp_path / loc / f"f{i}.txt", f"content-{i}")
        quiet = scan_locations(
            [("A", tmp_path / "A"), ("B", tmp_path / "B")],
            exclude_patterns=[], parallel_workers=2,
            hash_algorithm="sha256", show_progress=False,
        )
        loud = scan_locations(
            [("A", tmp_path / "A"), ("B", tmp_path / "B")],
            exclude_patterns=[], parallel_workers=2,
            hash_algorithm="sha256", show_progress=True,
        )
        for q, l in zip(quiet, loud):
            assert {k: v.hash for k, v in q.files.items()} == \
                   {k: v.hash for k, v in l.files.items()}

    def test_serial_hashing_with_progress(self, tmp_path: Path):
        """ワーカー1つ (直列) + 進捗バーの経路。"""
        _write(tmp_path / "A" / "a.txt", "x")
        _write(tmp_path / "A" / "b.txt", "y")
        _write(tmp_path / "B" / "a.txt", "x")
        _write(tmp_path / "B" / "b.txt", "y")
        scans = scan_locations(
            [("A", tmp_path / "A"), ("B", tmp_path / "B")],
            exclude_patterns=[], parallel_workers=1,
            hash_algorithm="sha256", show_progress=True,
        )
        assert scans[0].files["a.txt"].hash == scans[1].files["a.txt"].hash

    def test_walk_stops_on_cancel(self, tmp_path: Path):
        """列挙の途中でも中断フラグに反応する。"""
        _write(tmp_path / "A" / "sub" / "f.txt", "x")
        cancel = threading.Event()
        cancel.set()
        with pytest.raises(ScanCancelled):
            stat_location("A", tmp_path / "A", exclude_patterns=[], cancel=cancel)

    def test_unreadable_file_recorded_as_error(self, tmp_path: Path):
        if sys.platform == "win32":
            skip_or_fail("chmod-based unreadable test is POSIX only")
        if os.geteuid() == 0:  # type: ignore[attr-defined]
            skip_or_fail("running as root bypasses permission denial")
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
