"""comparator: 状態判定とハッシュ最頻値ロジックの境界テスト。"""
from __future__ import annotations

import pytest

from comparator import (
    STATUS_ERROR,
    STATUS_HASH_MISMATCH,
    STATUS_OK,
    STATUS_PARTIAL_MISSING,
    STATUS_PARTIAL_PRESENT,
    _classify,
    compare,
    minority_hashes,
    minority_sizes,
)

from .conftest import make_entry, make_scan


# ==========================================
# minority_hashes
# ==========================================
class TestMinorityHashes:
    def test_all_same_returns_empty(self):
        entries = {"A": make_entry("h"), "B": make_entry("h"), "C": make_entry("h")}
        assert minority_hashes(entries) == set()

    def test_clear_majority_returns_minority(self):
        # A=B=h1, C=h2  → C のみ少数派
        entries = {"A": make_entry("h1"), "B": make_entry("h1"), "C": make_entry("h2")}
        assert minority_hashes(entries) == {"h2"}

    def test_two_way_tie_highlights_both(self):
        # A=h1, B=h2 (1:1) → どちらが正か判断不能 → 両方
        entries = {"A": make_entry("h1"), "B": make_entry("h2")}
        assert minority_hashes(entries) == {"h1", "h2"}

    def test_all_different_highlights_all(self):
        entries = {"A": make_entry("h1"), "B": make_entry("h2"), "C": make_entry("h3")}
        assert minority_hashes(entries) == {"h1", "h2", "h3"}

    def test_balanced_two_groups_highlights_all(self):
        # 4拠点で 2:2 のタイ → 全強調
        entries = {
            "A": make_entry("h1"), "B": make_entry("h1"),
            "C": make_entry("h2"), "D": make_entry("h2"),
        }
        assert minority_hashes(entries) == {"h1", "h2"}

    def test_three_groups_with_clear_majority(self):
        # 5拠点で 3:1:1 → minority は2つの少数派ハッシュ
        entries = {
            "A": make_entry("h1"), "B": make_entry("h1"), "C": make_entry("h1"),
            "D": make_entry("h2"),
            "E": make_entry("h3"),
        }
        assert minority_hashes(entries) == {"h2", "h3"}

    def test_single_entry_returns_empty(self):
        entries = {"A": make_entry("h1")}
        assert minority_hashes(entries) == set()

    def test_all_none_returns_empty(self):
        entries = {"A": None, "B": None}
        assert minority_hashes(entries) == set()

    def test_none_entries_ignored(self):
        # 存在しない拠点は無視される。残り 1個なので minority は空。
        entries = {"A": make_entry("h1"), "B": None, "C": None}
        assert minority_hashes(entries) == set()


# ==========================================
# _classify
# ==========================================
class TestClassify:
    def _no_errors(self, keys):
        return {k: None for k in keys}

    def test_all_present_same_hash_is_ok(self):
        entries = {"A": make_entry("h"), "B": make_entry("h"), "C": make_entry("h")}
        assert _classify(entries, self._no_errors(entries)) == STATUS_OK

    def test_all_present_different_hash_is_mismatch(self):
        entries = {"A": make_entry("h1"), "B": make_entry("h2"), "C": make_entry("h1")}
        assert _classify(entries, self._no_errors(entries)) == STATUS_HASH_MISMATCH

    def test_3loc_2present_is_partial_missing(self):
        # 3拠点で 2:1 → 過半数に存在 → 欠落あり
        entries = {"A": make_entry("h"), "B": make_entry("h"), "C": None}
        assert _classify(entries, self._no_errors(entries)) == STATUS_PARTIAL_MISSING

    def test_3loc_1present_is_partial_present(self):
        # 3拠点で 1:2 → 少数派 → 一部のみ存在
        entries = {"A": make_entry("h"), "B": None, "C": None}
        assert _classify(entries, self._no_errors(entries)) == STATUS_PARTIAL_PRESENT

    def test_4loc_2present_is_partial_present_edge(self):
        # 4拠点で 2:2 (n_present > n_total/2 ではない) → 一部のみ存在
        entries = {
            "A": make_entry("h"), "B": make_entry("h"),
            "C": None, "D": None,
        }
        assert _classify(entries, self._no_errors(entries)) == STATUS_PARTIAL_PRESENT

    def test_4loc_3present_is_partial_missing(self):
        entries = {
            "A": make_entry("h"), "B": make_entry("h"), "C": make_entry("h"),
            "D": None,
        }
        assert _classify(entries, self._no_errors(entries)) == STATUS_PARTIAL_MISSING

    def test_any_error_returns_error_status(self):
        # エラーが1件でもあれば、欠落/不一致より優先
        entries = {"A": make_entry("h"), "B": make_entry("h"), "C": None}
        errors = {"A": None, "B": None, "C": "Permission denied"}
        assert _classify(entries, errors) == STATUS_ERROR

    def test_error_overrides_mismatch(self):
        entries = {"A": make_entry("h1"), "B": make_entry("h2"), "C": None}
        errors = {"A": None, "B": None, "C": "Read failed"}
        assert _classify(entries, errors) == STATUS_ERROR


class TestDisplayRelpath:
    def test_falls_back_to_match_key_when_no_real_path(self):
        """実パスが分からないファイルは照合キーをそのまま表示する。

        列挙フェーズで stat に失敗したファイル (リンク切れ等) は
        file_errors には載るが real_relpaths には載らないため、
        このフォールバックが実際に使われる。
        """
        a = make_scan("拠点A", files={"ok.txt": make_entry("h")})
        b = make_scan(
            "拠点B",
            files={"ok.txt": make_entry("h")},
            file_errors={"dangling.txt": "FileNotFoundError: [Errno 2]"},
            real_relpaths={"ok.txt": "ok.txt"},  # dangling.txt は実パス不明
        )
        result = compare([a, b])
        rels = [r.relpath for r in result.all_files]
        assert "dangling.txt" in rels
        row = next(r for r in result.all_files if r.relpath == "dangling.txt")
        assert row.status == STATUS_ERROR
        assert row.real_relpaths == {}

    def test_uses_first_location_that_has_the_file(self):
        """設定順で最初に見つかった拠点の実際の名前を表示に使う。"""
        a = make_scan("拠点A", files={})
        b = make_scan(
            "拠点B",
            files={"report.docx": make_entry("h")},
            real_relpaths={"report.docx": "Report.DOCX"},
        )
        result = compare([a, b])
        assert [r.relpath for r in result.all_files] == ["Report.DOCX"]


class TestClassifyWithoutHashes:
    """hash_mode=smart ではハッシュ未計算 (hash=None) の行が出る。"""

    def _no_errors(self, keys):
        return {k: None for k in keys}

    def test_size_difference_is_mismatch_without_hash(self):
        """サイズが違えばハッシュが無くても不一致と判定できる。"""
        entries = {
            "A": make_entry(None, size=100),
            "B": make_entry(None, size=200),
        }
        assert _classify(entries, self._no_errors(entries)) == STATUS_HASH_MISMATCH

    def test_same_size_no_hash_is_ok(self):
        """サイズ・更新日時一致でハッシュを省略した行は一致扱い。"""
        entries = {
            "A": make_entry(None, size=100),
            "B": make_entry(None, size=100),
            "C": make_entry(None, size=100),
        }
        assert _classify(entries, self._no_errors(entries)) == STATUS_OK

    def test_size_difference_wins_over_equal_hashes(self):
        """サイズ判定はハッシュより先。両立しない入力は不一致に倒す。"""
        entries = {
            "A": make_entry("h", size=100),
            "B": make_entry("h", size=200),
        }
        assert _classify(entries, self._no_errors(entries)) == STATUS_HASH_MISMATCH

    def test_minority_sizes_highlights_odd_location(self):
        entries = {
            "A": make_entry(None, size=100),
            "B": make_entry(None, size=100),
            "C": make_entry(None, size=999),
        }
        assert minority_sizes(entries) == {999}

    def test_no_location_has_the_file_is_treated_as_missing(self):
        """どの拠点にも実体が無くエラーも無い入力への防御。

        compare 経由では到達しない (和集合に載る時点でどこかに実体かエラーがある)
        が、_classify を単体で使ったときに黙って誤分類しないことを固定しておく。
        """
        entries = {"A": None, "B": None}
        assert _classify(entries, self._no_errors(entries)) == STATUS_PARTIAL_MISSING

    def test_minority_hashes_ignores_uncomputed(self):
        """ハッシュ未計算の行では強調対象なし (サイズ側で示す)。"""
        entries = {
            "A": make_entry(None, size=100),
            "B": make_entry(None, size=200),
        }
        assert minority_hashes(entries) == set()


# ==========================================
# compare
# ==========================================
class TestCompare:
    def test_requires_at_least_two_scans(self):
        with pytest.raises(ValueError):
            compare([])
        with pytest.raises(ValueError):
            compare([make_scan("A")])

    def test_two_identical_scans_have_no_diffs(self):
        files = {"a.txt": make_entry("h"), "b.txt": make_entry("h2")}
        scans = [make_scan("A", files=files), make_scan("B", files=files)]
        result = compare(scans)
        assert result.hash_mismatches == []
        assert result.missing_files == []
        assert result.extra_files == []
        assert result.errored_files == []
        assert len(result.all_files) == 2
        assert all(row.status == STATUS_OK for row in result.all_files)

    def test_three_way_categorization(self):
        a_files = {
            "ok.txt": make_entry("same"),
            "mismatch.txt": make_entry("v1"),
            "missing_in_C.txt": make_entry("x"),
            "only_in_A.txt": make_entry("y"),
        }
        b_files = {
            "ok.txt": make_entry("same"),
            "mismatch.txt": make_entry("v1"),
            "missing_in_C.txt": make_entry("x"),
        }
        c_files = {
            "ok.txt": make_entry("same"),
            "mismatch.txt": make_entry("v2"),  # 1拠点だけ違う
        }
        scans = [
            make_scan("A", files=a_files),
            make_scan("B", files=b_files),
            make_scan("C", files=c_files),
        ]
        result = compare(scans)

        rel_status = {r.relpath: r.status for r in result.all_files}
        assert rel_status["ok.txt"] == STATUS_OK
        assert rel_status["mismatch.txt"] == STATUS_HASH_MISMATCH
        assert rel_status["missing_in_C.txt"] == STATUS_PARTIAL_MISSING
        assert rel_status["only_in_A.txt"] == STATUS_PARTIAL_PRESENT

        # 各カテゴリの件数
        assert len(result.hash_mismatches) == 1
        assert len(result.missing_files) == 1
        assert len(result.extra_files) == 1
        assert len(result.errored_files) == 0

    def test_errored_file_separated_from_missing(self):
        """Bug fix B: A だけ読取失敗、B/C 正常 → 欠落ではなくエラー。"""
        files_ok = {"data.txt": make_entry("h")}
        scans = [
            make_scan("A", files={}, file_errors={"data.txt": "Permission denied"}),
            make_scan("B", files=files_ok),
            make_scan("C", files=files_ok),
        ]
        result = compare(scans)
        assert len(result.errored_files) == 1
        assert len(result.missing_files) == 0
        assert len(result.extra_files) == 0
        assert result.errored_files[0].relpath == "data.txt"
        assert result.errored_files[0].status == STATUS_ERROR
        assert result.errored_files[0].errors["A"] == "Permission denied"
        assert result.errored_files[0].errors["B"] is None

    def test_all_files_includes_error_only_relpaths(self):
        """ファイルが見つかった拠点が無くてもエラー記録があれば all_files に出る。"""
        scans = [
            make_scan("A", files={}, file_errors={"lost.txt": "I/O error"}),
            make_scan("B", files={}),
        ]
        result = compare(scans)
        assert len(result.all_files) == 1
        assert result.all_files[0].relpath == "lost.txt"
        assert result.all_files[0].status == STATUS_ERROR

    def test_dir_diff_detects_missing_dir(self):
        from scanner import DirEntry
        scans = [
            make_scan("A", dirs={"only_in_A": DirEntry(), "shared": DirEntry()}),
            make_scan("B", dirs={"shared": DirEntry()}),
        ]
        result = compare(scans)
        rels = {d.relpath: d.presence for d in result.dir_diffs}
        assert "only_in_A" in rels
        assert "shared" not in rels
        assert rels["only_in_A"] == {"A": True, "B": False}

    def test_sorted_output_by_relpath(self):
        files_a = {"z.txt": make_entry("h"), "a.txt": make_entry("h")}
        files_b = {"z.txt": make_entry("h"), "a.txt": make_entry("h")}
        scans = [make_scan("A", files=files_a), make_scan("B", files=files_b)]
        result = compare(scans)
        assert [r.relpath for r in result.all_files] == ["a.txt", "z.txt"]
