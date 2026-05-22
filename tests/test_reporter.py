"""reporter: Excel/HTML 出力スモークテスト。

I/O ロジックなので「正しく生成され、想定のシート/セクション・色塗り・プレースホルダが入る」
ことを確認する。色やレイアウトの細部までは確認しない。
"""
from __future__ import annotations

import html as html_mod
from datetime import datetime
from pathlib import Path

import pytest

from comparator import (
    STATUS_ERROR,
    STATUS_HASH_MISMATCH,
    STATUS_OK,
    STATUS_PARTIAL_MISSING,
    STATUS_PARTIAL_PRESENT,
    compare,
)
from reporter import (
    ERROR_PLACEHOLDER,
    MISSING_PLACEHOLDER,
    ReportContext,
    _build_paths,
    _looks_windows,
    write_excel,
    write_html,
)

from .conftest import make_entry, make_scan


# ============================================================
# パス組み立てヘルパー
# ============================================================
class TestLooksWindows:
    def test_unc_forward_slash(self):
        assert _looks_windows("//server/share/docs") is True

    def test_unc_backslash(self):
        assert _looks_windows("\\\\server\\share\\docs") is True

    def test_drive_letter(self):
        assert _looks_windows("C:/Users/foo") is True
        assert _looks_windows("D:") is True

    def test_unix_path(self):
        assert _looks_windows("/mnt/share/docs") is False
        assert _looks_windows("/Users/foo") is False

    def test_relative_path(self):
        assert _looks_windows("relative/path") is False


class TestBuildPaths:
    def test_unc_forward_slash_normalizes_to_backslash(self):
        """設定が //server/share/docs でも Windows 形式に統一する。"""
        full, folder = _build_paths("//server-a/share/docs", "設計/詳細/spec.xlsx")
        assert full == "\\\\server-a\\share\\docs\\設計\\詳細\\spec.xlsx"
        assert folder == "\\\\server-a\\share\\docs\\設計\\詳細"

    def test_file_at_root_of_share(self):
        full, folder = _build_paths("//server/share", "file.txt")
        assert full == "\\\\server\\share\\file.txt"
        assert folder == "\\\\server\\share"

    def test_unix_root_keeps_forward_slash(self):
        full, folder = _build_paths("/mnt/data", "sub/file.txt")
        assert full == "/mnt/data/sub/file.txt"
        assert folder == "/mnt/data/sub"

    def test_drive_letter_root(self):
        full, folder = _build_paths("C:/Users/foo", "docs/file.txt")
        assert full == "C:\\Users\\foo\\docs\\file.txt"
        assert folder == "C:\\Users\\foo\\docs"

    def test_trailing_separator_stripped(self):
        """ルートの末尾セパレータがあっても二重にならない。"""
        full, _ = _build_paths("//server/share/", "file.txt")
        assert "\\\\\\" not in full  # 三重セパレータが現れない
        assert full == "\\\\server\\share\\file.txt"


# ============================================================
# 共通: 全状態を含むコンテキストを作る
# ============================================================
@pytest.fixture
def rich_ctx(tmp_path: Path) -> ReportContext:
    """OK / 不一致 / 欠落 / 一部のみ / エラー / dir差分 を1件ずつ含むコンテキスト。"""
    # 3拠点
    a = make_scan(
        "拠点A",
        files={
            "ok.txt": make_entry("hash_ok"),
            "mismatch.txt": make_entry("hash_majority"),    # A=B≠C
            "missing_in_C.txt": make_entry("hash_x"),
            "only_in_A.txt": make_entry("hash_y"),
        },
        dirs={"only_A_dir": object()},
    )
    b = make_scan(
        "拠点B",
        files={
            "ok.txt": make_entry("hash_ok"),
            "mismatch.txt": make_entry("hash_majority"),
            "missing_in_C.txt": make_entry("hash_x"),
        },
    )
    c = make_scan(
        "拠点C",
        files={
            "ok.txt": make_entry("hash_ok"),
            "mismatch.txt": make_entry("hash_minority"),    # ← 少数派
        },
        file_errors={"locked.txt": "PermissionError: [Errno 13]"},
    )
    comparison = compare([a, b, c])

    return ReportContext(
        started_at=datetime(2026, 5, 20, 12, 0, 0),
        finished_at=datetime(2026, 5, 20, 12, 0, 3),
        config_path=tmp_path / "fake.yaml",
        scans=[a, b, c],
        comparison=comparison,
    )


# ============================================================
# Excel
# ============================================================
class TestExcel:
    def test_writes_xlsx_with_expected_sheets(self, rich_ctx: ReportContext, tmp_path: Path):
        out = write_excel(rich_ctx, tmp_path / "out.xlsx")
        assert out.is_file()
        assert out.stat().st_size > 0

        from openpyxl import load_workbook
        wb = load_workbook(out)
        expected = {
            "サマリー",
            "全ファイル一覧",
            "ハッシュ不一致",
            "ファイル欠落",
            "余分なファイル",
            "エラー対象ファイル",
            "フォルダ構造差分",
            "エラー",
        }
        assert expected.issubset(set(wb.sheetnames)), wb.sheetnames

    def test_all_files_sheet_has_one_row_per_relpath(self, rich_ctx: ReportContext, tmp_path: Path):
        out = write_excel(rich_ctx, tmp_path / "out.xlsx")
        from openpyxl import load_workbook
        wb = load_workbook(out)
        ws = wb["全ファイル一覧"]
        # ヘッダ + 5ファイル (ok, mismatch, missing_in_C, only_in_A, locked)
        assert ws.max_row == 1 + 5

        # 列: No, relpath, name, ext, status, [A:size,A:hash,A:mtime, B:..., C:...]
        # = 5 + 3*3 = 14
        assert ws.max_column == 14

    def test_status_column_values(self, rich_ctx: ReportContext, tmp_path: Path):
        out = write_excel(rich_ctx, tmp_path / "out.xlsx")
        from openpyxl import load_workbook
        wb = load_workbook(out)
        ws = wb["全ファイル一覧"]
        statuses = {ws.cell(row=r, column=2).value: ws.cell(row=r, column=5).value
                    for r in range(2, ws.max_row + 1)}
        assert statuses["ok.txt"] == STATUS_OK
        assert statuses["mismatch.txt"] == STATUS_HASH_MISMATCH
        assert statuses["missing_in_C.txt"] == STATUS_PARTIAL_MISSING
        assert statuses["only_in_A.txt"] == STATUS_PARTIAL_PRESENT
        assert statuses["locked.txt"] == STATUS_ERROR

    def test_minority_hash_highlighting(self, rich_ctx: ReportContext, tmp_path: Path):
        """A=B≠C のとき C のハッシュ列だけ赤背景になる (Bug A 修正)。"""
        out = write_excel(rich_ctx, tmp_path / "out.xlsx")
        from openpyxl import load_workbook
        wb = load_workbook(out)
        ws = wb["全ファイル一覧"]
        # mismatch.txt の行を探す
        for r in range(2, ws.max_row + 1):
            if ws.cell(row=r, column=2).value == "mismatch.txt":
                # 拠点A:hash=列7, 拠点B:hash=列10, 拠点C:hash=列13
                a_fill = ws.cell(row=r, column=7).fill.fgColor.rgb
                b_fill = ws.cell(row=r, column=10).fill.fgColor.rgb
                c_fill = ws.cell(row=r, column=13).fill.fgColor.rgb
                # 多数派 (A, B) は塗りなし
                assert a_fill in ("00000000", None)
                assert b_fill in ("00000000", None)
                # 少数派 (C) は赤 FILL_HASH_MISMATCH = FFC7CE
                assert c_fill == "00FFC7CE", f"expected red, got {c_fill}"
                return
        pytest.fail("mismatch.txt row not found")

    def test_missing_cells_use_half_width_placeholder(
        self, rich_ctx: ReportContext, tmp_path: Path
    ):
        """欠落セルは半角 '-' (項目 P 修正)。"""
        out = write_excel(rich_ctx, tmp_path / "out.xlsx")
        from openpyxl import load_workbook
        wb = load_workbook(out)
        ws = wb["全ファイル一覧"]
        for r in range(2, ws.max_row + 1):
            if ws.cell(row=r, column=2).value == "only_in_A.txt":
                # 拠点B, C は欠落 → 列 9,10,11 と 12,13,14
                for c in (9, 10, 11, 12, 13, 14):
                    assert ws.cell(row=r, column=c).value == MISSING_PLACEHOLDER
                assert MISSING_PLACEHOLDER == "-"
                return
        pytest.fail("only_in_A.txt row not found")

    def test_error_row_renders_error_placeholder_and_message(
        self, rich_ctx: ReportContext, tmp_path: Path
    ):
        """エラー対象ファイル: サイズ列に 'エラー'、ハッシュ列にメッセージ。"""
        out = write_excel(rich_ctx, tmp_path / "out.xlsx")
        from openpyxl import load_workbook
        wb = load_workbook(out)
        ws = wb["全ファイル一覧"]
        for r in range(2, ws.max_row + 1):
            if ws.cell(row=r, column=2).value == "locked.txt":
                # 拠点C: 列12=size, 13=hash, 14=mtime
                assert ws.cell(row=r, column=12).value == ERROR_PLACEHOLDER
                assert "PermissionError" in str(ws.cell(row=r, column=13).value)
                return
        pytest.fail("locked.txt row not found")

    def test_subset_sheets_only_contain_matching_rows(
        self, rich_ctx: ReportContext, tmp_path: Path
    ):
        out = write_excel(rich_ctx, tmp_path / "out.xlsx")
        from openpyxl import load_workbook
        wb = load_workbook(out)
        # ハッシュ不一致シートには mismatch.txt のみ
        mismatch_sheet = wb["ハッシュ不一致"]
        rels = [mismatch_sheet.cell(row=r, column=2).value
                for r in range(2, mismatch_sheet.max_row + 1)]
        assert rels == ["mismatch.txt"]

        # ファイル欠落: missing_in_C.txt のみ
        missing_sheet = wb["ファイル欠落"]
        rels = [missing_sheet.cell(row=r, column=2).value
                for r in range(2, missing_sheet.max_row + 1)]
        assert rels == ["missing_in_C.txt"]

        # 余分なファイル: only_in_A.txt のみ
        extra_sheet = wb["余分なファイル"]
        rels = [extra_sheet.cell(row=r, column=2).value
                for r in range(2, extra_sheet.max_row + 1)]
        assert rels == ["only_in_A.txt"]

        # エラー対象ファイル: locked.txt のみ
        errored_sheet = wb["エラー対象ファイル"]
        rels = [errored_sheet.cell(row=r, column=2).value
                for r in range(2, errored_sheet.max_row + 1)]
        assert rels == ["locked.txt"]

    def test_empty_subset_shows_placeholder(self, tmp_path: Path):
        """差分が無いケースでは '該当なし' が入る。"""
        a = make_scan("A", files={"x.txt": make_entry("h")})
        b = make_scan("B", files={"x.txt": make_entry("h")})
        comparison = compare([a, b])
        ctx = ReportContext(
            started_at=datetime(2026, 1, 1),
            finished_at=datetime(2026, 1, 1),
            config_path=tmp_path / "c.yaml",
            scans=[a, b],
            comparison=comparison,
        )
        out = write_excel(ctx, tmp_path / "empty.xlsx")
        from openpyxl import load_workbook
        wb = load_workbook(out)
        for sheet in ("ハッシュ不一致", "ファイル欠落", "余分なファイル", "エラー対象ファイル"):
            ws = wb[sheet]
            assert ws.cell(row=2, column=1).value == "該当なし", sheet

    def test_summary_sheet_includes_diff_counts(
        self, rich_ctx: ReportContext, tmp_path: Path
    ):
        out = write_excel(rich_ctx, tmp_path / "out.xlsx")
        from openpyxl import load_workbook
        wb = load_workbook(out)
        ws = wb["サマリー"]
        # 全セル走査でラベル+件数を回収
        all_cells = {}
        for row in ws.iter_rows(values_only=True):
            if len(row) >= 2 and row[0] is not None:
                all_cells[row[0]] = row[1]
        assert all_cells.get("ハッシュ不一致") == 1
        assert all_cells.get("ファイル欠落 (一部拠点になし)") == 1
        assert all_cells.get("余分なファイル (一部拠点のみ)") == 1
        assert all_cells.get("エラー対象ファイル (要確認)") == 1


# ============================================================
# HTML
# ============================================================
class TestHtml:
    def test_writes_valid_html(self, rich_ctx: ReportContext, tmp_path: Path):
        out = write_html(rich_ctx, tmp_path / "out.html")
        assert out.is_file()
        body = out.read_text(encoding="utf-8")
        assert body.startswith("\n<!DOCTYPE html>") or body.lstrip().startswith("<!DOCTYPE html>")
        assert "</html>" in body

    def test_contains_all_sections(self, rich_ctx: ReportContext, tmp_path: Path):
        out = write_html(rich_ctx, tmp_path / "out.html")
        body = out.read_text(encoding="utf-8")
        for sid in ("summary", "all", "mismatch", "missing", "extra", "errored", "dirs", "errors"):
            assert f'id="{sid}"' in body, f"missing section: {sid}"

    def test_status_classes_applied(self, rich_ctx: ReportContext, tmp_path: Path):
        out = write_html(rich_ctx, tmp_path / "out.html")
        body = out.read_text(encoding="utf-8")
        for cls in (
            "status-ok",
            "status-mismatch",
            "status-missing",
            "status-partial",
            "status-error",
        ):
            assert cls in body, f"missing class: {cls}"

    def test_minority_hash_uses_cell_mismatch_class(
        self, rich_ctx: ReportContext, tmp_path: Path
    ):
        """A=B≠C のとき少数派ハッシュだけが cell-mismatch クラスを持つはず。"""
        out = write_html(rich_ctx, tmp_path / "out.html")
        body = out.read_text(encoding="utf-8")
        # ハッシュ値は最初の12文字で短縮表示される
        # hash_majority → "hash_majorit…", hash_minority → "hash_minorit…"
        # cell-mismatch クラスが付くのは hash_minority 側のみ
        # 簡易チェック: cell-mismatch がページに存在
        assert "cell-mismatch" in body

    def test_missing_uses_half_width_placeholder(
        self, rich_ctx: ReportContext, tmp_path: Path
    ):
        out = write_html(rich_ctx, tmp_path / "out.html")
        body = out.read_text(encoding="utf-8")
        # 全角 '－' は使わない
        assert "－" not in body
        # 半角 '-' の欠落セルが存在
        assert "cell-missing'>-" in body

    def test_error_cell_shows_message_in_tooltip(
        self, rich_ctx: ReportContext, tmp_path: Path
    ):
        out = write_html(rich_ctx, tmp_path / "out.html")
        body = out.read_text(encoding="utf-8")
        # エラーメッセージが title 属性に含まれる
        assert "PermissionError" in body
        assert "cell-error" in body

    def test_sticky_column_has_explicit_background(
        self, rich_ctx: ReportContext, tmp_path: Path
    ):
        """Issue 1: 横スクロール時に sticky 列が透明になる問題が再発しないこと。

        - background: inherit を使っていない
        - thead 側に明示的な背景色
        - tbody の odd/even 両方に明示的な背景色
        """
        out = write_html(rich_ctx, tmp_path / "out.html")
        body = out.read_text(encoding="utf-8")
        assert "background: inherit" not in body, (
            "sticky 列に background: inherit を使うと透明になる。明示的な色を指定すること"
        )
        # ファイル名列 (3列目) を sticky 対象とする
        assert ".fixed-col-table thead th:nth-child(3)" in body
        # tbody 側の odd/even
        assert ".fixed-col-table tbody tr:nth-child(odd) td:nth-child(3)" in body
        assert ".fixed-col-table tbody tr:nth-child(even) td:nth-child(3)" in body
        # 旧仕様 (相対パス列 = 2 列目) が残っていないこと
        assert ".fixed-col-table tbody td:nth-child(2)" not in body

    def test_page_header_is_not_sticky(
        self, rich_ctx: ReportContext, tmp_path: Path
    ):
        """Issue 2: ページヘッダーを sticky にするとテーブル列ヘッダーと衝突するため、
        ページヘッダー側を解除した状態を維持する。"""
        out = write_html(rich_ctx, tmp_path / "out.html")
        body = out.read_text(encoding="utf-8")
        # 旧バグの組み合わせを直接禁止
        assert "header { background: #305496; color: #fff; padding: 16px 24px; position: sticky" not in body

    def test_section_has_scroll_margin_top(
        self, rich_ctx: ReportContext, tmp_path: Path
    ):
        """Issue 3: アンカージャンプ時、section 見出しが画面端に潜らないように余白がある。"""
        out = write_html(rich_ctx, tmp_path / "out.html")
        body = out.read_text(encoding="utf-8")
        assert "scroll-margin-top" in body, "section に scroll-margin-top を設定すること"

    def test_file_row_has_required_data_attributes(
        self, rich_ctx: ReportContext, tmp_path: Path
    ):
        """行クリックでモーダルを開くために data-relpath と data-detail-html が付与される。"""
        out = write_html(rich_ctx, tmp_path / "out.html")
        body = out.read_text(encoding="utf-8")
        assert 'data-relpath="mismatch.txt"' in body
        assert 'data-relpath="ok.txt"' in body
        # 詳細 HTML がエスケープされて埋め込まれている
        assert 'data-detail-html="' in body

    def test_no_inline_detail_row_rendered(
        self, rich_ctx: ReportContext, tmp_path: Path
    ):
        """詳細はモーダル経由になったため、インラインの detail-row は出力されない。"""
        out = write_html(rich_ctx, tmp_path / "out.html")
        body = out.read_text(encoding="utf-8")
        assert "detail-row" not in body

    def test_no_section_expand_controls(
        self, rich_ctx: ReportContext, tmp_path: Path
    ):
        """モーダル化により 全展開/全折りたたみ ボタンは削除済み。"""
        out = write_html(rich_ctx, tmp_path / "out.html")
        body = out.read_text(encoding="utf-8")
        assert "data-action=" not in body
        assert "expand-all" not in body
        assert "collapse-all" not in body

    def test_detail_html_contains_copy_paths(
        self, rich_ctx: ReportContext, tmp_path: Path
    ):
        """data-detail-html の中身を unescape した状態で各拠点のコピー用パスが含まれる。

        rich_ctx の拠点ルートは make_scan のデフォルト '/tmp/dummy' のため UNIX 形式。
        relpath='mismatch.txt' → full='/tmp/dummy/mismatch.txt', folder='/tmp/dummy'
        """
        out = write_html(rich_ctx, tmp_path / "out.html")
        body = out.read_text(encoding="utf-8")
        # data-detail-html 属性内の HTML は属性向けエスケープが掛かっているため
        # 一度デコードしてから検査する
        decoded = html_mod.unescape(body)
        assert 'data-copy="/tmp/dummy/mismatch.txt"' in decoded  # ファイル
        assert 'data-copy="/tmp/dummy"' in decoded  # フォルダ

    def test_detail_path_table_omits_folder_column(
        self, rich_ctx: ReportContext, tmp_path: Path
    ):
        """フォルダパス列は廃止 (ボタンには残るが列としては表示しない)。"""
        out = write_html(rich_ctx, tmp_path / "out.html")
        decoded = html_mod.unescape(out.read_text(encoding="utf-8"))
        assert "<th>フォルダパス</th>" not in decoded
        assert "<th>ファイルパス</th>" in decoded

    def test_modal_dialog_present(
        self, rich_ctx: ReportContext, tmp_path: Path
    ):
        """ページ末尾に <dialog id="detail-modal"> が1つ存在。"""
        out = write_html(rich_ctx, tmp_path / "out.html")
        body = out.read_text(encoding="utf-8")
        assert '<dialog id="detail-modal"' in body
        assert 'id="detail-modal-title"' in body
        assert 'id="detail-modal-content"' in body
        # close ボタン
        assert 'class="modal-close"' in body

    def test_modal_script_present(
        self, rich_ctx: ReportContext, tmp_path: Path
    ):
        """showModal / close の呼び出しを含む JS が埋め込まれている。"""
        out = write_html(rich_ctx, tmp_path / "out.html")
        body = out.read_text(encoding="utf-8")
        assert "showModal" in body
        assert "modal.close" in body

    def test_clipboard_script_with_fallback_present(
        self, rich_ctx: ReportContext, tmp_path: Path
    ):
        """新 API と旧 API フォールバック両方の呼び出しが script に含まれる。"""
        out = write_html(rich_ctx, tmp_path / "out.html")
        body = out.read_text(encoding="utf-8")
        # 新 API
        assert "navigator.clipboard.writeText" in body
        # 旧 API フォールバック
        assert "document.execCommand" in body

    def test_html_escapes_special_chars_in_paths(self, tmp_path: Path):
        """パス名に <script> を混ぜても素通りしないこと (HTMLエスケープ)。"""
        a = make_scan("A", files={"<script>.txt": make_entry("h")})
        b = make_scan("B", files={"<script>.txt": make_entry("h")})
        comparison = compare([a, b])
        ctx = ReportContext(
            started_at=datetime(2026, 1, 1),
            finished_at=datetime(2026, 1, 1),
            config_path=tmp_path / "c.yaml",
            scans=[a, b],
            comparison=comparison,
        )
        out = write_html(ctx, tmp_path / "evil.html")
        body = out.read_text(encoding="utf-8")
        # ファイル名由来の <script>.txt は素のまま出ない (エスケープされている)
        assert "<script>.txt" not in body
        assert "&lt;script&gt;.txt" in body
