"""レポート生成 (Excel / HTML)。"""
from __future__ import annotations

import html
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook as WB

from comparator import (
    STATUS_ERROR,
    STATUS_HASH_MISMATCH,
    STATUS_OK,
    STATUS_PARTIAL_MISSING,
    STATUS_PARTIAL_PRESENT,
    ComparisonResult,
    FileRow,
    minority_hashes,
)
from scanner import FileEntry, ScanResult
from utils import human_bytes


MISSING_PLACEHOLDER = "-"
ERROR_PLACEHOLDER = "エラー"
DATETIME_FMT = "%Y-%m-%d %H:%M:%S"


def _looks_windows(root_str: str) -> bool:
    """ルートが Windows UNC (`//` or `\\\\`) またはドライブレターか判定。"""
    return (
        root_str.startswith(("//", "\\\\"))
        or (len(root_str) >= 2 and root_str[1] == ":")
    )


def _build_paths(root_str: str, relpath: str) -> tuple[str, str]:
    """拠点のルートと相対パスからフルパスとフォルダパスを構築する。

    Windows ルート判定なら全セパレータをバックスラッシュに統一し、
    Windows エクスプローラのアドレス欄に貼り付けて開ける形式にする。

    Returns:
        (full_file_path, folder_path)
    """
    if _looks_windows(root_str):
        sep = "\\"
        norm_root = root_str.replace("/", "\\").rstrip("\\")
        rel = relpath.replace("/", "\\")
    else:
        sep = "/"
        norm_root = root_str.rstrip("/")
        rel = relpath
    full = f"{norm_root}{sep}{rel}"
    folder = full.rsplit(sep, 1)[0] if sep in full else full
    return full, folder


# === 配色 ===
FILL_OK = PatternFill("solid", fgColor="C6EFCE")           # 緑
FILL_HASH_MISMATCH = PatternFill("solid", fgColor="FFC7CE") # 赤
FILL_MISSING = PatternFill("solid", fgColor="FFD8A8")      # オレンジ
FILL_PARTIAL = PatternFill("solid", fgColor="FFE699")      # 黄
FILL_ERROR = PatternFill("solid", fgColor="F4B084")        # 濃いオレンジ (読み取り失敗)
FILL_GRAY = PatternFill("solid", fgColor="D9D9D9")
FILL_HEADER = PatternFill("solid", fgColor="305496")
FONT_HEADER = Font(bold=True, color="FFFFFF")


@dataclass
class ReportContext:
    started_at: datetime
    finished_at: datetime
    config_path: Path
    scans: List[ScanResult]
    comparison: ComparisonResult


# ============================================================
# Excel
# ============================================================
def write_excel(ctx: ReportContext, out_path: Path) -> Path:
    wb: WB = Workbook()
    # デフォルトで作成される空シートを削除
    default_ws = wb.active
    wb.remove(default_ws)

    _excel_summary(wb, ctx)
    _excel_all_files(wb, ctx)
    _excel_subset(wb, ctx, "ハッシュ不一致", ctx.comparison.hash_mismatches)
    _excel_subset(wb, ctx, "ファイル欠落", ctx.comparison.missing_files)
    _excel_subset(wb, ctx, "余分なファイル", ctx.comparison.extra_files)
    _excel_subset(wb, ctx, "エラー対象ファイル", ctx.comparison.errored_files)
    _excel_dir_diff(wb, ctx)
    _excel_errors(wb, ctx)

    wb.save(out_path)
    return out_path


def _set_header_row(ws, headers: List[str]) -> None:
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.fill = FILL_HEADER
        cell.font = FONT_HEADER
        cell.alignment = Alignment(horizontal="center", vertical="center")


def _excel_summary(wb: WB, ctx: ReportContext) -> None:
    ws = wb.create_sheet("サマリー")
    elapsed = (ctx.finished_at - ctx.started_at).total_seconds()

    rows: List[List[object]] = [
        ["項目", "値"],
        ["スキャン開始", ctx.started_at.strftime(DATETIME_FMT)],
        ["スキャン終了", ctx.finished_at.strftime(DATETIME_FMT)],
        ["所要時間 (秒)", f"{elapsed:.2f}"],
        ["設定ファイル", str(ctx.config_path)],
        ["拠点数", len(ctx.scans)],
        [],
        ["拠点", "ルートパス", "ファイル数", "総サイズ", "エラー数"],
    ]
    for s in ctx.scans:
        total_size = sum(f.size for f in s.files.values())
        rows.append([
            s.location_name,
            str(s.root),
            len(s.files),
            human_bytes(total_size),
            len(s.errors),
        ])

    rows.append([])
    rows.append(["差分種別", "件数"])
    rows.append(["ハッシュ不一致", len(ctx.comparison.hash_mismatches)])
    rows.append(["ファイル欠落 (一部拠点になし)", len(ctx.comparison.missing_files)])
    rows.append(["余分なファイル (一部拠点のみ)", len(ctx.comparison.extra_files)])
    rows.append(["エラー対象ファイル (要確認)", len(ctx.comparison.errored_files)])
    rows.append(["フォルダ構造差分", len(ctx.comparison.dir_diffs)])

    for r_idx, row in enumerate(rows, 1):
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            if r_idx == 1 or (isinstance(val, str) and val in {"拠点", "差分種別", "項目"}):
                # ヘッダ行の見た目
                if c_idx <= len(row):
                    cell.font = Font(bold=True)

    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 60
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 10


def _file_columns_for(location_names: List[str]) -> List[str]:
    cols = ["No.", "相対パス", "ファイル名", "拡張子", "状態"]
    for name in location_names:
        cols.extend([f"{name}: サイズ", f"{name}: ハッシュ", f"{name}: 更新日時"])
    return cols


def _status_fill(status: str) -> Optional[PatternFill]:
    return {
        STATUS_OK: FILL_OK,
        STATUS_HASH_MISMATCH: FILL_HASH_MISMATCH,
        STATUS_PARTIAL_MISSING: FILL_MISSING,
        STATUS_PARTIAL_PRESENT: FILL_PARTIAL,
        STATUS_ERROR: FILL_ERROR,
    }.get(status)


def _write_file_row(
    ws,
    row_idx: int,
    no: int,
    file_row: FileRow,
    location_names: List[str],
) -> None:
    rel = file_row.relpath
    name = rel.rsplit("/", 1)[-1]
    ext = Path(name).suffix

    ws.cell(row=row_idx, column=1, value=no)
    ws.cell(row=row_idx, column=2, value=rel)
    ws.cell(row=row_idx, column=3, value=name)
    ws.cell(row=row_idx, column=4, value=ext)

    status_cell = ws.cell(row=row_idx, column=5, value=file_row.status)
    fill = _status_fill(file_row.status)
    if fill is not None:
        status_cell.fill = fill

    # ハッシュ不一致時は「最頻値（過半数）と異なる拠点だけ」赤くする。
    minority = (
        minority_hashes(file_row.entries) if file_row.status == STATUS_HASH_MISMATCH else set()
    )

    col = 6
    for loc in location_names:
        err_msg = file_row.errors.get(loc)
        entry: Optional[FileEntry] = file_row.entries.get(loc)
        if err_msg is not None:
            # 読み取り失敗: サイズ列=エラー、ハッシュ列=メッセージ、更新日時列=―
            size_cell = ws.cell(row=row_idx, column=col, value=ERROR_PLACEHOLDER)
            msg_cell = ws.cell(row=row_idx, column=col + 1, value=err_msg)
            mtime_cell = ws.cell(row=row_idx, column=col + 2, value=MISSING_PLACEHOLDER)
            for c in (size_cell, msg_cell, mtime_cell):
                c.fill = FILL_ERROR
            size_cell.alignment = Alignment(horizontal="center")
            mtime_cell.alignment = Alignment(horizontal="center")
        elif entry is None:
            for offset in range(3):
                c = ws.cell(row=row_idx, column=col + offset, value=MISSING_PLACEHOLDER)
                c.fill = FILL_GRAY
                c.alignment = Alignment(horizontal="center")
        else:
            ws.cell(row=row_idx, column=col, value=entry.size)
            hash_cell = ws.cell(row=row_idx, column=col + 1, value=entry.hash)
            ws.cell(row=row_idx, column=col + 2, value=entry.mtime.strftime(DATETIME_FMT))
            if entry.hash in minority:
                hash_cell.fill = FILL_HASH_MISMATCH
        col += 3


def _excel_all_files(wb: WB, ctx: ReportContext) -> None:
    ws = wb.create_sheet("全ファイル一覧")
    loc_names = ctx.comparison.location_names
    headers = _file_columns_for(loc_names)
    _set_header_row(ws, headers)

    for i, row in enumerate(ctx.comparison.all_files, 1):
        _write_file_row(ws, i + 1, i, row, loc_names)

    # 列幅
    ws.column_dimensions["A"].width = 6   # No.
    ws.column_dimensions["B"].width = 50  # 相対パス
    ws.column_dimensions["C"].width = 28  # ファイル名
    ws.column_dimensions["D"].width = 8   # 拡張子
    ws.column_dimensions["E"].width = 14  # 状態
    base = 6
    for _ in loc_names:
        ws.column_dimensions[get_column_letter(base)].width = 12      # サイズ
        ws.column_dimensions[get_column_letter(base + 1)].width = 16  # ハッシュ (短く表示, セル選択でフル値確認)
        ws.column_dimensions[get_column_letter(base + 2)].width = 20  # 更新日時
        base += 3

    # 固定: 1行目 + 相対パス列 (B列まで固定 = C列以降スクロール)
    ws.freeze_panes = "C2"

    # オートフィルタ
    ws.auto_filter.ref = ws.dimensions


def _excel_subset(wb: WB, ctx: ReportContext, sheet_name: str, rows: List[FileRow]) -> None:
    ws = wb.create_sheet(sheet_name)
    loc_names = ctx.comparison.location_names
    headers = _file_columns_for(loc_names)
    _set_header_row(ws, headers)

    if not rows:
        ws.cell(row=2, column=1, value="該当なし")
        return

    for i, row in enumerate(rows, 1):
        _write_file_row(ws, i + 1, i, row, loc_names)

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 50
    ws.column_dimensions["C"].width = 28
    ws.column_dimensions["D"].width = 8
    ws.column_dimensions["E"].width = 14
    base = 6
    for _ in loc_names:
        ws.column_dimensions[get_column_letter(base)].width = 12
        ws.column_dimensions[get_column_letter(base + 1)].width = 16
        ws.column_dimensions[get_column_letter(base + 2)].width = 20
        base += 3
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = ws.dimensions


def _excel_dir_diff(wb: WB, ctx: ReportContext) -> None:
    ws = wb.create_sheet("フォルダ構造差分")
    loc_names = ctx.comparison.location_names
    headers = ["No.", "相対パス"] + loc_names
    _set_header_row(ws, headers)

    if not ctx.comparison.dir_diffs:
        ws.cell(row=2, column=1, value="該当なし")
        return

    for i, d in enumerate(ctx.comparison.dir_diffs, 1):
        ws.cell(row=i + 1, column=1, value=i)
        ws.cell(row=i + 1, column=2, value=d.relpath)
        for j, loc in enumerate(loc_names):
            v = "○" if d.presence[loc] else MISSING_PLACEHOLDER
            cell = ws.cell(row=i + 1, column=3 + j, value=v)
            cell.alignment = Alignment(horizontal="center")
            if not d.presence[loc]:
                cell.fill = FILL_GRAY

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 60
    for j in range(len(loc_names)):
        ws.column_dimensions[get_column_letter(3 + j)].width = 12
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = ws.dimensions


def _excel_errors(wb: WB, ctx: ReportContext) -> None:
    ws = wb.create_sheet("エラー")
    _set_header_row(ws, ["No.", "拠点", "相対パス", "メッセージ"])

    rows: List[tuple] = []
    for s in ctx.scans:
        for e in s.errors:
            rows.append((s.location_name, e.relpath, e.message))

    if not rows:
        ws.cell(row=2, column=1, value="該当なし")
        return

    for i, (loc, rel, msg) in enumerate(rows, 1):
        ws.cell(row=i + 1, column=1, value=i)
        ws.cell(row=i + 1, column=2, value=loc)
        ws.cell(row=i + 1, column=3, value=rel)
        ws.cell(row=i + 1, column=4, value=msg)

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 50
    ws.column_dimensions["D"].width = 80
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


# ============================================================
# HTML
# ============================================================
def write_html(ctx: ReportContext, out_path: Path) -> Path:
    elapsed = (ctx.finished_at - ctx.started_at).total_seconds()
    loc_names = ctx.comparison.location_names

    # 各拠点のルートパス文字列 (config.yaml に書かれた形式そのまま)
    location_roots = {s.location_name: str(s.root) for s in ctx.scans}

    summary_html = _html_summary_section(ctx, elapsed)
    all_files_html = _html_file_table("全ファイル一覧", ctx.comparison.all_files, loc_names, location_roots)
    mismatch_html = _html_file_table("ハッシュ不一致", ctx.comparison.hash_mismatches, loc_names, location_roots)
    missing_html = _html_file_table("ファイル欠落", ctx.comparison.missing_files, loc_names, location_roots)
    extra_html = _html_file_table("余分なファイル", ctx.comparison.extra_files, loc_names, location_roots)
    errored_html = _html_file_table("エラー対象ファイル", ctx.comparison.errored_files, loc_names, location_roots)
    dir_html = _html_dir_diff(ctx.comparison.dir_diffs, loc_names)
    err_html = _html_errors(ctx.scans)

    title = f"File Sync Check Report - {ctx.started_at.strftime(DATETIME_FMT)}"
    body = f"""
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<title>{html.escape(title)}</title>
<style>
{_HTML_STYLE}
</style>
</head>
<body>
<header>
  <h1>ファイル同期チェック レポート</h1>
  <nav>
    <a href="#summary">サマリー</a>
    <a href="#all">全ファイル一覧</a>
    <a href="#mismatch">ハッシュ不一致</a>
    <a href="#missing">ファイル欠落</a>
    <a href="#extra">余分なファイル</a>
    <a href="#errored">エラー対象ファイル</a>
    <a href="#dirs">フォルダ構造差分</a>
    <a href="#errors">エラー詳細</a>
  </nav>
</header>
<main>
  {summary_html}
  {all_files_html}
  {mismatch_html}
  {missing_html}
  {extra_html}
  {errored_html}
  {dir_html}
  {err_html}
</main>
<script>
{_HTML_SCRIPT}
</script>
</body>
</html>
"""
    out_path.write_text(body, encoding="utf-8")
    return out_path


_HTML_STYLE = """
* { box-sizing: border-box; }
body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Hiragino Sans",
       "Yu Gothic", sans-serif; margin: 0; color: #222; background: #fafafa; }
/* ページヘッダーは固定しない。代わりにテーブル列ヘッダーを sticky にして
   長い表でも常に列名が見える状態を保つ。 */
header { background: #305496; color: #fff; padding: 16px 24px; }
header h1 { margin: 0 0 8px; font-size: 1.3rem; }
header nav a { color: #fff; margin-right: 12px; padding: 4px 0;
               text-decoration: none; font-size: 0.9rem; }
header nav a:hover { text-decoration: underline; }
main { padding: 24px; }
section { background: #fff; border-radius: 6px; padding: 16px; margin-bottom: 24px;
          box-shadow: 0 1px 3px rgba(0,0,0,0.08);
          /* アンカージャンプで section 見出しが画面端に潜らないよう余白 */
          scroll-margin-top: 16px; }
section h2 { margin-top: 0; font-size: 1.1rem; border-bottom: 2px solid #305496; padding-bottom: 6px; }
table { border-collapse: collapse; width: 100%; font-size: 0.85rem; }
th, td { border: 1px solid #ddd; padding: 4px 8px; text-align: left; vertical-align: top;
         white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 320px; }
thead th { background: #305496; color: #fff; position: sticky; top: 0; z-index: 5; }
tbody tr:nth-child(odd) td { background: #fff; }
tbody tr:nth-child(even) td { background: #f6f8fb; }
td.hash { font-family: ui-monospace, "SF Mono", Menlo, Consolas, monospace; font-size: 0.78rem; }
td.num { text-align: right; font-variant-numeric: tabular-nums; }
/* ステータス・セル種別の塗りは zebra/sticky の背景を上書き */
.status-ok { background: #c6efce !important; color: #006100; font-weight: bold; }
.status-mismatch { background: #ffc7ce !important; color: #9c0006; font-weight: bold; }
.status-missing { background: #ffd8a8 !important; color: #9c4500; font-weight: bold; }
.status-partial { background: #ffe699 !important; color: #7f6000; font-weight: bold; }
.status-error { background: #f4b084 !important; color: #6a2900; font-weight: bold; }
.cell-mismatch { background: #ffc7ce !important; }
.cell-missing { background: #d9d9d9 !important; color: #888; text-align: center; }
.cell-error { background: #f4b084 !important; color: #6a2900; }
.summary-table td:first-child { font-weight: bold; width: 240px; background: #f0f3f8; }
/* テーブルラッパー: 横スクロール + 大量行時は内部で縦スクロール。
   max-height があると thead position:sticky が wrap スクロールに対して機能する。
   行数が少なければそのまま自然な高さで表示される。 */
.scroll-wrap { overflow: auto; max-width: 100%; max-height: 75vh; }
/* 相対パス列を左に固定。背景を明示しないと sticky 時に下のセルが透けるため、
   zebra に合わせて奇数行=白・偶数行=薄灰の背景を td に直接指定する。
   th 側もヘッダ色を明示し、行ヘッダより前面に配置する。 */
.fixed-col-table thead th:nth-child(2) {
  position: sticky; left: 0; z-index: 10; background: #305496;
}
.fixed-col-table tbody td:nth-child(2) {
  position: sticky; left: 0; z-index: 1;
}
.fixed-col-table tbody tr:nth-child(odd) td:nth-child(2) { background: #fff; }
.fixed-col-table tbody tr:nth-child(even) td:nth-child(2) { background: #f6f8fb; }
.empty { color: #888; font-style: italic; padding: 8px; }

/* ===== 行展開・コピー操作 ===== */
tr.file-row { cursor: pointer; }
tr.file-row:hover td { filter: brightness(0.97); }
tr.file-row .expand-indicator {
  display: inline-block; color: #888; margin-right: 6px;
  transition: transform 0.15s; user-select: none;
}
tr.file-row.expanded .expand-indicator { transform: rotate(90deg); }
tr.detail-row > td { padding: 0 !important; background: #f9fafe !important; }
/* 詳細パネルは viewport 幅以内に収め、内部テーブルが拡がりすぎないようにする。
   メインテーブルが横にスクロールされた場合は通常通りパネルも一緒にスクロールする
   (table 内での position:sticky はブラウザ実装が不安定なため非採用)。 */
.detail-panel {
  padding: 12px 24px;
  max-width: calc(100vw - 80px);
  box-sizing: border-box;
}
.detail-panel h4 {
  margin: 8px 0 4px; font-size: 0.85rem; color: #305496;
  border-bottom: 1px solid #d6deeb; padding-bottom: 2px;
}
.detail-table { width: 100%; max-width: 960px; margin-bottom: 8px; font-size: 0.8rem; }
.detail-table thead th { position: static; white-space: nowrap; }
.detail-table td, .detail-table th { max-width: none; padding: 4px 8px; }
.detail-table td { white-space: nowrap; }
/* ファイルパス・SHA-256 など長くなる列のみ折り返しを許可 */
.detail-table .path-cell, .detail-table .hash-cell {
  white-space: normal; overflow-wrap: anywhere; word-break: break-all;
}
/* ファイルパス列を最も広く、操作列はボタンに合わせて最小幅で */
.detail-table.detail-paths .path-cell { width: 100%; }
.detail-table.detail-paths .action-cell { white-space: nowrap; width: 1%; }
.detail-table code {
  background: #f4f4f8; padding: 2px 6px; border-radius: 3px;
  font-family: ui-monospace, "SF Mono", Menlo, Consolas, monospace;
  font-size: 0.78rem; word-break: break-all; overflow-wrap: anywhere;
  user-select: all; display: inline-block; max-width: 100%;
}
.detail-table .muted { color: #888; font-style: italic; }
button[data-copy], button[data-action] {
  background: #fff; border: 1px solid #305496; color: #305496;
  border-radius: 4px; padding: 3px 9px; cursor: pointer;
  font-size: 0.78rem; margin: 2px 4px 2px 0; font-family: inherit;
}
button[data-copy]:hover, button[data-action]:hover { background: #305496; color: #fff; }
button[data-copy].copied { background: #c6efce; color: #006100; border-color: #6b9e6e; }
.section-controls { float: right; font-weight: normal; }
.section-controls button { padding: 3px 10px; }
"""


_HTML_SCRIPT = """
(() => {
  // クリップボードコピー: 新 API → 旧 API (file:// など制限環境向け) フォールバック
  function copyText(text) {
    if (navigator.clipboard && navigator.clipboard.writeText && window.isSecureContext) {
      return navigator.clipboard.writeText(text).catch(() => fallbackCopy(text));
    }
    return Promise.resolve(fallbackCopy(text));
  }
  function fallbackCopy(text) {
    const ta = document.createElement("textarea");
    ta.value = text;
    ta.setAttribute("readonly", "");
    ta.style.position = "fixed";
    ta.style.top = "-1000px";
    ta.style.opacity = "0";
    document.body.appendChild(ta);
    ta.select();
    try { document.execCommand("copy"); } catch (e) { /* ignore */ }
    document.body.removeChild(ta);
  }

  function flashCopied(btn) {
    const original = btn.innerHTML;
    btn.classList.add("copied");
    btn.innerHTML = "&#10003; Copied";
    setTimeout(() => {
      btn.classList.remove("copied");
      btn.innerHTML = original;
    }, 1200);
  }

  document.addEventListener("click", (e) => {
    // 1) コピーボタン
    const copyBtn = e.target.closest("button[data-copy]");
    if (copyBtn) {
      e.stopPropagation();
      copyText(copyBtn.dataset.copy);
      flashCopied(copyBtn);
      return;
    }

    // 2) section レベルの 全展開 / 全折りたたみ
    const actionBtn = e.target.closest("button[data-action]");
    if (actionBtn) {
      e.stopPropagation();
      const expand = actionBtn.dataset.action === "expand-all";
      const section = actionBtn.closest("section");
      section.querySelectorAll("tr.detail-row").forEach(r => { r.hidden = !expand; });
      section.querySelectorAll("tr.file-row").forEach(r => r.classList.toggle("expanded", expand));
      return;
    }

    // 3) 行クリックで詳細行をトグル (テキスト選択中は除外)
    if (window.getSelection && window.getSelection().toString()) return;
    const row = e.target.closest("tr.file-row");
    if (row) {
      const detail = row.nextElementSibling;
      if (detail && detail.classList.contains("detail-row")) {
        detail.hidden = !detail.hidden;
        row.classList.toggle("expanded");
      }
    }
  });
})();
"""


def _status_class(status: str) -> str:
    return {
        STATUS_OK: "status-ok",
        STATUS_HASH_MISMATCH: "status-mismatch",
        STATUS_PARTIAL_MISSING: "status-missing",
        STATUS_PARTIAL_PRESENT: "status-partial",
        STATUS_ERROR: "status-error",
    }.get(status, "")


def _section_id(title: str) -> str:
    return {
        "全ファイル一覧": "all",
        "ハッシュ不一致": "mismatch",
        "ファイル欠落": "missing",
        "余分なファイル": "extra",
        "エラー対象ファイル": "errored",
    }.get(title, html.escape(title))


def _html_summary_section(ctx: ReportContext, elapsed: float) -> str:
    rows = [
        ("スキャン開始", ctx.started_at.strftime(DATETIME_FMT)),
        ("スキャン終了", ctx.finished_at.strftime(DATETIME_FMT)),
        ("所要時間 (秒)", f"{elapsed:.2f}"),
        ("設定ファイル", str(ctx.config_path)),
        ("拠点数", str(len(ctx.scans))),
    ]
    summary_kv = "".join(
        f"<tr><td>{html.escape(k)}</td><td>{html.escape(v)}</td></tr>" for k, v in rows
    )

    loc_rows = []
    for s in ctx.scans:
        total_size = sum(f.size for f in s.files.values())
        loc_rows.append(
            "<tr>"
            f"<td>{html.escape(s.location_name)}</td>"
            f"<td>{html.escape(str(s.root))}</td>"
            f"<td class='num'>{len(s.files):,}</td>"
            f"<td class='num'>{html.escape(human_bytes(total_size))}</td>"
            f"<td class='num'>{len(s.errors)}</td>"
            "</tr>"
        )

    diff_rows = [
        ("ハッシュ不一致", len(ctx.comparison.hash_mismatches)),
        ("ファイル欠落 (一部拠点になし)", len(ctx.comparison.missing_files)),
        ("余分なファイル (一部拠点のみ)", len(ctx.comparison.extra_files)),
        ("エラー対象ファイル (要確認)", len(ctx.comparison.errored_files)),
        ("フォルダ構造差分", len(ctx.comparison.dir_diffs)),
    ]
    diff_html = "".join(
        f"<tr><td>{html.escape(k)}</td><td class='num'>{v:,}</td></tr>" for k, v in diff_rows
    )

    return f"""
<section id="summary">
  <h2>サマリー</h2>
  <table class="summary-table">{summary_kv}</table>
  <h3>拠点別</h3>
  <div class="scroll-wrap"><table>
    <thead><tr><th>拠点</th><th>ルートパス</th><th>ファイル数</th><th>総サイズ</th><th>エラー数</th></tr></thead>
    <tbody>{"".join(loc_rows)}</tbody>
  </table></div>
  <h3>差分件数</h3>
  <table class="summary-table">{diff_html}</table>
</section>
"""


def _render_detail_row(
    row: FileRow,
    loc_names: List[str],
    location_roots: Dict[str, str],
    colspan: int,
) -> str:
    """行クリックで展開される詳細パネル (パス・ハッシュ + コピーボタン)。"""
    path_rows: List[str] = []
    for loc in loc_names:
        root = location_roots.get(loc, "")
        full, folder = _build_paths(root, row.relpath)
        err = row.errors.get(loc)
        entry = row.entries.get(loc)
        if err is not None:
            status_html = f"<td class='muted'>エラー</td>"
        elif entry is None:
            status_html = f"<td class='muted'>欠落</td>"
        else:
            status_html = f"<td>あり</td>"
        # フォルダパスは「📁 フォルダ」ボタンの data-copy に格納する (列としては非表示)
        path_rows.append(
            "<tr>"
            f"<td>{html.escape(loc)}</td>"
            f"{status_html}"
            f"<td class='path-cell'><code>{html.escape(full)}</code></td>"
            f"<td class='action-cell'>"
            f"<button type='button' data-copy=\"{html.escape(full, quote=True)}\" "
            f"title='ファイルのフルパスをコピー'>📋 ファイル</button>"
            f"<button type='button' data-copy=\"{html.escape(folder, quote=True)}\" "
            f"title='フォルダパスをコピー (Windows エクスプローラに貼り付けて開く)'>📁 フォルダ</button>"
            f"</td>"
            f"</tr>"
        )

    hash_rows: List[str] = []
    for loc in loc_names:
        entry = row.entries.get(loc)
        err = row.errors.get(loc)
        if err is not None:
            hash_rows.append(
                "<tr>"
                f"<td>{html.escape(loc)}</td>"
                f"<td colspan='2' class='muted'>エラー: {html.escape(err)}</td>"
                f"<td></td>"
                "</tr>"
            )
        elif entry is None:
            hash_rows.append(
                "<tr>"
                f"<td>{html.escape(loc)}</td>"
                f"<td colspan='2' class='muted'>欠落</td>"
                f"<td></td>"
                "</tr>"
            )
        else:
            hash_rows.append(
                "<tr>"
                f"<td>{html.escape(loc)}</td>"
                f"<td class='hash-cell'><code>{html.escape(entry.hash)}</code></td>"
                f"<td class='num'>{entry.size:,} B</td>"
                f"<td class='action-cell'><button type='button' data-copy=\"{html.escape(entry.hash, quote=True)}\" "
                f"title='SHA-256 フル値をコピー'>📋 ハッシュ</button></td>"
                "</tr>"
            )

    return (
        f'<tr class="detail-row" hidden><td colspan="{colspan}">'
        f'<div class="detail-panel">'
        f'<h4>パス情報</h4>'
        f'<table class="detail-table detail-paths">'
        f'<thead><tr><th>拠点</th><th>状態</th><th>ファイルパス</th><th>操作</th></tr></thead>'
        f'<tbody>{"".join(path_rows)}</tbody>'
        f'</table>'
        f'<h4>ハッシュ詳細 (SHA-256 フル値)</h4>'
        f'<table class="detail-table">'
        f'<thead><tr><th>拠点</th><th>SHA-256</th><th>サイズ</th><th>操作</th></tr></thead>'
        f'<tbody>{"".join(hash_rows)}</tbody>'
        f'</table>'
        f'</div></td></tr>'
    )


def _html_file_table(
    title: str,
    rows: List[FileRow],
    loc_names: List[str],
    location_roots: Dict[str, str],
) -> str:
    sid = _section_id(title)
    if not rows:
        return f"""
<section id="{sid}">
  <h2>{html.escape(title)}</h2>
  <p class="empty">該当なし</p>
</section>
"""

    headers = ["No.", "相対パス", "ファイル名", "拡張子", "状態"]
    for n in loc_names:
        headers.extend([f"{n}: サイズ", f"{n}: ハッシュ", f"{n}: 更新日時"])

    head_html = "".join(f"<th>{html.escape(h)}</th>" for h in headers)
    colspan = 5 + 3 * len(loc_names)

    body_lines: List[str] = []
    for i, row in enumerate(rows, 1):
        name = row.relpath.rsplit("/", 1)[-1]
        ext = Path(name).suffix
        status_cls = _status_class(row.status)
        # 相対パスセルに展開インジケータ (▶) を埋める
        relpath_cell = (
            f"<td title='{html.escape(row.relpath)}'>"
            f"<span class='expand-indicator'>&#9654;</span>"
            f"{html.escape(row.relpath)}</td>"
        )
        cells: List[str] = [
            f"<td class='num'>{i}</td>",
            relpath_cell,
            f"<td>{html.escape(name)}</td>",
            f"<td>{html.escape(ext)}</td>",
            f"<td class='{status_cls}'>{html.escape(row.status)}</td>",
        ]
        # ハッシュ不一致時は最頻値以外のハッシュだけ強調
        minority = (
            minority_hashes(row.entries) if row.status == STATUS_HASH_MISMATCH else set()
        )

        for loc in loc_names:
            err_msg = row.errors.get(loc)
            entry = row.entries.get(loc)
            if err_msg is not None:
                cells.append(f"<td class='cell-error'>{ERROR_PLACEHOLDER}</td>")
                cells.append(
                    f"<td class='cell-error' title='{html.escape(err_msg)}'>"
                    f"{html.escape(err_msg[:60])}</td>"
                )
                cells.append(f"<td class='cell-error'>{MISSING_PLACEHOLDER}</td>")
            elif entry is None:
                cells.append(f"<td class='cell-missing'>{MISSING_PLACEHOLDER}</td>")
                cells.append(f"<td class='cell-missing'>{MISSING_PLACEHOLDER}</td>")
                cells.append(f"<td class='cell-missing'>{MISSING_PLACEHOLDER}</td>")
            else:
                hash_cls = "hash"
                if entry.hash in minority:
                    hash_cls += " cell-mismatch"
                cells.append(f"<td class='num'>{entry.size:,}</td>")
                cells.append(
                    f"<td class='{hash_cls}' title='{html.escape(entry.hash)}'>"
                    f"{html.escape(entry.hash[:12])}…</td>"
                )
                cells.append(f"<td>{html.escape(entry.mtime.strftime(DATETIME_FMT))}</td>")
        body_lines.append(
            f'<tr class="file-row" data-relpath="{html.escape(row.relpath, quote=True)}">'
            + "".join(cells)
            + "</tr>"
        )
        body_lines.append(_render_detail_row(row, loc_names, location_roots, colspan))

    return f"""
<section id="{sid}">
  <h2>{html.escape(title)} ({len(rows):,} 件)
    <span class="section-controls">
      <button type="button" data-action="expand-all">全展開</button>
      <button type="button" data-action="collapse-all">全折りたたみ</button>
    </span>
  </h2>
  <div class="scroll-wrap"><table class="fixed-col-table">
    <thead><tr>{head_html}</tr></thead>
    <tbody>{"".join(body_lines)}</tbody>
  </table></div>
</section>
"""


def _html_dir_diff(dir_diffs, loc_names: List[str]) -> str:
    if not dir_diffs:
        return """
<section id="dirs"><h2>フォルダ構造差分</h2><p class="empty">該当なし</p></section>
"""
    headers = ["No.", "相対パス"] + loc_names
    head_html = "".join(f"<th>{html.escape(h)}</th>" for h in headers)
    rows = []
    for i, d in enumerate(dir_diffs, 1):
        cells = [f"<td class='num'>{i}</td>", f"<td>{html.escape(d.relpath)}</td>"]
        for loc in loc_names:
            if d.presence[loc]:
                cells.append("<td style='text-align:center;'>○</td>")
            else:
                cells.append(f"<td class='cell-missing'>{MISSING_PLACEHOLDER}</td>")
        rows.append("<tr>" + "".join(cells) + "</tr>")
    return f"""
<section id="dirs">
  <h2>フォルダ構造差分 ({len(dir_diffs):,} 件)</h2>
  <div class="scroll-wrap"><table>
    <thead><tr>{head_html}</tr></thead>
    <tbody>{"".join(rows)}</tbody>
  </table></div>
</section>
"""


def _html_errors(scans: List[ScanResult]) -> str:
    rows = []
    n = 0
    for s in scans:
        for e in s.errors:
            n += 1
            rows.append(
                "<tr>"
                f"<td class='num'>{n}</td>"
                f"<td>{html.escape(s.location_name)}</td>"
                f"<td>{html.escape(e.relpath)}</td>"
                f"<td>{html.escape(e.message)}</td>"
                "</tr>"
            )
    if not rows:
        return """
<section id="errors"><h2>エラー</h2><p class="empty">該当なし</p></section>
"""
    return f"""
<section id="errors">
  <h2>エラー ({n:,} 件)</h2>
  <div class="scroll-wrap"><table>
    <thead><tr><th>No.</th><th>拠点</th><th>相対パス</th><th>メッセージ</th></tr></thead>
    <tbody>{"".join(rows)}</tbody>
  </table></div>
</section>
"""
