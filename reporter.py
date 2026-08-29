"""レポート生成 (Excel / HTML)。"""
from __future__ import annotations

import html
import json
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterator, List, Optional

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
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
    minority_sizes,
)
from scanner import FileEntry, ScanResult
from utils import human_bytes


MISSING_PLACEHOLDER = "-"
ERROR_PLACEHOLDER = "エラー"
# hash_mode=smart でハッシュ計算を省略したファイルのハッシュ欄。
SKIPPED_HASH_PLACEHOLDER = "(未計算)"
DATETIME_FMT = "%Y-%m-%d %H:%M:%S"

# Excel の 1 シートあたりの行数上限 (1,048,576)。
# ヘッダー行と末尾の省略注記行を除いた分だけデータ行に使える。
EXCEL_MAX_DATA_ROWS = 1_048_576 - 2

# レポートにそのまま載せられない文字。
#   \x00-\x1f (改行・タブを除く): XML 1.0 に書けず openpyxl が例外を投げる。
#   \ud800-\udfff: 不正な UTF-8 のファイル名を Python が surrogateescape で
#       返したもの。UTF-8 にエンコードできず、HTML の書き出しが例外になる。
#       Excel は書けてしまうが、生成された xlsx が壊れて開けなくなる。
# 共有フォルダにはレガシー機器が付けた名前や文字コード不一致のファイルが実在するため、
# 1 ファイルでレポート全体を失わないよう、見える表記に置換して出力を続ける。
_UNSAFE_TEXT_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff]")


def _sanitize(text: str) -> str:
    """出力できない文字を `\\x07` / `\\udcff` のような見える表記に置き換える。

    置換後の文字列は Excel・HTML の双方で同じになる必要がある
    (HTML では行の照合キーとして使うため、表と詳細データで食い違うと
    詳細が引けなくなる)。
    """
    def _escape(m: "re.Match[str]") -> str:
        code = ord(m.group())
        return f"\\x{code:02x}" if code < 0x20 else f"\\u{code:04x}"

    return _UNSAFE_TEXT_RE.sub(_escape, text)


def _xl(value):
    """Excel セルに入れる値を整える (文字列以外はそのまま)。"""
    return _sanitize(value) if isinstance(value, str) else value


def _condition_rows(ctx: "ReportContext") -> List[tuple]:
    """サマリーに出す実行条件の (項目, 値) を返す。設定が無ければ空。"""
    s = ctx.settings
    if s is None:
        return []

    if s.hash_mode == "smart":
        mode = (
            "smart — サイズが違うファイルと、"
            "全拠点でサイズ・更新日時が一致するファイルは読み取りを省略"
        )
    else:
        mode = "always — 全ファイルの内容を読んで照合"

    rows = [
        ("ハッシュ方式", f"{s.hash_algorithm} / {mode}"),
        ("ハッシュ省略件数", f"{sum(sc.skipped_hashes for sc in ctx.scans):,} 件"),
    ]
    if s.hash_mode == "smart":
        rows.append(("更新日時の許容誤差", f"{s.mtime_tolerance_sec:g} 秒"))
    rows.append((
        "除外パターン",
        ", ".join(s.exclude_patterns) if s.exclude_patterns else "なし",
    ))
    rows.append((
        "ファイル名の照合",
        ("NFC 正規化あり" if s.normalize_unicode else "正規化なし")
        + " / "
        + ("大文字小文字を区別する" if s.case_sensitive else "大文字小文字を区別しない"),
    ))
    return rows


def _minority_targets(row: FileRow) -> tuple[set, set]:
    """不一致行で強調表示すべき (ハッシュ集合, サイズ集合) を返す。

    通常はハッシュの少数派を強調する。ハッシュを計算していない行
    (サイズ相違だけで不一致が確定したケース) ではサイズの少数派を強調する。
    """
    if row.status != STATUS_HASH_MISMATCH:
        return set(), set()
    hashes = minority_hashes(row.entries)
    if hashes:
        return hashes, set()
    return set(), minority_sizes(row.entries)


def _looks_windows(root_str: str) -> bool:
    """ルートが Windows UNC (`//` or `\\\\`) またはドライブレターか判定。"""
    return (
        root_str.startswith(("//", "\\\\"))
        or (len(root_str) >= 2 and root_str[1] == ":")
    )


def _normalize_root(root_str: str) -> tuple[str, str]:
    """拠点ルートを正規化し、(正規化ルート, セパレータ) を返す。

    Windows ルート判定ならセパレータをバックスラッシュに統一し、
    Windows エクスプローラのアドレス欄に貼り付けて開ける形式にする。

    ルート単位で1回だけ計算すればよいため、ファイル行ごとの再計算を避ける目的で
    `_build_paths` から切り出してある (HTML 側は正規化済みルートを JSON で受け取り、
    ファイルパスの組み立てはブラウザ側で行う)。
    """
    if _looks_windows(root_str):
        return root_str.replace("/", "\\").rstrip("\\"), "\\"
    return root_str.rstrip("/"), "/"


def _build_paths(root_str: str, relpath: str) -> tuple[str, str]:
    """拠点のルートと相対パスからフルパスとフォルダパスを構築する。

    Returns:
        (full_file_path, folder_path)
    """
    norm_root, sep = _normalize_root(root_str)
    rel = relpath.replace("/", sep) if sep == "\\" else relpath
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
class ReportSettings:
    """レポートに記録する「どういう条件で照合したか」。

    設定ファイルのパスだけでは、そのレポートがどの程度の保証を意味するのか
    読み手に分からない。特に `hash_mode: smart` は更新日時に基づく推定を含み、
    除外パターンに至っては「検査していないファイルがある」ことすら見えない。
    レポート単体で条件が追えるよう、実際に使われた値を持たせる。
    """

    hash_mode: str
    hash_algorithm: str
    mtime_tolerance_sec: float
    exclude_patterns: List[str]
    normalize_unicode: bool
    case_sensitive: bool


@dataclass
class ReportContext:
    started_at: datetime
    finished_at: datetime
    config_path: Path
    scans: List[ScanResult]
    comparison: ComparisonResult
    # 省略時は実行条件のセクションを出さない (呼び出し側が未対応の場合)
    settings: Optional[ReportSettings] = None


# ============================================================
# Excel
# ============================================================
def write_excel(ctx: ReportContext, out_path: Path) -> Path:
    # write_only: 行を書いた端から解放するため、シート全体をメモリに持たない。
    # 通常モードは 50,000 行 × 3 拠点で約 236MB を使い、行数に比例して増える。
    # 代わりにセルへのランダムアクセスができないので、各シートは 1 行ずつ append する。
    wb: WB = Workbook(write_only=True)

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


def _cell(ws, value, *, fill=None, font=None, align=None):
    """write_only シート用のセルを作る。

    write_only では ws.cell(row=, column=) が使えないため、
    書式付きのセルは WriteOnlyCell を組み立てて append する。
    """
    c = WriteOnlyCell(ws, value=_xl(value))
    if fill is not None:
        c.fill = fill
    if font is not None:
        c.font = font
    if align is not None:
        c.alignment = align
    return c


def _start_table(wb: WB, sheet_name: str, headers: List[str], *, freeze: str = "C2"):
    """シートを作り、ペイン固定を設定してからヘッダー行を書く。

    write_only ではシートビューが行より先に出力されるため、`freeze_panes` は
    最初の append より前に設定しないと保存時に捨てられる。
    """
    ws = wb.create_sheet(sheet_name)
    ws.freeze_panes = freeze
    ws.append([
        _cell(ws, h, fill=FILL_HEADER, font=FONT_HEADER,
              align=Alignment(horizontal="center", vertical="center"))
        for h in headers
    ])
    return ws


def _finish_table(ws, n_columns: int, n_rows: int) -> None:
    """オートフィルタの範囲を設定する。

    write_only では ws.dimensions がまだ確定していないため、範囲を自分で組み立てる。
    """
    if n_rows > 0:
        ws.auto_filter.ref = f"A1:{get_column_letter(n_columns)}{n_rows + 1}"


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

    conditions = _condition_rows(ctx)
    if conditions:
        rows.append([])
        rows.append(["実行条件", "値"])
        rows.extend([k, v] for k, v in conditions)

    rows.append([])
    rows.append(["差分種別", "件数"])
    rows.append(["ハッシュ不一致", len(ctx.comparison.hash_mismatches)])
    rows.append(["ファイル欠落 (一部拠点になし)", len(ctx.comparison.missing_files)])
    rows.append(["余分なファイル (一部拠点のみ)", len(ctx.comparison.extra_files)])
    rows.append(["エラー対象ファイル (要確認)", len(ctx.comparison.errored_files)])
    rows.append(["フォルダ構造差分", len(ctx.comparison.dir_diffs)])

    bold = Font(bold=True)
    for r_idx, row in enumerate(rows, 1):
        # 見出し行 (1行目と各表の先頭行) だけ太字にする
        is_header = r_idx == 1 or (
            row and isinstance(row[0], str)
            and row[0] in {"拠点", "差分種別", "項目", "実行条件"}
        )
        ws.append([
            _cell(ws, val, font=bold) if is_header else _xl(val) for val in row
        ])

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


def _write_file_row(ws, no: int, file_row: FileRow, location_names: List[str]) -> None:
    rel = file_row.relpath
    name = rel.rsplit("/", 1)[-1]
    ext = Path(name).suffix
    center = Alignment(horizontal="center")

    cells = [
        no,
        _xl(rel),
        _xl(name),
        _xl(ext),
        _cell(ws, file_row.status, fill=_status_fill(file_row.status)),
    ]

    # ハッシュ不一致時は「最頻値（過半数）と異なる拠点だけ」赤くする。
    # ハッシュ未計算 (サイズ相違だけで不一致が確定) の場合はサイズ列を強調する。
    minority, minority_size = _minority_targets(file_row)

    for loc in location_names:
        err_msg = file_row.errors.get(loc)
        entry: Optional[FileEntry] = file_row.entries.get(loc)
        if err_msg is not None:
            # 読み取り失敗: サイズ列=エラー、ハッシュ列=メッセージ、更新日時列=―
            cells.extend([
                _cell(ws, ERROR_PLACEHOLDER, fill=FILL_ERROR, align=center),
                _cell(ws, err_msg, fill=FILL_ERROR),
                _cell(ws, MISSING_PLACEHOLDER, fill=FILL_ERROR, align=center),
            ])
        elif entry is None:
            cells.extend(
                _cell(ws, MISSING_PLACEHOLDER, fill=FILL_GRAY, align=center)
                for _ in range(3)
            )
        else:
            cells.extend([
                _cell(
                    ws, entry.size,
                    fill=FILL_HASH_MISMATCH if entry.size in minority_size else None,
                ),
                _cell(
                    ws,
                    entry.hash if entry.hash is not None else SKIPPED_HASH_PLACEHOLDER,
                    fill=(
                        FILL_HASH_MISMATCH
                        if entry.hash is not None and entry.hash in minority
                        else None
                    ),
                ),
                entry.mtime.strftime(DATETIME_FMT),
            ])
    ws.append(cells)


def _truncate_for_excel(rows: List[FileRow]) -> tuple[List[FileRow], int]:
    """Excel の行数上限に収まるよう切り詰め、(表示する行, 省略した件数) を返す。

    上限を超えると openpyxl が保存時に失敗し、レポートが 1 枚も残らない。
    切り詰めたことは末尾の注記行で明示し、全件は HTML レポート側で見てもらう。
    """
    if len(rows) <= EXCEL_MAX_DATA_ROWS:
        return rows, 0
    return rows[:EXCEL_MAX_DATA_ROWS], len(rows) - EXCEL_MAX_DATA_ROWS


def _write_truncation_note(ws, omitted: int) -> None:
    ws.append([_cell(
        ws,
        f"... 他 {omitted:,} 件は Excel の行数上限のため省略しました "
        f"(全件は HTML レポートを参照してください)",
        font=Font(bold=True, color="9C0006"),
    )])


def _set_file_column_widths(ws, loc_names: List[str]) -> None:
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


def _excel_all_files(wb: WB, ctx: ReportContext) -> None:
    loc_names = ctx.comparison.location_names
    headers = _file_columns_for(loc_names)
    # 固定: 1行目 + 相対パス列 (B列まで固定 = C列以降スクロール)
    ws = _start_table(wb, "全ファイル一覧", headers)

    rows, omitted = _truncate_for_excel(ctx.comparison.all_files)
    for i, row in enumerate(rows, 1):
        _write_file_row(ws, i, row, loc_names)
    if omitted:
        _write_truncation_note(ws, omitted)

    _set_file_column_widths(ws, loc_names)
    _finish_table(ws, len(headers), len(rows))


def _excel_subset(wb: WB, ctx: ReportContext, sheet_name: str, rows: List[FileRow]) -> None:
    loc_names = ctx.comparison.location_names
    headers = _file_columns_for(loc_names)
    ws = _start_table(wb, sheet_name, headers)

    if not rows:
        ws.append(["該当なし"])
        return

    rows, omitted = _truncate_for_excel(rows)
    for i, row in enumerate(rows, 1):
        _write_file_row(ws, i, row, loc_names)
    if omitted:
        _write_truncation_note(ws, omitted)

    _set_file_column_widths(ws, loc_names)
    _finish_table(ws, len(headers), len(rows))


def _excel_dir_diff(wb: WB, ctx: ReportContext) -> None:
    loc_names = ctx.comparison.location_names
    headers = ["No.", "相対パス"] + loc_names
    ws = _start_table(wb, "フォルダ構造差分", headers)

    if not ctx.comparison.dir_diffs:
        ws.append(["該当なし"])
        return

    center = Alignment(horizontal="center")
    for i, d in enumerate(ctx.comparison.dir_diffs, 1):
        cells = [i, _xl(d.relpath)]
        for loc in loc_names:
            present = d.presence[loc]
            cells.append(_cell(
                ws,
                "○" if present else MISSING_PLACEHOLDER,
                fill=None if present else FILL_GRAY,
                align=center,
            ))
        ws.append(cells)

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 60
    for j in range(len(loc_names)):
        ws.column_dimensions[get_column_letter(3 + j)].width = 12
    _finish_table(ws, len(headers), len(ctx.comparison.dir_diffs))


def _excel_errors(wb: WB, ctx: ReportContext) -> None:
    ws = _start_table(
        wb, "エラー", ["No.", "拠点", "相対パス", "メッセージ"], freeze="A2"
    )

    rows: List[tuple] = []
    for s in ctx.scans:
        for e in s.errors:
            rows.append((s.location_name, e.relpath, e.message))

    if not rows:
        ws.append(["該当なし"])
        return

    for i, (loc, rel, msg) in enumerate(rows, 1):
        ws.append([i, _xl(loc), _xl(rel), _xl(msg)])

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 50
    ws.column_dimensions["D"].width = 80
    _finish_table(ws, 4, len(rows))


# ============================================================
# HTML
# ============================================================
def write_html(ctx: ReportContext, out_path: Path) -> Path:
    """HTML レポートを書き出す。

    レポート全体を 1 つの文字列に組み立てると、行数に比例してメモリを食う
    (50,000 行 × 3 拠点でピーク約 500MB)。断片を順にファイルへ書き出す。
    """
    with out_path.open("w", encoding="utf-8") as f:
        for chunk in _iter_html_document(ctx):
            f.write(chunk)
    return out_path


def _iter_html_document(ctx: ReportContext) -> Iterator[str]:
    elapsed = (ctx.finished_at - ctx.started_at).total_seconds()
    loc_names = ctx.comparison.location_names
    title = f"File Sync Check Report - {ctx.started_at.strftime(DATETIME_FMT)}"

    yield f"""
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
  """
    yield _html_summary_section(ctx, elapsed)
    for section_title, rows in (
        ("全ファイル一覧", ctx.comparison.all_files),
        ("ハッシュ不一致", ctx.comparison.hash_mismatches),
        ("ファイル欠落", ctx.comparison.missing_files),
        ("余分なファイル", ctx.comparison.extra_files),
        ("エラー対象ファイル", ctx.comparison.errored_files),
    ):
        yield from _iter_html_file_table(section_title, rows, loc_names)
    yield _html_dir_diff(ctx.comparison.dir_diffs, loc_names)
    yield _html_errors(ctx.scans)
    yield """
</main>
<dialog id="detail-modal" aria-labelledby="detail-modal-title">
  <div class="modal-header">
    <h3 id="detail-modal-title">詳細情報</h3>
    <button type="button" class="modal-close" aria-label="閉じる">&times;</button>
  </div>
  <div class="modal-content" id="detail-modal-content"></div>
  <div class="modal-footer">
    <button type="button" class="modal-close">閉じる</button>
  </div>
</dialog>
<script type="application/json" id="report-data">
"""
    yield from _iter_report_data_json(ctx)
    yield f"""
</script>
<script>
{_HTML_SCRIPT}
</script>
</body>
</html>
"""


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
td.hash.muted { color: #888; font-style: italic; }
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
/* パスは省略すると確認手段が無くなるため折り返して全体を出す。
   行数が少ない表 (サマリー・拠点別・エラー) でのみ使う。 */
.wrap-path {
  white-space: normal; overflow: visible; text-overflow: clip;
  max-width: none; overflow-wrap: anywhere; word-break: break-all;
}
.wrap-path button[data-copy] { margin-left: 6px; vertical-align: baseline; }
/* テーブルラッパー: 横スクロール + 大量行時は内部で縦スクロール。
   max-height があると thead position:sticky が wrap スクロールに対して機能する。
   行数が少なければそのまま自然な高さで表示される。 */
.scroll-wrap { overflow: auto; max-width: 100%; max-height: 75vh; }
/* ファイル名列 (3列目) を左に固定。横スクロール時に「どのファイルの行か」を
   常に見えるようにするため。背景を明示しないと sticky 時に下のセルが透けるため、
   zebra に合わせて奇数行=白・偶数行=薄灰の背景を td に直接指定する。
   th 側もヘッダ色を明示し、行ヘッダより前面に配置する。 */
.fixed-col-table thead th:nth-child(3) {
  position: sticky; left: 0; z-index: 10; background: #305496;
}
.fixed-col-table tbody td:nth-child(3) {
  position: sticky; left: 0; z-index: 1;
}
.fixed-col-table tbody tr:nth-child(odd) td:nth-child(3) { background: #fff; }
.fixed-col-table tbody tr:nth-child(even) td:nth-child(3) { background: #f6f8fb; }
.empty { color: #888; font-style: italic; padding: 8px; }

/* ===== 絞り込みツールバー ===== */
.table-tools {
  display: flex; align-items: center; gap: 8px; flex-wrap: wrap;
  margin: 0 0 8px 0; font-size: 0.85rem;
}
.table-tools input.filter-input {
  flex: 1 1 260px; min-width: 160px; max-width: 420px;
  padding: 5px 9px; border: 1px solid #c3cbd8; border-radius: 4px;
  font-family: inherit; font-size: 0.85rem;
}
.table-tools select.filter-status {
  padding: 5px 8px; border: 1px solid #c3cbd8; border-radius: 4px;
  font-family: inherit; font-size: 0.85rem; background: #fff;
}
.table-tools .filter-count { color: #555; font-variant-numeric: tabular-nums; }
.table-tools .filter-count.filtered { color: #305496; font-weight: bold; }

/* ===== ファイル行 (クリックでモーダル表示) ===== */
tr.file-row { cursor: pointer; }
tr.file-row:hover td { filter: brightness(0.97); }

/* ===== 詳細モーダル ===== */
dialog#detail-modal {
  padding: 0; border: none; border-radius: 8px;
  width: min(900px, 92vw); max-height: 90vh;
  box-shadow: 0 12px 32px rgba(0,0,0,0.25);
  background: #fff; overflow: hidden;
  /* native <dialog> centering */
}
dialog#detail-modal::backdrop { background: rgba(0,0,0,0.45); }
.modal-header {
  background: #305496; color: #fff; padding: 12px 20px;
  display: flex; align-items: center; justify-content: space-between;
}
.modal-header h3 { margin: 0; font-size: 1rem; font-weight: normal; }
.modal-close {
  background: transparent; border: none; color: inherit;
  cursor: pointer; line-height: 1; font-family: inherit; padding: 0 4px;
}
.modal-header .modal-close { font-size: 1.4rem; }
.modal-content {
  padding: 16px 24px; overflow-y: auto;
  max-height: calc(90vh - 110px);
}
.modal-footer {
  padding: 8px 20px; border-top: 1px solid #e0e0e0; text-align: right;
}
.modal-footer .modal-close {
  background: #305496; color: #fff; padding: 6px 18px;
  border-radius: 4px; font-size: 0.9rem;
}
.modal-footer .modal-close:hover { background: #243d6f; }
.modal-summary {
  display: grid; grid-template-columns: max-content 1fr; gap: 4px 12px;
  margin: 0 0 12px 0; font-size: 0.85rem;
}
.modal-summary dt { font-weight: bold; color: #555; }
.modal-summary dd { margin: 0; }
.modal-summary dd code {
  background: #f4f4f8; padding: 2px 6px; border-radius: 3px;
  font-family: ui-monospace, "SF Mono", Menlo, Consolas, monospace;
  font-size: 0.82rem; word-break: break-all;
}

/* ===== モーダル内詳細テーブル ===== */
.detail-table { width: 100%; margin-bottom: 12px; font-size: 0.82rem; }
.detail-table thead th { position: static; white-space: nowrap; }
.detail-table td, .detail-table th { max-width: none; padding: 4px 8px; }
.detail-table td { white-space: nowrap; }
.detail-table .path-cell, .detail-table .hash-cell {
  white-space: normal; overflow-wrap: anywhere; word-break: break-all;
}
.detail-table.detail-paths .path-cell { width: 100%; }
.detail-table.detail-paths .action-cell { white-space: nowrap; width: 1%; }
.detail-table code {
  background: #f4f4f8; padding: 2px 6px; border-radius: 3px;
  font-family: ui-monospace, "SF Mono", Menlo, Consolas, monospace;
  font-size: 0.78rem; word-break: break-all; overflow-wrap: anywhere;
  user-select: all; display: inline-block; max-width: 100%;
}
.detail-table .muted { color: #888; font-style: italic; }

/* ===== コピーボタン ===== */
button[data-copy] {
  background: #fff; border: 1px solid #305496; color: #305496;
  border-radius: 4px; padding: 3px 9px; cursor: pointer;
  font-size: 0.78rem; margin: 2px 4px 2px 0; font-family: inherit;
}
button[data-copy]:hover { background: #305496; color: #fff; }
button[data-copy].copied { background: #c6efce; color: #006100; border-color: #6b9e6e; }
"""


_HTML_SCRIPT = """
(() => {
  const modal = document.getElementById("detail-modal");
  const modalTitle = document.getElementById("detail-modal-title");
  const modalContent = document.getElementById("detail-modal-content");

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
    const original = btn.textContent;
    btn.classList.add("copied");
    btn.textContent = "\\u2713 Copied";
    setTimeout(() => {
      btn.classList.remove("copied");
      btn.textContent = original;
    }, 1200);
  }

  // ===== 詳細データ =====
  // 詳細 DOM は行ごとに埋め込まず、#report-data の JSON から都度組み立てる。
  // 値は textContent 経由でのみ挿入するため、パス名に HTML が混ざっても解釈されない。
  const DATA = JSON.parse(document.getElementById("report-data").textContent);
  const LOCS = DATA.locations;
  const SCLS = DATA.statusClasses;

  function el(tag, props, children) {
    const node = document.createElement(tag);
    if (props) {
      for (const key of Object.keys(props)) {
        if (props[key] === null || props[key] === undefined) continue;
        if (key === "class") node.className = props[key];
        else if (key === "text") node.textContent = props[key];
        else node.setAttribute(key, props[key]);
      }
    }
    (children || []).forEach((c) => node.appendChild(c));
    return node;
  }

  function headerRow(labels) {
    return el("thead", null, [
      el("tr", null, labels.map((l) => el("th", { text: l }))),
    ]);
  }

  function copyButton(value, label, title) {
    return el("button", { type: "button", "data-copy": value, title: title, text: label });
  }

  function fullPathFor(loc, relpath) {
    const rel = loc.sep === "\\\\" ? relpath.split("/").join("\\\\") : relpath;
    return loc.root + loc.sep + rel;
  }

  function folderPathFor(full, sep) {
    const i = full.lastIndexOf(sep);
    return i >= 0 ? full.slice(0, i) : full;
  }

  function buildDetail(relpath, row) {
    const frag = document.createDocumentFragment();

    frag.appendChild(el("dl", { class: "modal-summary" }, [
      el("dt", { text: "相対パス" }),
      el("dd", null, [el("code", { text: relpath })]),
      el("dt", { text: "状態" }),
      el("dd", { class: SCLS[row.s] || "", text: row.s }),
    ]));

    // --- パス情報 ---
    frag.appendChild(el("h4", { text: "パス情報" }));
    const pathBody = el("tbody");
    LOCS.forEach((loc, i) => {
      const err = row.x ? row.x[i] : null;
      const cell = row.c[i];
      // 拠点ごとに実ファイル名が違う場合は row.p にその拠点でのパスが入る
      const rel = (row.p && row.p[i]) || relpath;
      const full = fullPathFor(loc, rel);
      let state;
      if (err) state = el("td", { class: "muted", text: "エラー" });
      else if (!cell) state = el("td", { class: "muted", text: "欠落" });
      else state = el("td", { text: "あり" });
      pathBody.appendChild(el("tr", null, [
        el("td", { text: loc.name }),
        state,
        el("td", { class: "path-cell" }, [el("code", { text: full })]),
        // フォルダパスは列としては表示せず「📁 フォルダ」ボタンからコピーする
        el("td", { class: "action-cell" }, [
          copyButton(full, "📋 ファイル", "ファイルのフルパスをコピー"),
          copyButton(folderPathFor(full, loc.sep), "📁 フォルダ",
                     "フォルダパスをコピー (Windows エクスプローラに貼り付けて開く)"),
        ]),
      ]));
    });
    frag.appendChild(el("table", { class: "detail-table detail-paths" }, [
      headerRow(["拠点", "状態", "ファイルパス", "操作"]), pathBody,
    ]));

    // --- ハッシュ詳細 ---
    frag.appendChild(el("h4", { text: "ハッシュ詳細 (SHA-256 フル値)" }));
    const hashBody = el("tbody");
    LOCS.forEach((loc, i) => {
      const err = row.x ? row.x[i] : null;
      const cell = row.c[i];
      if (err) {
        hashBody.appendChild(el("tr", null, [
          el("td", { text: loc.name }),
          el("td", { colspan: "2", class: "muted", text: "エラー: " + err }),
          el("td"),
        ]));
      } else if (!cell) {
        hashBody.appendChild(el("tr", null, [
          el("td", { text: loc.name }),
          el("td", { colspan: "2", class: "muted", text: "欠落" }),
          el("td"),
        ]));
      } else if (cell[1] === null) {
        // hash_mode=smart でハッシュ計算を省略したファイル
        hashBody.appendChild(el("tr", null, [
          el("td", { text: loc.name }),
          el("td", { class: "muted", text: "(未計算 — サイズ・更新日時で判定)" }),
          el("td", { class: "num", text: cell[0].toLocaleString() + " B" }),
          el("td"),
        ]));
      } else {
        hashBody.appendChild(el("tr", null, [
          el("td", { text: loc.name }),
          el("td", { class: "hash-cell" }, [el("code", { text: cell[1] })]),
          el("td", { class: "num", text: cell[0].toLocaleString() + " B" }),
          el("td", { class: "action-cell" }, [
            copyButton(cell[1], "📋 ハッシュ", "SHA-256 フル値をコピー"),
          ]),
        ]));
      }
    });
    frag.appendChild(el("table", { class: "detail-table" }, [
      headerRow(["拠点", "SHA-256", "サイズ", "操作"]), hashBody,
    ]));

    return frag;
  }

  function openDetailModal(tr) {
    const relpath = tr.dataset.relpath;
    const row = DATA.rows[relpath];
    modalTitle.textContent = "詳細情報: " + relpath;
    modalContent.textContent = "";
    if (row) modalContent.appendChild(buildDetail(relpath, row));
    if (typeof modal.showModal === "function") {
      modal.showModal();
    } else {
      // <dialog> 非対応ブラウザ向けフォールバック (非常に古い環境のみ)
      modal.setAttribute("open", "");
    }
  }

  function closeModal() {
    if (typeof modal.close === "function" && modal.open) {
      modal.close();
    } else {
      modal.removeAttribute("open");
    }
  }

  // ESC キーでモーダルを閉じる (native dialog でも効くが、明示しておく)
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && modal.open) {
      e.preventDefault();
      closeModal();
    }
  });

  // ===== 絞り込み =====
  // 行データは DATA.rows にあるので、状態の判定に data 属性を増やす必要はない。
  function setUpFilters(section) {
    const input = section.querySelector(".filter-input");
    if (!input) return;
    const select = section.querySelector(".filter-status");
    const countEl = section.querySelector(".filter-count");
    const rows = Array.from(section.querySelectorAll("tbody tr.file-row"));
    const total = rows.length;

    function apply() {
      const q = input.value.trim().toLowerCase();
      const status = select ? select.value : "";
      let shown = 0;
      for (const tr of rows) {
        const rel = tr.dataset.relpath || "";
        const data = DATA.rows[rel];
        const hit =
          (!q || rel.toLowerCase().indexOf(q) !== -1) &&
          (!status || (data && data.s === status));
        tr.style.display = hit ? "" : "none";
        if (hit) shown++;
      }
      const filtered = shown !== total;
      countEl.textContent = filtered
        ? shown.toLocaleString() + " / " + total.toLocaleString() + " 件"
        : total.toLocaleString() + " 件";
      countEl.classList.toggle("filtered", filtered);
    }

    // 行数が多いと 1 打鍵ごとの再計算が重くなるため少し待ってからまとめて処理する
    let timer = null;
    function schedule() {
      clearTimeout(timer);
      timer = setTimeout(apply, total > 5000 ? 200 : 60);
    }
    input.addEventListener("input", schedule);
    if (select) select.addEventListener("change", apply);
    apply();
  }
  document.querySelectorAll("section").forEach(setUpFilters);

  document.addEventListener("click", (e) => {
    // 1) コピーボタン (モーダル内・モーダル外どちらでも)
    const copyBtn = e.target.closest("button[data-copy]");
    if (copyBtn) {
      e.stopPropagation();
      copyText(copyBtn.dataset.copy);
      flashCopied(copyBtn);
      return;
    }

    // 2) モーダル閉じるボタン
    if (e.target.closest(".modal-close")) {
      e.stopPropagation();
      closeModal();
      return;
    }

    // 3) バックドロップクリックで閉じる (dialog 自身が target になる場合)
    if (e.target === modal) {
      closeModal();
      return;
    }

    // 4) 行クリックでモーダル展開 (テキスト選択中は除外)
    if (window.getSelection && window.getSelection().toString()) return;
    const row = e.target.closest("tr.file-row");
    if (row) {
      openDetailModal(row);
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


def _path_cell(value: str) -> str:
    """省略せず全体を表示し、コピーもできるパスセルを返す。

    サマリーのパスは行数が少ないので、他の表のように 1 行に詰めて
    末尾を省略する必要がない。省略するとフルパスを確認する手段が無くなる
    (ファイル行と違ってモーダルが無いため)。
    """
    value = _sanitize(value)
    return (
        f"<td class='wrap-path'>{html.escape(value)}"
        f"<button type='button' data-copy=\"{html.escape(value, quote=True)}\" "
        f"title='パスをコピー'>📋</button></td>"
    )


def _html_summary_section(ctx: ReportContext, elapsed: float) -> str:
    rows = [
        ("スキャン開始", ctx.started_at.strftime(DATETIME_FMT), False),
        ("スキャン終了", ctx.finished_at.strftime(DATETIME_FMT), False),
        ("所要時間 (秒)", f"{elapsed:.2f}", False),
        ("設定ファイル", str(ctx.config_path), True),
        ("拠点数", str(len(ctx.scans)), False),
    ]
    summary_kv = "".join(
        f"<tr><td>{html.escape(k)}</td>"
        + (_path_cell(v) if is_path else f"<td>{html.escape(v)}</td>")
        + "</tr>"
        for k, v, is_path in rows
    )

    loc_rows = []
    for s in ctx.scans:
        total_size = sum(f.size for f in s.files.values())
        loc_rows.append(
            "<tr>"
            f"<td>{html.escape(_sanitize(s.location_name))}</td>"
            f"{_path_cell(str(s.root))}"
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

    conditions = _condition_rows(ctx)
    conditions_html = ""
    if conditions:
        body = "".join(
            f"<tr><td>{html.escape(k)}</td>"
            f"<td class='wrap-path'>{html.escape(v)}</td></tr>"
            for k, v in conditions
        )
        conditions_html = f"""
  <h3>実行条件</h3>
  <table class="summary-table">{body}</table>
"""

    return f"""
<section id="summary">
  <h2>サマリー</h2>
  <table class="summary-table">{summary_kv}</table>
{conditions_html}
  <h3>拠点別</h3>
  <div class="scroll-wrap"><table>
    <thead><tr><th>拠点</th><th>ルートパス</th><th>ファイル数</th><th>総サイズ</th><th>エラー数</th></tr></thead>
    <tbody>{"".join(loc_rows)}</tbody>
  </table></div>
  <h3>差分件数</h3>
  <table class="summary-table">{diff_html}</table>
</section>
"""


def _json_chunk(value) -> str:
    """JSON 断片を出力用にエンコードする。

    `</script>` でブロックが閉じられるのを防ぐため '<' をエスケープする。
    ensure_ascii=False で日本語パスをそのまま出す (ドキュメントは UTF-8)。
    """
    return json.dumps(value, ensure_ascii=False, separators=(",", ":")).replace(
        "<", "\\u003c"
    )


def _iter_report_data_json(ctx: ReportContext) -> Iterator[str]:
    """詳細表示用データを JSON として少しずつ書き出す。

    全行分を 1 つの dict に組み立ててから dumps すると、行数に比例して
    メモリを食う。行ごとにエンコードして流す。
    """
    yield "{"
    yield f'"locations":{_json_chunk(_report_locations(ctx))},'
    yield f'"statusClasses":{_json_chunk(_STATUS_CLASSES)},'
    yield '"rows":{'
    for i, (key, data) in enumerate(_iter_report_rows(ctx)):
        if i:
            yield ","
        yield f"{_json_chunk(key)}:{_json_chunk(data)}"
    yield "}}"


_STATUS_CLASSES = {
    STATUS_OK: "status-ok",
    STATUS_HASH_MISMATCH: "status-mismatch",
    STATUS_PARTIAL_MISSING: "status-missing",
    STATUS_PARTIAL_PRESENT: "status-partial",
    STATUS_ERROR: "status-error",
}


def _report_locations(ctx: ReportContext) -> List[dict]:
    """拠点ごとの正規化済みルートとセパレータ。

    フルパスはブラウザ側が root + sep + 相対パスで組み立てるため、
    ルートの正規化は拠点あたり 1 回で済む。
    """
    roots = {s.location_name: str(s.root) for s in ctx.scans}
    locations = []
    for name in ctx.comparison.location_names:
        norm_root, sep = _normalize_root(roots.get(name, ""))
        locations.append(
            {"name": _sanitize(name), "root": _sanitize(norm_root), "sep": sep}
        )
    return locations


def _iter_report_rows(ctx: ReportContext) -> Iterator[tuple]:
    """詳細表示用の (相対パス, 行データ) を 1 行ずつ返す。

    行ごとに詳細 HTML を組み立てて data 属性に埋め込むと、拠点数 × 行数分のマークアップが
    そのまま出力に乗る (実測: 1,000 ファイル × 3 拠点で HTML 6.7MB のうち 87% が詳細属性)。
    同一の相対パスが「全ファイル一覧」と各差分セクションに重複して現れる分も二重に載る。

    そこで相対パスをキーにしたデータブロック 1 つに集約し、詳細 DOM の組み立ては
    ブラウザ側 (`_HTML_SCRIPT`) に任せる。重複はキーで自然に解消される。

    行データの形式 (バイト数を抑えるため短いキー・拠点順の配列を使う):
        {s: 状態,
         c: [[size, hash, mtime] | null, ...],
         x: [エラーメッセージ | null, ...]   ← エラーが無い行では省略,
         p: [拠点ごとの実際の相対パス, ...]  ← 表示用パスと同じなら省略}
    """
    loc_names = ctx.comparison.location_names
    for row in ctx.comparison.all_files:
        cells: List[Optional[list]] = []
        errs: List[Optional[str]] = []
        for loc in loc_names:
            entry = row.entries.get(loc)
            cells.append(
                None if entry is None
                else [entry.size, entry.hash, entry.mtime.strftime(DATETIME_FMT)]
            )
            errs.append(row.errors.get(loc))
        data: dict = {"s": row.status, "c": cells}
        if any(e is not None for e in errs):
            data["x"] = [_sanitize(e) if e is not None else None for e in errs]
        # 拠点ごとに実際のファイル名が違う場合 (Unicode 正規化形・大小文字) のみ、
        # フルパス組み立て用の相対パスを持たせる。通常は表示用パスと同じなので省略する。
        relpath = _sanitize(row.relpath)
        reals = [
            _sanitize(r) if (r := row.real_relpaths.get(loc)) is not None else None
            for loc in loc_names
        ]
        if any(r is not None and r != relpath for r in reals):
            data["p"] = reals
        yield relpath, data


def _iter_html_file_table(
    title: str,
    rows: List[FileRow],
    loc_names: List[str],
) -> Iterator[str]:
    """ファイル一覧セクションを行単位で少しずつ生成する。"""
    sid = _section_id(title)
    if not rows:
        yield f"""
<section id="{sid}">
  <h2>{html.escape(title)}</h2>
  <p class="empty">該当なし</p>
</section>
"""
        return

    headers = ["No.", "相対パス", "ファイル名", "拡張子", "状態"]
    for n in loc_names:
        headers.extend([f"{n}: サイズ", f"{n}: ハッシュ", f"{n}: 更新日時"])

    head_html = "".join(f"<th>{html.escape(h)}</th>" for h in headers)

    # 状態が 1 種類しかないセクション (各差分セクション) では状態フィルタは意味がない
    statuses = sorted({r.status for r in rows})
    if len(statuses) > 1:
        options = "".join(
            f'<option value="{html.escape(s, quote=True)}">{html.escape(s)}</option>'
            for s in statuses
        )
        status_select = (
            f'<select class="filter-status" aria-label="状態で絞り込み">'
            f'<option value="">すべての状態</option>{options}</select>'
        )
    else:
        status_select = ""

    yield f"""
<section id="{sid}">
  <h2>{html.escape(title)} ({len(rows):,} 件)</h2>
  <div class="table-tools">
    <input type="search" class="filter-input" placeholder="パスで絞り込み"
           aria-label="相対パスで絞り込み">
    {status_select}
    <span class="filter-count"></span>
  </div>
  <div class="scroll-wrap"><table class="fixed-col-table">
    <thead><tr>{head_html}</tr></thead>
    <tbody>"""

    for i, row in enumerate(rows, 1):
        # 表と詳細データで同じ文字列になるよう、どちらも _sanitize を通す
        relpath = _sanitize(row.relpath)
        name = relpath.rsplit("/", 1)[-1]
        ext = Path(name).suffix
        status_cls = _status_class(row.status)
        cells: List[str] = [
            f"<td class='num'>{i}</td>",
            f"<td title='{html.escape(relpath)}'>{html.escape(relpath)}</td>",
            f"<td>{html.escape(name)}</td>",
            f"<td>{html.escape(ext)}</td>",
            f"<td class='{status_cls}'>{html.escape(row.status)}</td>",
        ]
        # ハッシュ不一致時は最頻値以外のハッシュだけ強調
        # (ハッシュ未計算ならサイズ側を強調)
        minority, minority_size = _minority_targets(row)

        for loc in loc_names:
            err_msg = row.errors.get(loc)
            entry = row.entries.get(loc)
            if err_msg is not None:
                cells.append(f"<td class='cell-error'>{ERROR_PLACEHOLDER}</td>")
                err_msg = _sanitize(err_msg)
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
                size_cls = "num cell-mismatch" if entry.size in minority_size else "num"
                cells.append(f"<td class='{size_cls}'>{entry.size:,}</td>")
                if entry.hash is None:
                    cells.append(
                        f"<td class='hash muted' title='hash_mode=smart により省略'>"
                        f"{SKIPPED_HASH_PLACEHOLDER}</td>"
                    )
                else:
                    hash_cls = "hash"
                    if entry.hash in minority:
                        hash_cls += " cell-mismatch"
                    cells.append(
                        f"<td class='{hash_cls}' title='{html.escape(entry.hash)}'>"
                        f"{html.escape(entry.hash[:12])}…</td>"
                    )
                cells.append(f"<td>{html.escape(entry.mtime.strftime(DATETIME_FMT))}</td>")
        # 詳細は relpath をキーに JSON ブロック (#report-data) から JS が引く
        yield (
            f'<tr class="file-row" '
            f'data-relpath="{html.escape(relpath, quote=True)}">'
            + "".join(cells)
            + "</tr>"
        )

    yield """</tbody>
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
        # フォルダ行にはモーダルが無いので、省略せず全体を出す
        cells = [
            f"<td class='num'>{i}</td>",
            f"<td class='wrap-path'>{html.escape(_sanitize(d.relpath))}</td>",
        ]
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
            # エラー行にもモーダルが無いため、パスとメッセージは省略せず全体を出す
            rows.append(
                "<tr>"
                f"<td class='num'>{n}</td>"
                f"<td>{html.escape(_sanitize(s.location_name))}</td>"
                f"<td class='wrap-path'>{html.escape(_sanitize(e.relpath))}</td>"
                f"<td class='wrap-path'>{html.escape(_sanitize(e.message))}</td>"
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
