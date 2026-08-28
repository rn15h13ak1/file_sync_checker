"""N 拠点間の差分検出。"""
from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Set

from scanner import FileEntry, ScanResult, match_key


# 状態ラベル
STATUS_OK = "OK"
STATUS_HASH_MISMATCH = "ハッシュ不一致"
STATUS_PARTIAL_MISSING = "欠落あり"
STATUS_PARTIAL_PRESENT = "一部のみ存在"
STATUS_ERROR = "エラー"  # 一部の拠点で読み取り失敗 — 欠落と区別する


@dataclass
class FileRow:
    """全ファイル一覧シート1行分。

    - relpath:  表示用の相対パス (NFC 正規化済み)。
    - entries:  location_name -> FileEntry。存在しない or 読み取り失敗の拠点は None。
    - errors:   location_name -> エラーメッセージ。エラーが無い拠点は None。
    - real_relpaths: location_name -> その拠点での実際の相対パス。
      Unicode 正規化形や大小文字が拠点ごとに違う場合、実ファイルを開くパスは
      拠点ごとに異なるため保持する (`scanner.match_key` 参照)。
    """

    relpath: str
    entries: Dict[str, Optional[FileEntry]]
    errors: Dict[str, Optional[str]]
    status: str
    real_relpaths: Dict[str, str] = field(default_factory=dict)


@dataclass
class DirRow:
    relpath: str
    presence: Dict[str, bool]  # location_name -> 存在するか


@dataclass
class ComparisonResult:
    location_names: List[str]
    all_files: List[FileRow]                                          # 和集合
    hash_mismatches: List[FileRow] = field(default_factory=list)
    missing_files: List[FileRow] = field(default_factory=list)        # 一部の拠点に欠落 (多数派)
    extra_files: List[FileRow] = field(default_factory=list)          # 一部の拠点のみ存在 (少数派)
    errored_files: List[FileRow] = field(default_factory=list)        # いずれかの拠点で読み取り失敗
    dir_diffs: List[DirRow] = field(default_factory=list)


def _minority(values: List) -> Set:
    """「過半数（最頻値）と異なる」値の集合を返す。

    最頻値が複数（タイ）の場合は全値を返す（どれが「正」か判定不能なので全てハイライト）。
    値が2つ未満や全部同一の場合は空集合を返す。
    """
    if len(values) < 2:
        return set()
    counts = Counter(values)
    if len(counts) <= 1:
        return set()  # 全部同じ
    max_count = max(counts.values())
    leaders = [v for v, c in counts.items() if c == max_count]
    if len(leaders) > 1:
        # 最頻値が複数 → 判定不能なので全ハイライト
        return set(counts.keys())
    return {v for v, c in counts.items() if c < max_count}


def minority_hashes(entries: Dict[str, Optional[FileEntry]]) -> Set[str]:
    """ハッシュ不一致状態の中で「過半数（最頻値）と異なる」ハッシュ集合を返す。"""
    hashes = [e.hash for e in entries.values() if e is not None and e.hash is not None]
    return _minority(hashes)


def minority_sizes(entries: Dict[str, Optional[FileEntry]]) -> Set[int]:
    """「過半数（最頻値）と異なる」サイズ集合を返す。

    ハッシュ未計算（サイズ相違だけで不一致が確定したケース）で、
    どの拠点が少数派かを示すために使う。
    """
    return _minority([e.size for e in entries.values() if e is not None])


def _classify(
    entries: Dict[str, Optional[FileEntry]],
    errors: Dict[str, Optional[str]],
) -> str:
    # いずれかの拠点で読み取り失敗があれば、まずエラーとして分離する（欠落と区別）。
    if any(e is not None for e in errors.values()):
        return STATUS_ERROR

    present = [e for e in entries.values() if e is not None]
    n_total = len(entries)
    n_present = len(present)

    if n_present == n_total:
        # 全拠点に存在 → サイズ・ハッシュで判定。
        # サイズが違えば内容も違うため、ハッシュ未計算 (hash=None) でも不一致と判定できる。
        if len({e.size for e in present}) > 1:
            return STATUS_HASH_MISMATCH
        # ハッシュ計算をスキップしたファイルは全拠点 None になり、同一とみなされる
        # (どの拠点でハッシュを取るかは相対パス単位で決まる — scanner.plan_hash_targets)。
        hashes = {e.hash for e in present}
        return STATUS_OK if len(hashes) == 1 else STATUS_HASH_MISMATCH

    if n_present == 0:
        # 理論上ありえないが念のため
        return STATUS_PARTIAL_MISSING

    # 一部の拠点に存在
    # 「過半数に存在 = 欠落あり」「少数に存在 = 一部のみ存在」とラベルを分ける
    if n_present > n_total / 2:
        return STATUS_PARTIAL_MISSING
    return STATUS_PARTIAL_PRESENT


def _display_relpath(
    key: str, real_relpaths: Dict[str, str], location_names: List[str]
) -> str:
    """レポートに表示する相対パスを決める。

    照合キーは大小文字を潰していることがあるため、そのままでは表示に使えない。
    設定順で最初に見つかった拠点の実際のパスを NFC 正規化して使う
    (どの拠点も持っていなければ照合キーにフォールバック)。
    """
    for name in location_names:
        real = real_relpaths.get(name)
        if real is not None:
            return match_key(real, normalize_unicode=True, case_sensitive=True)
    return key


def compare(scans: List[ScanResult]) -> ComparisonResult:
    """N 拠点のスキャン結果を比較する。"""
    if len(scans) < 2:
        raise ValueError("compare には2件以上のスキャン結果が必要です")

    location_names = [s.location_name for s in scans]

    # --- ファイル比較 ---
    # 和集合: 「ファイルが見つかった拠点」+「読み取り失敗した拠点」両方を含める。
    # キーは照合キー (`scanner.match_key`) なので、Unicode 正規化形や大小文字が
    # 拠点ごとに違っても同一ファイルとして 1 行にまとまる。
    all_keys = sorted(
        {rp for s in scans for rp in s.files.keys()}
        | {rp for s in scans for rp in s.file_errors.keys()}
    )

    all_files: List[FileRow] = []
    hash_mismatches: List[FileRow] = []
    missing_files: List[FileRow] = []
    extra_files: List[FileRow] = []
    errored_files: List[FileRow] = []

    for key in all_keys:
        entries: Dict[str, Optional[FileEntry]] = {
            s.location_name: s.files.get(key) for s in scans
        }
        errors: Dict[str, Optional[str]] = {
            s.location_name: s.file_errors.get(key) for s in scans
        }
        real_relpaths: Dict[str, str] = {
            s.location_name: s.real_relpaths[key]
            for s in scans
            if key in s.real_relpaths
        }
        status = _classify(entries, errors)
        row = FileRow(
            relpath=_display_relpath(key, real_relpaths, location_names),
            entries=entries,
            errors=errors,
            status=status,
            real_relpaths=real_relpaths,
        )
        all_files.append(row)

        if status == STATUS_HASH_MISMATCH:
            hash_mismatches.append(row)
        elif status == STATUS_PARTIAL_MISSING:
            missing_files.append(row)
        elif status == STATUS_PARTIAL_PRESENT:
            extra_files.append(row)
        elif status == STATUS_ERROR:
            errored_files.append(row)

    # --- ディレクトリ比較 ---
    all_dirs = sorted({d for s in scans for d in s.dirs.keys()})
    dir_diffs: List[DirRow] = []
    for d in all_dirs:
        presence = {s.location_name: (d in s.dirs) for s in scans}
        if not all(presence.values()):
            dir_diffs.append(DirRow(relpath=d, presence=presence))

    return ComparisonResult(
        location_names=location_names,
        all_files=all_files,
        hash_mismatches=hash_mismatches,
        missing_files=missing_files,
        extra_files=extra_files,
        errored_files=errored_files,
        dir_diffs=dir_diffs,
    )
