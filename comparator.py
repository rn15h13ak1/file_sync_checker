"""N 拠点間の差分検出。"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional

from scanner import FileEntry, ScanResult


# 状態ラベル
STATUS_OK = "OK"
STATUS_HASH_MISMATCH = "ハッシュ不一致"
STATUS_PARTIAL_MISSING = "欠落あり"
STATUS_PARTIAL_PRESENT = "一部のみ存在"


@dataclass
class FileRow:
    """全ファイル一覧シート1行分。

    location_name -> FileEntry (存在しない拠点は None) のマッピングを持つ。
    """

    relpath: str
    entries: Dict[str, Optional[FileEntry]]
    status: str


@dataclass
class DirRow:
    relpath: str
    presence: Dict[str, bool]  # location_name -> 存在するか


@dataclass
class ComparisonResult:
    location_names: List[str]
    all_files: List[FileRow]                    # 和集合
    hash_mismatches: List[FileRow] = field(default_factory=list)
    missing_files: List[FileRow] = field(default_factory=list)         # 一部の拠点に欠落
    extra_files: List[FileRow] = field(default_factory=list)           # 一部の拠点にのみ存在
    dir_diffs: List[DirRow] = field(default_factory=list)


def _classify(entries: Dict[str, Optional[FileEntry]]) -> str:
    present = [e for e in entries.values() if e is not None]
    n_total = len(entries)
    n_present = len(present)

    if n_present == n_total:
        # 全拠点に存在 → ハッシュで判定
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


def compare(scans: List[ScanResult]) -> ComparisonResult:
    """N 拠点のスキャン結果を比較する。"""
    if len(scans) < 2:
        raise ValueError("compare には2件以上のスキャン結果が必要です")

    location_names = [s.location_name for s in scans]

    # --- ファイル比較 ---
    all_relpaths = sorted({rp for s in scans for rp in s.files.keys()})

    all_files: List[FileRow] = []
    hash_mismatches: List[FileRow] = []
    missing_files: List[FileRow] = []
    extra_files: List[FileRow] = []

    for rel in all_relpaths:
        entries: Dict[str, Optional[FileEntry]] = {
            s.location_name: s.files.get(rel) for s in scans
        }
        status = _classify(entries)
        row = FileRow(relpath=rel, entries=entries, status=status)
        all_files.append(row)

        if status == STATUS_HASH_MISMATCH:
            hash_mismatches.append(row)
        elif status == STATUS_PARTIAL_MISSING:
            missing_files.append(row)
        elif status == STATUS_PARTIAL_PRESENT:
            extra_files.append(row)

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
        dir_diffs=dir_diffs,
    )
