"""フォルダスキャンとハッシュ計算。"""
from __future__ import annotations

import fnmatch
import hashlib
import os
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path, PurePosixPath
from typing import Dict, Iterable, List, Optional, Tuple

from tqdm import tqdm


HASH_CHUNK_SIZE = 1024 * 1024  # 1 MiB


@dataclass(frozen=True)
class FileEntry:
    size: int
    mtime: datetime
    hash: str


@dataclass(frozen=True)
class DirEntry:
    pass


@dataclass(frozen=True)
class ScanError:
    relpath: str
    message: str


@dataclass
class ScanResult:
    """1 拠点分のスキャン結果。"""

    location_name: str
    root: Path
    files: Dict[str, FileEntry]
    dirs: Dict[str, DirEntry]
    errors: List[ScanError]


def _is_excluded(name: str, patterns: Iterable[str]) -> bool:
    return any(fnmatch.fnmatch(name, p) for p in patterns)


def _to_relpath(root: Path, target: Path) -> str:
    """ルートからの相対パスを '/' 区切りで返す。"""
    rel = target.relative_to(root)
    return str(PurePosixPath(*rel.parts))


def _walk(
    root: Path, exclude_patterns: List[str]
) -> Tuple[List[Tuple[str, Path]], List[str], List[ScanError]]:
    """ルート配下を再帰スキャン。

    Returns:
        files: [(relpath, abspath), ...]
        dirs:  [relpath, ...]
        errors: 走査中に発生したエラー
    """
    files: List[Tuple[str, Path]] = []
    dirs: List[str] = []
    errors: List[ScanError] = []

    try:
        # followlinks=False: シンボリックリンクループ防止
        walker = os.walk(root, followlinks=False, onerror=lambda e: errors.append(
            ScanError(relpath=_safe_relpath(root, e.filename), message=f"walk error: {e}")
        ))
        for dirpath, dirnames, filenames in walker:
            # ディレクトリ除外: 配下にも降りないように in-place で削除
            dirnames[:] = [d for d in dirnames if not _is_excluded(d, exclude_patterns)]

            dir_abs = Path(dirpath)
            if dir_abs != root:
                dirs.append(_to_relpath(root, dir_abs))

            for fname in filenames:
                if _is_excluded(fname, exclude_patterns):
                    continue
                fpath = dir_abs / fname
                files.append((_to_relpath(root, fpath), fpath))
    except OSError as e:
        errors.append(ScanError(relpath="", message=f"root walk failed: {e}"))

    return files, dirs, errors


def _safe_relpath(root: Path, target: Optional[str]) -> str:
    if not target:
        return ""
    try:
        return _to_relpath(root, Path(target))
    except Exception:
        return str(target)


def _hash_file(path: Path, algorithm: str) -> str:
    h = hashlib.new(algorithm)
    with path.open("rb") as f:
        while True:
            chunk = f.read(HASH_CHUNK_SIZE)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()


def _process_file(
    relpath: str, abspath: Path, algorithm: str
) -> Tuple[str, Optional[FileEntry], Optional[ScanError]]:
    try:
        stat = abspath.stat()
        digest = _hash_file(abspath, algorithm)
        return (
            relpath,
            FileEntry(
                size=stat.st_size,
                mtime=datetime.fromtimestamp(stat.st_mtime),
                hash=digest,
            ),
            None,
        )
    except (OSError, PermissionError) as e:
        return relpath, None, ScanError(relpath=relpath, message=f"{type(e).__name__}: {e}")


def scan_location(
    name: str,
    root: Path,
    *,
    exclude_patterns: List[str],
    parallel_workers: int,
    hash_algorithm: str,
    show_progress: bool = True,
) -> ScanResult:
    """1 拠点をスキャンしてハッシュを計算する。"""
    if not root.exists():
        return ScanResult(
            location_name=name,
            root=root,
            files={},
            dirs={},
            errors=[ScanError(relpath="", message=f"root path does not exist: {root}")],
        )
    if not root.is_dir():
        return ScanResult(
            location_name=name,
            root=root,
            files={},
            dirs={},
            errors=[ScanError(relpath="", message=f"root path is not a directory: {root}")],
        )

    files, dirs, errors = _walk(root, exclude_patterns)

    file_entries: Dict[str, FileEntry] = {}
    desc = f"[{name}] hashing"
    iterator: Iterable[Tuple[str, Optional[FileEntry], Optional[ScanError]]]

    if parallel_workers <= 1 or len(files) <= 1:
        iterator = (_process_file(rp, ap, hash_algorithm) for rp, ap in files)
        if show_progress:
            iterator = tqdm(iterator, total=len(files), desc=desc, unit="file")
        for relpath, entry, err in iterator:
            if entry is not None:
                file_entries[relpath] = entry
            if err is not None:
                errors.append(err)
    else:
        with ThreadPoolExecutor(max_workers=parallel_workers) as ex:
            futures = [ex.submit(_process_file, rp, ap, hash_algorithm) for rp, ap in files]
            progress = (
                tqdm(as_completed(futures), total=len(futures), desc=desc, unit="file")
                if show_progress
                else as_completed(futures)
            )
            for fut in progress:
                relpath, entry, err = fut.result()
                if entry is not None:
                    file_entries[relpath] = entry
                if err is not None:
                    errors.append(err)

    dir_entries = {d: DirEntry() for d in dirs}

    return ScanResult(
        location_name=name,
        root=root,
        files=file_entries,
        dirs=dir_entries,
        errors=errors,
    )
