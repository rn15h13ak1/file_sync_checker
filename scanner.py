"""フォルダスキャンとハッシュ計算。

スキャンは 2 フェーズに分かれる:

1. `stat_location` — 再帰列挙してサイズ・更新日時だけを集める (中身は読まない)。
2. `hash_locations` — 拠点をまたいで「ハッシュが必要なファイル」だけを読む。

ネットワーク共有では全ファイルの中身を読むことが支配的なコストになるため、
フェーズ 1 の情報だけで結論が出るファイル (サイズ相違 = 内容相違が確定、
サイズも更新日時も全拠点一致 = 同一とみなす) をフェーズ 2 から外せる。
どこまで外すかは `hash_mode` で選ぶ (`plan_hash_targets` 参照)。
"""
from __future__ import annotations

import fnmatch
import hashlib
import os
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Set, Tuple

from tqdm import tqdm


HASH_CHUNK_SIZE = 1024 * 1024  # 1 MiB

# 更新日時の一致判定に使う既定の許容誤差 (秒)。
# ファイルシステムによって更新日時の粒度が異なり (NTFS 100ns / HFS+ 1s / FAT 2s)、
# コピーの過程で丸められるため、厳密比較では同一ファイルでも不一致になる。
DEFAULT_MTIME_TOLERANCE_SEC = 2.0

HASH_MODE_ALWAYS = "always"
HASH_MODE_SMART = "smart"
HASH_MODES = (HASH_MODE_ALWAYS, HASH_MODE_SMART)


@dataclass(frozen=True)
class FileStat:
    """フェーズ 1 で集めるメタデータ。"""

    size: int
    mtime: datetime
    mtime_ns: int


@dataclass(frozen=True)
class FileEntry:
    size: int
    mtime: datetime
    # ハッシュ未計算 (フェーズ 2 から外れたファイル) では None。
    hash: Optional[str] = None


@dataclass(frozen=True)
class DirEntry:
    pass


@dataclass(frozen=True)
class ScanError:
    relpath: str
    message: str


@dataclass
class StatResult:
    """1 拠点分のフェーズ 1 結果。"""

    location_name: str
    root: Path
    stats: Dict[str, FileStat]
    dirs: Dict[str, DirEntry]
    errors: List[ScanError] = field(default_factory=list)
    file_errors: Dict[str, str] = field(default_factory=dict)


@dataclass
class ScanResult:
    """1 拠点分のスキャン結果。"""

    location_name: str
    root: Path
    files: Dict[str, FileEntry]
    dirs: Dict[str, DirEntry]
    errors: List[ScanError]
    # ファイル単位の読み取り失敗 (relpath -> message)。
    # walk 中のディレクトリエラーやルートエラーは含まない。
    file_errors: Dict[str, str] = field(default_factory=dict)
    # ハッシュを計算しなかったファイル数 (レポートのサマリー表示用)。
    skipped_hashes: int = 0


class ScanCancelled(Exception):
    """スキャンが中断された (Ctrl+C 等)。"""


@contextmanager
def _managed_pool(max_workers: int, cancel: threading.Event):
    """中断に反応するスレッドプール。

    `with ThreadPoolExecutor(...)` は終了時に `shutdown(wait=True)` するため、
    Ctrl+C を押してもキューに積んだ全ファイルのハッシュが終わるまで戻らない
    (10 万ファイル規模では事実上停止できない)。中断時は
    キューを破棄し、実行中のタスクにも `cancel` で降りてもらう。
    """
    executor = ThreadPoolExecutor(max_workers=max_workers)
    try:
        yield executor
    except BaseException:
        cancel.set()
        executor.shutdown(wait=False, cancel_futures=True)
        raise
    else:
        executor.shutdown(wait=True)


def _is_excluded(name: str, patterns: Iterable[str]) -> bool:
    return any(fnmatch.fnmatch(name, p) for p in patterns)


def _abspath(root: Path, relpath: str) -> Path:
    """'/' 区切りの相対パスから実パスを組み立てる。"""
    return root.joinpath(*relpath.split("/"))


def _walk_stats(
    root: Path,
    exclude_patterns: List[str],
    cancel: Optional[threading.Event] = None,
) -> Tuple[Dict[str, FileStat], List[str], List[ScanError], Dict[str, str]]:
    """ルート配下を再帰列挙し、サイズと更新日時を集める。

    `os.walk` ではなく `os.scandir` を直接使う。Windows では
    ディレクトリ列挙の時点でサイズ・更新日時が返るため `DirEntry.stat()` が
    追加のシステムコールを伴わず、SMB 共有でのフェーズ 1 がほぼ無コストになる。

    シンボリックリンクは `os.walk(followlinks=False)` と同じ扱いにする:
    リンク先がディレクトリならディレクトリとして記録するが、配下には降りない。
    """
    stats: Dict[str, FileStat] = {}
    dirs: List[str] = []
    errors: List[ScanError] = []
    file_errors: Dict[str, str] = {}

    stack: List[Tuple[Path, str]] = [(root, "")]
    while stack:
        if cancel is not None and cancel.is_set():
            raise ScanCancelled()
        current, prefix = stack.pop()
        try:
            with os.scandir(current) as it:
                entries = list(it)
        except OSError as e:
            errors.append(ScanError(relpath=prefix, message=f"walk error: {e}"))
            continue

        for de in entries:
            if _is_excluded(de.name, exclude_patterns):
                continue
            rel = f"{prefix}/{de.name}" if prefix else de.name
            try:
                if de.is_dir():
                    dirs.append(rel)
                    # シンボリックリンクのループを避けるため配下には降りない
                    if not de.is_symlink():
                        stack.append((Path(de.path), rel))
                    continue
                st = de.stat()
                stats[rel] = FileStat(
                    size=st.st_size,
                    mtime=datetime.fromtimestamp(st.st_mtime),
                    mtime_ns=st.st_mtime_ns,
                )
            except OSError as e:
                msg = f"{type(e).__name__}: {e}"
                file_errors[rel] = msg
                errors.append(ScanError(relpath=rel, message=msg))

    return stats, dirs, errors, file_errors


def stat_location(
    name: str,
    root: Path,
    *,
    exclude_patterns: List[str],
    cancel: Optional[threading.Event] = None,
) -> StatResult:
    """フェーズ 1: 1 拠点を列挙してメタデータだけを集める。"""
    if not root.exists():
        return StatResult(
            location_name=name, root=root, stats={}, dirs={},
            errors=[ScanError(relpath="", message=f"root path does not exist: {root}")],
        )
    if not root.is_dir():
        return StatResult(
            location_name=name, root=root, stats={}, dirs={},
            errors=[ScanError(relpath="", message=f"root path is not a directory: {root}")],
        )

    stats, dirs, errors, file_errors = _walk_stats(root, exclude_patterns, cancel)
    return StatResult(
        location_name=name,
        root=root,
        stats=stats,
        dirs={d: DirEntry() for d in dirs},
        errors=errors,
        file_errors=file_errors,
    )


def plan_hash_targets(
    stat_results: Sequence[StatResult],
    *,
    hash_mode: str = HASH_MODE_ALWAYS,
    mtime_tolerance_sec: float = DEFAULT_MTIME_TOLERANCE_SEC,
) -> Set[str]:
    """フェーズ 2 でハッシュを計算する相対パスの集合を返す。

    判定は相対パス単位で行い、「ある拠点だけハッシュがある」状態は作らない
    (比較器がサイズとハッシュを一貫して扱えるようにするため)。

    `always`: 全ファイル。ハッシュを常に取るので更新日時を一切信用しない。
    `smart`: 以下をフェーズ 2 から外す。
      - サイズが拠点間で相違 → 内容が違うことは確定しており、ハッシュは結論を変えない。
      - 存在する全拠点でサイズが一致し、更新日時も許容誤差内で一致 → 同一とみなす。
        更新日時に基づく推定であり、サイズを保ったまま書き換えて更新日時も復元された
        ファイルは検出できない。厳密さが要る場合は `always` を使う。
    どちらのモードでも、1 拠点にしか無いファイルはハッシュを計算する
    (比較には不要だが、レポート上のハッシュ欄を空にしないため)。
    """
    all_rel: Set[str] = set()
    for s in stat_results:
        all_rel |= s.stats.keys()

    if hash_mode == HASH_MODE_ALWAYS:
        return all_rel

    tolerance_ns = int(mtime_tolerance_sec * 1_000_000_000)
    targets: Set[str] = set()
    for rel in all_rel:
        present = [s.stats[rel] for s in stat_results if rel in s.stats]
        if len(present) < 2:
            targets.add(rel)
            continue
        if len({p.size for p in present}) > 1:
            continue  # サイズ相違 = 内容相違が確定
        mtimes = [p.mtime_ns for p in present]
        if max(mtimes) - min(mtimes) <= tolerance_ns:
            continue  # サイズ・更新日時とも一致 → 同一とみなす
        targets.add(rel)
    return targets


def _hash_file(
    path: Path, algorithm: str, cancel: Optional[threading.Event] = None
) -> str:
    h = hashlib.new(algorithm)
    with path.open("rb") as f:
        while True:
            # 大きいファイルの途中でも中断できるようチャンクごとに確認する
            if cancel is not None and cancel.is_set():
                raise ScanCancelled()
            chunk = f.read(HASH_CHUNK_SIZE)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()


def _hash_task(
    loc_index: int,
    relpath: str,
    abspath: Path,
    algorithm: str,
    cancel: Optional[threading.Event] = None,
) -> Tuple[int, str, Optional[str], Optional[str]]:
    """Returns: (loc_index, relpath, hash, error_message)"""
    try:
        return loc_index, relpath, _hash_file(abspath, algorithm, cancel), None
    except OSError as e:
        return loc_index, relpath, None, f"{type(e).__name__}: {e}"


def hash_locations(
    stat_results: Sequence[StatResult],
    targets: Set[str],
    *,
    algorithm: str,
    parallel_workers: int,
    show_progress: bool = True,
) -> List[ScanResult]:
    """フェーズ 2: 対象ファイルのハッシュを計算し、拠点ごとの ScanResult を組み立てる。

    スレッドプールは拠点をまたいで 1 つだけ使う。拠点はそれぞれ別のサーバなので、
    拠点ごとに直列化するより全体を 1 つのキューで回した方がネットワークを使い切れる。
    """
    tasks: List[Tuple[int, str, Path]] = []
    for i, s in enumerate(stat_results):
        for rel in s.stats:
            if rel in targets:
                tasks.append((i, rel, _abspath(s.root, rel)))

    hashes: Dict[Tuple[int, str], str] = {}
    hash_errors: Dict[Tuple[int, str], str] = {}

    def _accumulate(loc_index: int, rel: str, digest: Optional[str], err: Optional[str]) -> None:
        if digest is not None:
            hashes[(loc_index, rel)] = digest
        if err is not None:
            hash_errors[(loc_index, rel)] = err

    if tasks:
        desc = "hashing"
        cancel = threading.Event()
        if parallel_workers <= 1 or len(tasks) <= 1:
            iterator: Iterable = (
                _hash_task(i, rel, ap, algorithm) for i, rel, ap in tasks
            )
            if show_progress:
                iterator = tqdm(iterator, total=len(tasks), desc=desc, unit="file")
            for loc_index, rel, digest, err in iterator:
                _accumulate(loc_index, rel, digest, err)
        else:
            with _managed_pool(parallel_workers, cancel) as ex:
                futures = [
                    ex.submit(_hash_task, i, rel, ap, algorithm, cancel)
                    for i, rel, ap in tasks
                ]
                completed: Iterable = as_completed(futures)
                if show_progress:
                    completed = tqdm(completed, total=len(futures), desc=desc, unit="file")
                for fut in completed:
                    loc_index, rel, digest, err = fut.result()
                    _accumulate(loc_index, rel, digest, err)

    results: List[ScanResult] = []
    for i, s in enumerate(stat_results):
        files: Dict[str, FileEntry] = {}
        errors = list(s.errors)
        file_errors = dict(s.file_errors)
        skipped = 0
        for rel, st in s.stats.items():
            err = hash_errors.get((i, rel))
            if err is not None:
                # 読み取り失敗したファイルは files に載せず、エラーとして記録する
                file_errors[rel] = err
                errors.append(ScanError(relpath=rel, message=err))
                continue
            digest = hashes.get((i, rel))
            if digest is None and rel not in targets:
                skipped += 1
            files[rel] = FileEntry(size=st.size, mtime=st.mtime, hash=digest)
        results.append(
            ScanResult(
                location_name=s.location_name,
                root=s.root,
                files=files,
                dirs=s.dirs,
                errors=errors,
                file_errors=file_errors,
                skipped_hashes=skipped,
            )
        )
    return results


def scan_locations(
    locations: Sequence[Tuple[str, Path]],
    *,
    exclude_patterns: List[str],
    parallel_workers: int,
    hash_algorithm: str,
    hash_mode: str = HASH_MODE_ALWAYS,
    mtime_tolerance_sec: float = DEFAULT_MTIME_TOLERANCE_SEC,
    show_progress: bool = True,
    on_stat_done=None,
) -> List[ScanResult]:
    """全拠点を 2 フェーズでスキャンする。

    フェーズ 1 は拠点ごとに並列化する (拠点は別サーバなので待ち時間が重なる)。
    `on_stat_done(StatResult)` が渡されていれば、拠点の列挙が終わるたびに呼ぶ。
    """
    stat_results: List[Optional[StatResult]] = [None] * len(locations)
    if len(locations) <= 1:
        for i, (name, root) in enumerate(locations):
            sr = stat_location(name, root, exclude_patterns=exclude_patterns)
            stat_results[i] = sr
            if on_stat_done:
                on_stat_done(sr)
    else:
        cancel = threading.Event()
        with _managed_pool(len(locations), cancel) as ex:
            futures = {
                ex.submit(
                    stat_location, name, root,
                    exclude_patterns=exclude_patterns, cancel=cancel,
                ): i
                for i, (name, root) in enumerate(locations)
            }
            for fut in as_completed(futures):
                i = futures[fut]
                sr = fut.result()
                stat_results[i] = sr
                if on_stat_done:
                    on_stat_done(sr)

    stats: List[StatResult] = [s for s in stat_results if s is not None]
    targets = plan_hash_targets(
        stats, hash_mode=hash_mode, mtime_tolerance_sec=mtime_tolerance_sec
    )
    return hash_locations(
        stats,
        targets,
        algorithm=hash_algorithm,
        parallel_workers=parallel_workers,
        show_progress=show_progress,
    )


def scan_location(
    name: str,
    root: Path,
    *,
    exclude_patterns: List[str],
    parallel_workers: int,
    hash_algorithm: str,
    show_progress: bool = True,
) -> ScanResult:
    """1 拠点をスキャンして全ファイルのハッシュを計算する。

    単独拠点では拠点間の比較ができないため、常に全ファイルをハッシュする。
    """
    sr = stat_location(name, root, exclude_patterns=exclude_patterns)
    return hash_locations(
        [sr],
        set(sr.stats.keys()),
        algorithm=hash_algorithm,
        parallel_workers=parallel_workers,
        show_progress=show_progress,
    )[0]
