"""設定ファイル(YAML)の読み込みと検証。"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import List

import yaml


SUPPORTED_FORMATS = {"excel", "html", "both"}
SUPPORTED_HASH_ALGOS = {"sha256"}


@dataclass(frozen=True)
class Location:
    name: str
    path: Path


@dataclass(frozen=True)
class OutputConfig:
    format: str
    output_dir: Path


@dataclass(frozen=True)
class PerformanceConfig:
    parallel_workers: int
    hash_algorithm: str


@dataclass(frozen=True)
class Config:
    locations: List[Location]
    exclude_patterns: List[str] = field(default_factory=list)
    output: OutputConfig = field(default_factory=lambda: OutputConfig("excel", Path("./reports")))
    performance: PerformanceConfig = field(
        default_factory=lambda: PerformanceConfig(4, "sha256")
    )


class ConfigError(ValueError):
    pass


def load_config(path: str | Path) -> Config:
    config_path = Path(path)
    if not config_path.is_file():
        raise ConfigError(f"設定ファイルが見つかりません: {config_path}")

    with config_path.open("r", encoding="utf-8") as f:
        raw = yaml.safe_load(f) or {}

    locations_raw = raw.get("locations")
    if not locations_raw or not isinstance(locations_raw, list):
        raise ConfigError("locations は1件以上指定してください")

    locations: List[Location] = []
    seen_names: set[str] = set()
    for i, item in enumerate(locations_raw):
        if not isinstance(item, dict):
            raise ConfigError(f"locations[{i}] が不正です")
        name = item.get("name")
        path = item.get("path")
        if not name or not path:
            raise ConfigError(f"locations[{i}] は name と path が必須です")
        if name in seen_names:
            raise ConfigError(f"拠点名が重複しています: {name}")
        seen_names.add(name)
        locations.append(Location(name=str(name), path=Path(str(path))))

    if len(locations) < 2:
        raise ConfigError("locations は2件以上必要です")

    exclude_patterns = raw.get("exclude_patterns") or []
    if not isinstance(exclude_patterns, list):
        raise ConfigError("exclude_patterns はリスト形式で指定してください")
    exclude_patterns = [str(p) for p in exclude_patterns]

    output_raw = raw.get("output") or {}
    fmt = str(output_raw.get("format", "excel")).lower()
    if fmt not in SUPPORTED_FORMATS:
        raise ConfigError(
            f"output.format は {sorted(SUPPORTED_FORMATS)} のいずれかを指定してください (現在: {fmt})"
        )
    output_dir = Path(str(output_raw.get("output_dir", "./reports")))
    output = OutputConfig(format=fmt, output_dir=output_dir)

    perf_raw = raw.get("performance") or {}
    workers = int(perf_raw.get("parallel_workers", 4))
    if workers < 1:
        raise ConfigError("performance.parallel_workers は1以上を指定してください")
    algo = str(perf_raw.get("hash_algorithm", "sha256")).lower()
    if algo not in SUPPORTED_HASH_ALGOS:
        raise ConfigError(
            f"performance.hash_algorithm は {sorted(SUPPORTED_HASH_ALGOS)} のみ対応 (現在: {algo})"
        )
    performance = PerformanceConfig(parallel_workers=workers, hash_algorithm=algo)

    return Config(
        locations=locations,
        exclude_patterns=exclude_patterns,
        output=output,
        performance=performance,
    )
