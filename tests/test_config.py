"""config: YAML 検証ルール（重複拒否・相対パス解決など）。"""
from __future__ import annotations

from pathlib import Path

import pytest

from config import ConfigError, load_config


def _write(path: Path, body: str) -> Path:
    path.write_text(body, encoding="utf-8")
    return path


class TestLoading:
    def test_missing_file(self, tmp_path: Path):
        with pytest.raises(ConfigError, match="設定ファイルが見つかりません"):
            load_config(tmp_path / "nope.yaml")

    def test_minimal_valid_config(self, tmp_path: Path):
        cfg = _write(tmp_path / "c.yaml", """
locations:
  - name: A
    path: /opt/a
  - name: B
    path: /opt/b
""")
        config = load_config(cfg)
        assert len(config.locations) == 2
        assert config.locations[0].name == "A"
        assert config.output.format == "excel"  # default


class TestValidation:
    def test_single_location_rejected(self, tmp_path: Path):
        cfg = _write(tmp_path / "c.yaml", """
locations:
  - name: A
    path: /opt/a
""")
        with pytest.raises(ConfigError, match="2件以上"):
            load_config(cfg)

    def test_duplicate_name_rejected(self, tmp_path: Path):
        cfg = _write(tmp_path / "c.yaml", """
locations:
  - name: A
    path: /opt/a
  - name: A
    path: /opt/b
""")
        with pytest.raises(ConfigError, match="拠点名が重複"):
            load_config(cfg)

    def test_duplicate_path_rejected(self, tmp_path: Path):
        """N: 同じ実体パスを2拠点に登録すると弾かれる。"""
        cfg = _write(tmp_path / "c.yaml", """
locations:
  - name: A
    path: /opt/shared
  - name: B
    path: /opt/shared
""")
        with pytest.raises(ConfigError, match="重複"):
            load_config(cfg)

    def test_duplicate_path_with_trailing_slash_rejected(self, tmp_path: Path):
        """末尾スラッシュの違いは normpath で同一視される。"""
        cfg = _write(tmp_path / "c.yaml", """
locations:
  - name: A
    path: /opt/shared
  - name: B
    path: /opt/shared/
""")
        with pytest.raises(ConfigError, match="重複"):
            load_config(cfg)

    def test_invalid_format_rejected(self, tmp_path: Path):
        cfg = _write(tmp_path / "c.yaml", """
locations:
  - name: A
    path: /opt/a
  - name: B
    path: /opt/b
output:
  format: pdf
""")
        with pytest.raises(ConfigError, match="output.format"):
            load_config(cfg)

    def test_invalid_workers_rejected(self, tmp_path: Path):
        cfg = _write(tmp_path / "c.yaml", """
locations:
  - name: A
    path: /opt/a
  - name: B
    path: /opt/b
performance:
  parallel_workers: 0
""")
        with pytest.raises(ConfigError, match="parallel_workers"):
            load_config(cfg)

    def test_invalid_hash_algo_rejected(self, tmp_path: Path):
        cfg = _write(tmp_path / "c.yaml", """
locations:
  - name: A
    path: /opt/a
  - name: B
    path: /opt/b
performance:
  hash_algorithm: md5
""")
        with pytest.raises(ConfigError, match="hash_algorithm"):
            load_config(cfg)

    def test_hash_mode_defaults_to_always(self, tmp_path: Path):
        """既定では更新日時を信用せず全ファイルをハッシュする。"""
        cfg = _write(tmp_path / "c.yaml", """
locations:
  - name: A
    path: /opt/a
  - name: B
    path: /opt/b
""")
        config = load_config(cfg)
        assert config.performance.hash_mode == "always"
        assert config.performance.mtime_tolerance_sec == 2.0

    def test_hash_mode_smart_accepted(self, tmp_path: Path):
        cfg = _write(tmp_path / "c.yaml", """
locations:
  - name: A
    path: /opt/a
  - name: B
    path: /opt/b
performance:
  hash_mode: smart
  mtime_tolerance_sec: 0.5
""")
        config = load_config(cfg)
        assert config.performance.hash_mode == "smart"
        assert config.performance.mtime_tolerance_sec == 0.5

    def test_invalid_hash_mode_rejected(self, tmp_path: Path):
        cfg = _write(tmp_path / "c.yaml", """
locations:
  - name: A
    path: /opt/a
  - name: B
    path: /opt/b
performance:
  hash_mode: trust_me
""")
        with pytest.raises(ConfigError, match="hash_mode"):
            load_config(cfg)

    def test_negative_mtime_tolerance_rejected(self, tmp_path: Path):
        cfg = _write(tmp_path / "c.yaml", """
locations:
  - name: A
    path: /opt/a
  - name: B
    path: /opt/b
performance:
  mtime_tolerance_sec: -1
""")
        with pytest.raises(ConfigError, match="mtime_tolerance_sec"):
            load_config(cfg)

    def test_matching_defaults(self, tmp_path: Path):
        """既定で Unicode 正規化あり・大小文字は区別する。"""
        cfg = _write(tmp_path / "c.yaml", """
locations:
  - name: A
    path: /opt/a
  - name: B
    path: /opt/b
""")
        config = load_config(cfg)
        assert config.matching.normalize_unicode is True
        assert config.matching.case_sensitive is True

    def test_matching_options_parsed(self, tmp_path: Path):
        cfg = _write(tmp_path / "c.yaml", """
locations:
  - name: A
    path: /opt/a
  - name: B
    path: /opt/b
matching:
  normalize_unicode: false
  case_sensitive: false
""")
        config = load_config(cfg)
        assert config.matching.normalize_unicode is False
        assert config.matching.case_sensitive is False

    def test_matching_non_bool_rejected(self, tmp_path: Path):
        cfg = _write(tmp_path / "c.yaml", """
locations:
  - name: A
    path: /opt/a
  - name: B
    path: /opt/b
matching:
  case_sensitive: "yes"
""")
        with pytest.raises(ConfigError, match="case_sensitive"):
            load_config(cfg)

    def test_backslash_path_rejected(self, tmp_path: Path):
        """バックスラッシュ表記 (\\\\server\\share) は YAML パース時に潰れて壊れがちなので拒否。"""
        cfg = _write(tmp_path / "c.yaml", """
locations:
  - name: A
    path: "\\\\\\\\server-a\\\\share\\\\docs"
  - name: B
    path: /opt/b
""")
        with pytest.raises(ConfigError, match="バックスラッシュ"):
            load_config(cfg)

    def test_backslash_anywhere_in_path_rejected(self, tmp_path: Path):
        """混在パスも拒否 (フォワードスラッシュ統一の方針)。"""
        cfg = _write(tmp_path / "c.yaml", """
locations:
  - name: A
    path: "//server-a/share\\\\docs"
  - name: B
    path: /opt/b
""")
        with pytest.raises(ConfigError, match="バックスラッシュ"):
            load_config(cfg)

    def test_missing_path_in_location(self, tmp_path: Path):
        cfg = _write(tmp_path / "c.yaml", """
locations:
  - name: A
  - name: B
    path: /opt/b
""")
        with pytest.raises(ConfigError, match="name と path"):
            load_config(cfg)


class TestRelativePathResolution:
    def test_relative_location_path_resolved_against_config_dir(self, tmp_path: Path):
        # config.yaml の隣に locA, locB を作る
        (tmp_path / "locA").mkdir()
        (tmp_path / "locB").mkdir()
        cfg = _write(tmp_path / "c.yaml", """
locations:
  - name: A
    path: ./locA
  - name: B
    path: ./locB
""")
        config = load_config(cfg)
        assert config.locations[0].path == (tmp_path / "locA").resolve()
        assert config.locations[1].path == (tmp_path / "locB").resolve()

    def test_relative_output_dir_resolved_against_config_dir(self, tmp_path: Path):
        cfg = _write(tmp_path / "c.yaml", """
locations:
  - name: A
    path: /opt/a
  - name: B
    path: /opt/b
output:
  output_dir: ./out
""")
        config = load_config(cfg)
        assert config.output.output_dir == (tmp_path / "out").resolve()

    def test_absolute_paths_preserved(self, tmp_path: Path):
        cfg = _write(tmp_path / "c.yaml", """
locations:
  - name: A
    path: /opt/absolute/a
  - name: B
    path: /opt/absolute/b
output:
  output_dir: /var/reports
""")
        config = load_config(cfg)
        assert config.locations[0].path == Path("/opt/absolute/a")
        assert config.output.output_dir == Path("/var/reports")
