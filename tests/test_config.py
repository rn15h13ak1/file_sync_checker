"""config: YAML 検証ルール（重複拒否・相対パス解決など）。"""
from __future__ import annotations

from pathlib import Path

import pytest

from config import ConfigError, apply_overrides, load_config


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

    def test_max_table_rows_defaults_when_absent(self, tmp_path: Path):
        """既に配布済みの config.yaml (項目なし) でも既定値で動く。"""
        cfg = _write(tmp_path / "c.yaml", """
locations:
  - name: A
    path: /opt/a
  - name: B
    path: /opt/b
""")
        assert load_config(cfg).output.max_table_rows == 20_000

    def test_max_table_rows_explicit(self, tmp_path: Path):
        cfg = _write(tmp_path / "c.yaml", """
locations:
  - name: A
    path: /opt/a
  - name: B
    path: /opt/b
output:
  max_table_rows: 5000
""")
        assert load_config(cfg).output.max_table_rows == 5000

    def test_max_table_rows_zero_means_unlimited(self, tmp_path: Path):
        cfg = _write(tmp_path / "c.yaml", """
locations:
  - name: A
    path: /opt/a
  - name: B
    path: /opt/b
output:
  max_table_rows: 0
""")
        assert load_config(cfg).output.max_table_rows == 0

    def test_negative_max_table_rows_rejected(self, tmp_path: Path):
        cfg = _write(tmp_path / "c.yaml", """
locations:
  - name: A
    path: /opt/a
  - name: B
    path: /opt/b
output:
  max_table_rows: -1
""")
        with pytest.raises(ConfigError, match="max_table_rows"):
            load_config(cfg)

    def test_non_integer_max_table_rows_rejected(self, tmp_path: Path):
        cfg = _write(tmp_path / "c.yaml", """
locations:
  - name: A
    path: /opt/a
  - name: B
    path: /opt/b
output:
  max_table_rows: "たくさん"
""")
        with pytest.raises(ConfigError, match="max_table_rows"):
            load_config(cfg)

    def test_cli_overrides_keep_max_table_rows(self, tmp_path: Path):
        """--format や -o で上書きしても表示上限は維持される。"""
        cfg = _write(tmp_path / "c.yaml", """
locations:
  - name: A
    path: /opt/a
  - name: B
    path: /opt/b
output:
  max_table_rows: 1234
""")
        config = apply_overrides(load_config(cfg), output_format="html")
        assert config.output.max_table_rows == 1234

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

    def test_missing_locations_rejected(self, tmp_path: Path):
        cfg = _write(tmp_path / "c.yaml", "output:\n  format: html\n")
        with pytest.raises(ConfigError, match="locations"):
            load_config(cfg)

    def test_locations_not_a_list_rejected(self, tmp_path: Path):
        cfg = _write(tmp_path / "c.yaml", "locations: not-a-list\n")
        with pytest.raises(ConfigError, match="locations"):
            load_config(cfg)

    def test_location_entry_not_a_mapping_rejected(self, tmp_path: Path):
        cfg = _write(tmp_path / "c.yaml", """
locations:
  - "/opt/a"
  - name: B
    path: /opt/b
""")
        with pytest.raises(ConfigError, match=r"locations\[0\]"):
            load_config(cfg)

    def test_exclude_patterns_not_a_list_rejected(self, tmp_path: Path):
        cfg = _write(tmp_path / "c.yaml", """
locations:
  - name: A
    path: /opt/a
  - name: B
    path: /opt/b
exclude_patterns: "*.tmp"
""")
        with pytest.raises(ConfigError, match="exclude_patterns"):
            load_config(cfg)

    def test_non_numeric_mtime_tolerance_rejected(self, tmp_path: Path):
        cfg = _write(tmp_path / "c.yaml", """
locations:
  - name: A
    path: /opt/a
  - name: B
    path: /opt/b
performance:
  mtime_tolerance_sec: "だいたい2秒"
""")
        with pytest.raises(ConfigError, match="mtime_tolerance_sec"):
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


class TestCliOverrides:
    def _config(self, tmp_path: Path):
        cfg = _write(tmp_path / "c.yaml", """
locations:
  - name: A
    path: /opt/a
  - name: B
    path: /opt/b
output:
  format: excel
  output_dir: ./reports
performance:
  hash_mode: always
""")
        return load_config(cfg)

    def test_no_overrides_keeps_config(self, tmp_path: Path):
        config = self._config(tmp_path)
        assert apply_overrides(config) == config

    def test_format_and_hash_mode_overridden(self, tmp_path: Path):
        config = apply_overrides(
            self._config(tmp_path), output_format="html", hash_mode="smart"
        )
        assert config.output.format == "html"
        assert config.performance.hash_mode == "smart"
        # 他の項目は維持される
        assert [loc.name for loc in config.locations] == ["A", "B"]

    def test_output_dir_resolved_from_cwd(self, tmp_path: Path, monkeypatch):
        """設定ファイル内の相対パスと違い、CLI 指定は CWD 基準で解決する。"""
        monkeypatch.chdir(tmp_path)
        config = apply_overrides(self._config(tmp_path), output_dir="./out")
        assert config.output.output_dir == (tmp_path / "out").resolve()

    def test_invalid_format_rejected(self, tmp_path: Path):
        with pytest.raises(ConfigError, match="--format"):
            apply_overrides(self._config(tmp_path), output_format="pdf")

    def test_invalid_hash_mode_rejected(self, tmp_path: Path):
        with pytest.raises(ConfigError, match="--hash-mode"):
            apply_overrides(self._config(tmp_path), hash_mode="trust_me")


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


class TestComparisons:
    """複数の比較を定義する `comparisons:` 形式。"""

    def _write(self, tmp_path: Path, body: str) -> Path:
        for d in ("a", "b", "c", "d"):
            (tmp_path / d).mkdir(exist_ok=True)
        return _write(tmp_path / "c.yaml", body)

    def _two(self, tmp_path: Path) -> Path:
        return self._write(tmp_path, f"""
comparisons:
  契約書:
    - {{name: 本社, path: {tmp_path}/a}}
    - {{name: 大阪, path: {tmp_path}/b}}
  設計:
    - {{name: 本社, path: {tmp_path}/c}}
    - {{name: 大阪, path: {tmp_path}/d}}
""")

    def test_legacy_locations_still_work(self, tmp_path: Path):
        """従来形式はそのまま動く (既に配布済みのため)。"""
        cfg = self._write(tmp_path, f"""
locations:
  - {{name: A, path: {tmp_path}/a}}
  - {{name: B, path: {tmp_path}/b}}
""")
        config = load_config(cfg)
        assert config.selected == ""
        assert config.comparison_names == []
        assert [loc.name for loc in config.locations] == ["A", "B"]

    def test_selects_the_named_comparison(self, tmp_path: Path):
        config = load_config(self._two(tmp_path), comparison="設計")
        assert config.selected == "設計"
        assert config.locations[0].path == tmp_path / "c"

    def test_lists_the_defined_names(self, tmp_path: Path):
        config = load_config(self._two(tmp_path), comparison="契約書")
        assert config.comparison_names == ["契約書", "設計"]

    def test_omitting_the_name_errors_with_candidates(self, tmp_path: Path):
        with pytest.raises(ConfigError, match="契約書, 設計"):
            load_config(self._two(tmp_path))

    def test_unknown_name_errors_with_candidates(self, tmp_path: Path):
        with pytest.raises(ConfigError, match="経理"):
            load_config(self._two(tmp_path), comparison="経理")

    def test_single_comparison_needs_no_name(self, tmp_path: Path):
        cfg = self._write(tmp_path, f"""
comparisons:
  契約書:
    - {{name: A, path: {tmp_path}/a}}
    - {{name: B, path: {tmp_path}/b}}
""")
        assert load_config(cfg).selected == "契約書"

    def test_listing_does_not_require_a_selection(self, tmp_path: Path):
        """一覧表示のために名前が要る、という状態を避ける。"""
        config = load_config(self._two(tmp_path), require_selection=False)
        assert config.comparison_names == ["契約書", "設計"]

    def test_locations_and_comparisons_are_exclusive(self, tmp_path: Path):
        cfg = self._write(tmp_path, f"""
locations:
  - {{name: A, path: {tmp_path}/a}}
  - {{name: B, path: {tmp_path}/b}}
comparisons:
  x:
    - {{name: A, path: {tmp_path}/c}}
    - {{name: B, path: {tmp_path}/d}}
""")
        with pytest.raises(ConfigError, match="同時に指定できません"):
            load_config(cfg)

    def test_each_comparison_needs_two_locations(self, tmp_path: Path):
        cfg = self._write(tmp_path, f"""
comparisons:
  x:
    - {{name: A, path: {tmp_path}/a}}
""")
        with pytest.raises(ConfigError, match="comparisons.x"):
            load_config(cfg, comparison="x")

    def test_duplicate_paths_are_rejected_within_a_comparison(self, tmp_path: Path):
        cfg = self._write(tmp_path, f"""
comparisons:
  x:
    - {{name: A, path: {tmp_path}/a}}
    - {{name: B, path: {tmp_path}/a}}
""")
        with pytest.raises(ConfigError, match="重複"):
            load_config(cfg, comparison="x")

    def test_same_path_may_appear_in_different_comparisons(self, tmp_path: Path):
        """比較をまたいで同じ拠点を使うのは正当なので弾かない。"""
        cfg = self._write(tmp_path, f"""
comparisons:
  x:
    - {{name: A, path: {tmp_path}/a}}
    - {{name: B, path: {tmp_path}/b}}
  y:
    - {{name: A, path: {tmp_path}/a}}
    - {{name: C, path: {tmp_path}/c}}
""")
        assert load_config(cfg, comparison="y").locations[0].path == tmp_path / "a"

    @pytest.mark.parametrize("name", ["a/b", "a:b", "a*b", "a?b", 'a"b', "a<b", "a|b"])
    def test_names_with_filename_unsafe_characters_are_rejected(
        self, tmp_path: Path, name: str
    ):
        """比較名はレポートのファイル名に入るため、使えない文字を弾く。"""
        # 名前は YAML のシングルクォートで囲む (ダブルクォートを含む名前も試すため)
        cfg = self._write(tmp_path, f"""
comparisons:
  '{name}':
    - {{name: A, path: {tmp_path}/a}}
    - {{name: B, path: {tmp_path}/b}}
""")
        with pytest.raises(ConfigError, match="使えない文字"):
            load_config(cfg, comparison=name)

    def test_empty_comparisons_is_rejected(self, tmp_path: Path):
        cfg = self._write(tmp_path, "comparisons: {}\n")
        with pytest.raises(ConfigError, match="comparisons"):
            load_config(cfg)
