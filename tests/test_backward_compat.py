"""配布済みの設定ファイルが動き続けることを守るテスト。

このツールは既に利用者へ配布されている。手元の `config.yaml` は
トップレベルに `locations:` を書く形式で、書き換えを求めずに
新しい版へ入れ替えられる必要がある。

`comparisons:` を追加したあとも従来形式が同じ結果になることを、
実際に `main.py` をプロセスとして起動して確認する。
"""
from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

import pytest
import yaml

REPO_ROOT = Path(__file__).resolve().parent.parent
MAIN = REPO_ROOT / "main.py"
MENU = REPO_ROOT / "menu.py"

# 配布済みの config.yaml と同じ形。ここは「変えてはいけない入力」なので、
# ヘルパーで組み立てずに YAML をそのまま書いておく。
LEGACY_CONFIG = """
locations:
  - name: "拠点A"
    path: "{a}"
  - name: "拠点B"
    path: "{b}"

exclude_patterns:
  - "~$*"
  - "*.tmp"
  - "Thumbs.db"
  - ".DS_Store"
  - "*.lnk"

output:
  format: "both"
  output_dir: "{out}"

performance:
  parallel_workers: 4
  hash_algorithm: "sha256"
"""


@pytest.fixture
def legacy(tmp_path: Path):
    """従来形式の設定と、一致する 2 拠点を用意する。"""
    a, b = tmp_path / "拠点A", tmp_path / "拠点B"
    for d in (a, b):
        d.mkdir()
        (d / "資料.docx").write_text("same", encoding="utf-8")
    cfg = tmp_path / "config.yaml"
    cfg.write_text(
        LEGACY_CONFIG.format(a=a, b=b, out=tmp_path / "reports"), encoding="utf-8"
    )
    return cfg, tmp_path / "reports"


def _run(*args: str) -> subprocess.CompletedProcess:
    return subprocess.run(
        [sys.executable, str(MAIN), *args],
        capture_output=True, text=True, timeout=120,
    )


class TestLegacyConfigStillRuns:
    def test_runs_without_any_new_option(self, legacy):
        """比較名を指定しなくても実行できる。"""
        cfg, _ = legacy
        r = _run("-c", str(cfg), "--no-progress")
        assert r.returncode == 0, r.stderr

    def test_detects_differences(self, legacy):
        cfg, _ = legacy
        (cfg.parent / "拠点A" / "only_a.txt").write_text("x", encoding="utf-8")
        r = _run("-c", str(cfg), "--no-progress")
        assert r.returncode == 1, r.stderr

    def test_report_names_are_unchanged(self, legacy):
        """レポート名に比較名が入らない (ブックマークが切れない)。"""
        cfg, out = legacy
        _run("-c", str(cfg), "--no-progress")

        names = sorted(p.name for p in out.iterdir())
        # 固定名は sync-check.html のまま
        assert "sync-check.html" in names
        # タイムスタンプ付きは sync-check-YYYYMMDD-HHMMSS.{html,xlsx}
        stamped = [n for n in names if n != "sync-check.html"]
        assert stamped, names
        for n in stamped:
            assert re.fullmatch(r"sync-check-\d{8}-\d{6}\.(html|xlsx)", n), n

    def test_report_has_no_comparison_row(self, legacy):
        """実行条件に「比較」の行を出さない (名前が無いため)。"""
        cfg, out = legacy
        _run("-c", str(cfg), "--no-progress")

        html = (out / "sync-check.html").read_text(encoding="utf-8")
        conditions = html.split("<h3>実行条件</h3>")[1].split("</table>")[0]
        assert "<td>比較</td>" not in conditions

    def test_list_comparisons_says_there_are_none(self, legacy):
        cfg, _ = legacy
        r = _run("-c", str(cfg), "--list-comparisons")
        assert r.returncode == 0
        assert "locations" in r.stdout

    def test_comparison_option_is_rejected_with_guidance(self, legacy):
        """従来形式に比較名を渡したら、理由が分かるエラーにする。"""
        cfg, _ = legacy
        r = _run("-c", str(cfg), "--no-progress", "--comparison", "契約書")
        assert r.returncode == 2
        assert "契約書" in r.stderr

    def test_old_report_pruning_still_works(self, legacy):
        """従来形式のレポートも keep_reports で整理できる。"""
        cfg, out = legacy
        data = yaml.safe_load(cfg.read_text(encoding="utf-8"))
        data["output"]["keep_reports"] = 1
        cfg.write_text(yaml.safe_dump(data, allow_unicode=True), encoding="utf-8")

        out.mkdir(exist_ok=True)
        for slug in ("20250101-090000", "20250102-090000"):
            (out / f"sync-check-{slug}.html").write_text("old", encoding="utf-8")

        _run("-c", str(cfg), "--no-progress")
        remaining = sorted(p.name for p in out.iterdir())
        assert not any("2025" in n for n in remaining), remaining


class TestLegacyMenu:
    def test_menu_starts_and_shows_the_locations(self, legacy):
        cfg, _ = legacy
        r = subprocess.run(
            [sys.executable, str(MENU), "--config", str(cfg)],
            input="", capture_output=True, text=True, timeout=120,
        )
        assert "拠点A / 拠点B" in r.stdout
        # 比較が 1 組なので「比較対象を変更」は出さない
        assert "比較対象を変更" not in r.stdout
        assert "Traceback" not in r.stderr


class TestMixingIsRejected:
    def test_locations_and_comparisons_cannot_coexist(self, legacy):
        """両方書かれていたら、どちらを使うか曖昧なのでエラーにする。"""
        cfg, _ = legacy
        data = yaml.safe_load(cfg.read_text(encoding="utf-8"))
        data["comparisons"] = {
            "x": [{"name": "A", "path": str(cfg.parent / "拠点A")},
                  {"name": "B", "path": str(cfg.parent / "拠点B")}]
        }
        cfg.write_text(yaml.safe_dump(data, allow_unicode=True), encoding="utf-8")

        r = _run("-c", str(cfg), "--no-progress")
        assert r.returncode == 2
        assert "同時に指定できません" in r.stderr
