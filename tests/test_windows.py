"""Windows 環境への対応を守るテスト。

macOS / Linux では気づけない退行を検出する。本ツールの主な対象は
Windows の UNC 共有 (`\\\\server\\share\\...`) なので、Windows で
動かないと使い物にならない。
"""
from __future__ import annotations

import re
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent

# コンソールに出力するソース。
# Windows の日本語コンソールは既定で CP932 のため、CP932 に無い文字を print すると
# 文字化けではなく UnicodeEncodeError で落ちる。menu.bat をダブルクリックした場合、
# メニューの最初の画面すら出ない。
#
# reporter.py は対象外。出力先が HTML (UTF-8) と Excel (XML) で、
# コンソールには何も書かないため、絵文字などを使ってよい。
CONSOLE_SOURCES = ("main.py", "menu.py", "scanner.py", "comparator.py",
                   "config.py", "utils.py")


def _unencodable(text: str) -> set:
    bad = set()
    for ch in set(text):
        if ord(ch) < 128:
            continue
        try:
            ch.encode("cp932")
        except UnicodeEncodeError:
            bad.add(ch)
    return bad


class TestCp932Safety:
    """画面に出る文字がすべて CP932 で表現できること。"""

    @pytest.mark.parametrize("name", CONSOLE_SOURCES)
    def test_sources_are_cp932_safe(self, name):
        bad = _unencodable((ROOT / name).read_text(encoding="utf-8"))
        assert bad == set(), (
            f"{name} に CP932 で表現できない文字があります: "
            + ", ".join(f"U+{ord(c):04X} {c!r}" for c in sorted(bad))
            + "。Windows の日本語コンソールで UnicodeEncodeError になります。"
        )

    def test_detects_a_known_bad_character(self):
        """検査自体が機能していることを確かめる (em dash は CP932 に無い)。"""
        assert _unencodable("作成完了 — 件名") == {"—"}

    @pytest.mark.parametrize("ch", "←※【】―…○")
    def test_common_symbols_are_safe(self, ch):
        """実際に使っている記号が CP932 にあること。"""
        assert _unencodable(ch) == set()


class TestMenuBat:
    """menu.bat が cmd.exe で正しく解釈される形であること。"""

    @pytest.fixture
    def raw(self) -> bytes:
        return (ROOT / "menu.bat").read_bytes()

    def test_uses_crlf(self, raw):
        """cmd.exe は LF だけの .bat を誤動作させることがある。"""
        assert b"\r\n" in raw
        assert re.search(rb"[^\r]\n", raw) is None, "LF だけの行があります"

    def test_is_ascii_only(self, raw):
        """コンソールのコードページに依存しないよう ASCII に収める。"""
        raw.decode("ascii")   # 例外が出なければ OK

    def test_calls_menu_py(self, raw):
        assert b"menu.py" in raw

    def test_falls_back_from_py_launcher_to_python(self, raw):
        text = raw.decode("ascii")
        assert "py -3" in text and "python --version" in text

    def test_pauses_so_the_window_stays_open(self, raw):
        """ダブルクリック起動でエラーを読めるようにする。"""
        assert b"pause" in raw

    def test_uses_pushd_for_unc_paths(self, raw):
        """共有ドライブ上の UNC パスから起動しても動くようにする。"""
        text = raw.decode("ascii")
        assert "pushd" in text and "popd" in text

    def test_gitattributes_pins_crlf(self):
        """チェックアウト時に LF へ正規化されないよう固定する。"""
        attrs = (ROOT / ".gitattributes").read_text(encoding="utf-8")
        assert "*.bat" in attrs and "eol=crlf" in attrs


class TestWindowsPaths:
    """Windows のパス表記が壊れないこと。"""

    def test_unc_root_produces_backslash_paths(self):
        from reporter import _build_paths

        full, folder = _build_paths("//server/share/docs", "設計/仕様.xlsx")
        assert full == "\\\\server\\share\\docs\\設計\\仕様.xlsx"
        assert folder == "\\\\server\\share\\docs\\設計"

    def test_backslash_in_config_is_rejected_with_guidance(self, tmp_path):
        """YAML でバックスラッシュが潰れる事故を、理由付きで弾く。"""
        from config import ConfigError, load_config

        cfg = tmp_path / "c.yaml"
        cfg.write_text(
            'locations:\n'
            '  - {name: A, path: "\\\\\\\\server\\\\share"}\n'
            '  - {name: B, path: /opt/b}\n',
            encoding="utf-8",
        )
        with pytest.raises(ConfigError, match="バックスラッシュ"):
            load_config(cfg)
