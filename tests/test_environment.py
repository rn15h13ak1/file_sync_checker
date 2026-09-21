"""README が「動作要件」に書いた環境で起動できることを守るテスト。

開発機の Python は 3.12 以降で、依存も入っている。**そのため、
動作要件どおりの環境で起動できなくなっても、他のテストは全部通る。**
配布先は Windows で、同梱の Python を使うとは限らない。

ここでは 2 つを固定する。

- Python 3.9 で評価できない注釈を書かないこと（`str | None` は 3.10 以降）
- 依存が入っていない Python で起動したとき、トレースバックではなく
  対処を表示し、**差分 (1) と紛れない終了コードを返すこと**
"""
from __future__ import annotations

import ast
import os
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent

# 依存が入っていない環境を作るための細工。site が起動時に読み込む。
# ライブラリをアンインストールする代わりに、import を名前で撥ねる。
BLOCKER = """
import sys


class _Blocker:
    def find_spec(self, name, path=None, target=None):
        if name.split(".")[0] == {blocked!r}:
            raise ModuleNotFoundError(f"No module named {{name!r}}", name=name)
        return None


sys.meta_path.insert(0, _Blocker())
"""


class TestPython39Annotations:
    """`str | None` は Python 3.10 以降の書き方 (PEP 604)。

    `from __future__ import annotations` があれば注釈は文字列のまま
    残るので 3.9 でも書けるが、無いと **def の時点で評価され、
    モジュールの import そのものが TypeError で落ちる。**
    関数を呼ぶ前に落ちるため、起動した瞬間にトレースバックになる。
    """

    @staticmethod
    def _pep604_annotations(tree: ast.AST) -> list:
        """注釈のうち `X | Y` を使っているものを返す。"""
        found = []
        for node in ast.walk(tree):
            annotations = []
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                args = node.args
                for a in args.args + args.posonlyargs + args.kwonlyargs:
                    if a.annotation:
                        annotations.append(a.annotation)
                if node.returns:
                    annotations.append(node.returns)
            elif isinstance(node, ast.AnnAssign):
                annotations.append(node.annotation)
            for ann in annotations:
                if any(
                    isinstance(n, ast.BinOp) and isinstance(n.op, ast.BitOr)
                    for n in ast.walk(ann)
                ):
                    found.append(f"{ann.lineno}: {ast.unparse(ann)}")
        return found

    @staticmethod
    def _has_future_annotations(tree: ast.AST) -> bool:
        for node in tree.body:
            if isinstance(node, ast.ImportFrom) and node.module == "__future__":
                if any(alias.name == "annotations" for alias in node.names):
                    return True
        return False

    @pytest.mark.parametrize(
        "name", sorted(p.name for p in ROOT.glob("*.py"))
    )
    def test_module_is_importable_on_python_39(self, name: str):
        tree = ast.parse((ROOT / name).read_text(encoding="utf-8"))
        if self._has_future_annotations(tree):
            return
        hits = self._pep604_annotations(tree)
        assert not hits, (
            f"{name} は `from __future__ import annotations` が無いまま "
            f"`X | Y` の注釈を使っており、Python 3.9 では import できない: {hits}"
        )

    def test_the_check_would_catch_a_regression(self):
        """検査自体が効いていることを確かめる (常に通る検査を防ぐ)。"""
        tree = ast.parse("def f(x: int | None) -> str | None: ...")
        assert not self._has_future_annotations(tree)
        assert len(self._pep604_annotations(tree)) == 2


@pytest.fixture
def without_tqdm(tmp_path: Path):
    """tqdm を import できない状態で子プロセスを起動する env を返す。"""
    (tmp_path / "sitecustomize.py").write_text(
        BLOCKER.format(blocked="tqdm"), encoding="utf-8"
    )
    env = dict(os.environ)
    existing = env.get("PYTHONPATH")
    env["PYTHONPATH"] = (
        f"{tmp_path}{os.pathsep}{existing}" if existing else str(tmp_path)
    )
    return env


class TestMissingDependency:
    """依存が欠けた状態で起動したときの振る舞い。

    終了コード 1 は「差分を検出」。環境が壊れているのに 1 を返すと、
    定期実行の通知では正常な検出と区別が付かず、**レポートが一枚も
    出ていないことに気付けない。**
    """

    def _run(self, script: str, env: dict) -> subprocess.CompletedProcess:
        return subprocess.run(
            [sys.executable, str(ROOT / script)],
            input="", capture_output=True, text=True, timeout=60, env=env,
        )

    def test_blocker_actually_blocks(self, without_tqdm):
        """細工が効いていることを先に確かめる。"""
        r = subprocess.run(
            [sys.executable, "-c", "import tqdm"],
            capture_output=True, text=True, timeout=60, env=without_tqdm,
        )
        assert r.returncode != 0
        assert "tqdm" in r.stderr

    @pytest.mark.parametrize("script", ["main.py", "menu.py"])
    def test_exit_code_is_not_the_diff_code(self, script: str, without_tqdm):
        r = self._run(script, without_tqdm)
        assert r.returncode == 2, (r.returncode, r.stdout, r.stderr)

    @pytest.mark.parametrize("script", ["main.py", "menu.py"])
    def test_shows_what_to_do_instead_of_a_traceback(self, script: str, without_tqdm):
        r = self._run(script, without_tqdm)
        out = r.stdout + r.stderr
        assert "Traceback" not in out, out
        assert "tqdm" in out
        assert "pip install -r requirements.txt" in out

    def test_main_names_the_interpreter(self, without_tqdm):
        """どの Python に入れればよいかが分かること。

        Windows には py ランチャ経由の版が複数入っていることがあり、
        「pip install した」のに直らない、が起こりやすい。
        """
        r = self._run("main.py", without_tqdm)
        assert sys.executable in r.stdout + r.stderr

    def test_unrelated_import_errors_still_raise(self, tmp_path: Path):
        """本ツールと無関係な ModuleNotFoundError は握り潰さない。"""
        (tmp_path / "sitecustomize.py").write_text(
            BLOCKER.format(blocked="pathlib"), encoding="utf-8"
        )
        env = dict(os.environ)
        env["PYTHONPATH"] = str(tmp_path)
        r = self._run("main.py", env)
        assert "pip install" not in r.stdout + r.stderr
