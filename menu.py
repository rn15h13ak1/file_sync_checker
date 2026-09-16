"""
ファイル同期チェック 対話メニュー
=================================
CLI オプションを覚えなくても、番号を選ぶだけで実行できる。

  python3 menu.py                  # メニューを表示
  python3 menu.py --config my.yaml # 設定ファイルを指定

実行前に比較対象の拠点を表示するので、設定の取り違えに気付ける。

無人実行 (cron / タスクスケジューラ) はこのメニューではなく本体を直接呼ぶこと:
  python3 main.py --no-progress
"""

import argparse
import json
import subprocess
import sys
from pathlib import Path

try:
    import config as fsc_config
except ModuleNotFoundError as e:
    # menu.bat のダブルクリック起動で依存が入っていない場合に、
    # トレースバックではなく対処を表示する
    if e.name not in ("yaml", "tqdm", "openpyxl"):
        raise
    print(f"必要なライブラリ {e.name} が入っていません。")
    print("次のコマンドでインストールしてください:")
    print("    pip install -r requirements.txt")
    sys.exit(2)

WIDTH = 60
TOOL_DIR = Path(__file__).resolve().parent
SCRIPT = TOOL_DIR / "main.py"
HISTORY_PATH = Path.home() / ".file_sync_checker_menu.json"

MODE_RUN = "run"
MODE_CUSTOM = "custom"
MODE_SWITCH = "switch"
MODE_SHOW = "show"
MODE_OPEN = "open"

MODES = [
    (MODE_RUN, "チェックを実行", "設定ファイルのとおりに照合する"),
    (MODE_CUSTOM, "条件を選んで実行", "ハッシュ方式・出力形式・再試行回数を選ぶ"),
    # 比較が複数定義されているときだけ出す
    (MODE_SWITCH, "比較対象を変更", "設定に定義された比較を切り替える"),
    (MODE_SHOW, "設定を確認", "比較する拠点・除外パターン・出力先を表示する"),
    (MODE_OPEN, "前回のレポートを開く", "出力先にある最新のレポートを開く"),
]

# 本体の終了コード (README の Exit code と対応)
EXIT_MEANINGS = {
    0: "全拠点一致・エラーなし",
    1: "差分を検出 (スキャン自体は完走)",
    2: "設定エラー",
    3: "読み取りエラーあり (スキャンが不完全)",
    4: "予期しないエラー",
    130: "中断",
}

# メニュー自身の終了コード (tool_launcher と揃える)
EXIT_OK = 0
EXIT_EOF = 1
EXIT_USAGE = 2
EXIT_INTERRUPTED = 130


# ===========================================================================
# 表示・入力
# ===========================================================================


def hr(char="="):
    print(char * WIDTH)


def print_menu(title: str, items: list, back_label: str = "戻る",
               default: int | None = None, default_mark: str = "既定") -> int:
    """
    メニューを表示して選択番号を返す。0 = 戻る / 終了。
    default_mark は既定値の由来を示すラベル (前回の選択なら「前回」)。
    """
    while True:
        print()
        hr()
        print(f"  {title}")
        hr()
        for i, item in enumerate(items, 1):
            mark = f" <-{default_mark}" if default == i else ""
            print(f"  {i}. {item}{mark}")
        hr("-")
        print(f"  0. {back_label}")
        hr()
        prompt = ("番号を入力してください"
                  + (f" [Enter={default}]: " if default else ": "))
        choice = input(prompt).strip()
        if not choice and default:
            return default
        if choice == "0":
            return 0
        if choice.isdigit() and 1 <= int(choice) <= len(items):
            return int(choice)
        print("  ※ 無効な入力です。もう一度入力してください。")


def input_text(prompt: str, default: str = "", validate=None,
               hint: str = "") -> str | None:
    """
    文字列を入力させる。空 Enter は default を採用する。
    default が無い状態で空 Enter を押した場合は None (キャンセル) を返す。
    """
    suffix = f" [Enter={default}]" if default else " (空 Enter で戻る)"
    while True:
        answer = input(f"  {prompt}{suffix}: ").strip()
        if not answer:
            if default:
                return default
            return None
        if validate is None or validate(answer):
            return answer
        print(f"  ※ {hint}")


def is_retry_count(value: str) -> bool:
    return value.isdigit()


# ===========================================================================
# 前回値の記憶
# ===========================================================================


def load_history() -> dict:
    """前回の入力値を読む。壊れていても既定値で続行する。"""
    try:
        with open(HISTORY_PATH, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def save_history(data: dict) -> None:
    """入力値を保存する。書けなくても実行は妨げない。"""
    try:
        with open(HISTORY_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except OSError:
        pass


# ===========================================================================
# 設定の読み取り
# ===========================================================================


def load_config_or_none(config_path: str, comparison: str = ""):
    """設定を読む。読めない場合は (None, 理由) を返す。

    比較名を渡さない場合でも、選択を求めずに読む (拠点名の表示や
    比較の一覧取得に使うため)。
    """
    try:
        return fsc_config.load_config(
            config_path, comparison=comparison or None,
            require_selection=bool(comparison),
        ), ""
    except fsc_config.ConfigError as e:
        return None, str(e)
    except Exception as e:  # YAML パースエラー等
        return None, f"設定ファイルを読めません: {e}"


def list_comparisons(config_path: str) -> list:
    """設定に定義されている比較の名前。従来形式なら空リスト。"""
    config, _ = load_config_or_none(config_path)
    return config.comparison_names if config else []


def choose_comparison(config_path: str, current: str) -> str | None:
    """比較対象を選ぶ。戻る場合は None。"""
    names = list_comparisons(config_path)
    if not names:
        print("  ※ この設定には比較が 1 組しかありません (comparisons 未使用)。")
        return None
    default = names.index(current) + 1 if current in names else 1
    items = []
    for name in names:
        config, _ = load_config_or_none(config_path, name)
        locs = " / ".join(loc.name for loc in config.locations) if config else "?"
        items.append(f"{name} ({locs})")
    choice = print_menu("比較対象を選択", items, default=default)
    if choice == 0:
        return None
    return names[choice - 1]


def describe_targets(config_path: str, comparison: str = "") -> str:
    """比較対象を 1 行で表す。読めない場合は理由を返す。

    設定を取り違えたまま実行する事故を、実行前に気付けるようにする。
    """
    config, reason = load_config_or_none(config_path, comparison)
    if config is None:
        return f"※ {reason}"
    names = " / ".join(loc.name for loc in config.locations)
    label = f"[{comparison}] " if comparison else ""
    return f"{label}{names} ({len(config.locations)} 拠点)"


def print_config(config_path: str, comparison: str = "") -> None:
    """設定の内容を一覧表示する。"""
    config, reason = load_config_or_none(config_path, comparison)
    print()
    hr()
    print("  設定内容")
    hr()
    print(f"  設定ファイル: {config_path}")
    if config is None:
        print(f"  ※ {reason}")
        return

    names = config.comparison_names
    if names:
        print(f"  定義されている比較: {', '.join(names)}")

    print()
    print(f"  比較する拠点{f' [{comparison}]' if comparison else ''}:")
    for loc in config.locations:
        print(f"    - {loc.name}: {loc.path}")

    print()
    print("  除外パターン:")
    if config.exclude_patterns:
        for pat in config.exclude_patterns:
            target = "相対パス全体" if "/" in pat else "ファイル名"
            print(f"    - {pat}  ({target}と照合)")
    else:
        print("    (なし)")

    print()
    print("  出力:")
    print(f"    形式        : {config.output.format}")
    print(f"    出力先      : {config.output.output_dir}")
    print(f"    表の行数上限: {config.output.max_table_rows or '無制限'}")
    keep = config.output.keep_reports
    print(f"    レポート保持: {f'直近 {keep} 回分' if keep else '削除しない'}")

    print()
    print("  照合:")
    print(f"    ハッシュ方式: {config.performance.hash_algorithm} / "
          f"{config.performance.hash_mode}")
    if config.performance.hash_mode == fsc_config.HASH_MODE_ALWAYS:
        print("      (全ファイルの内容を読んで照合する)")
    else:
        print("      (サイズ・更新日時が全拠点一致なら読み取りを省略する)")
    print(f"    並列数      : {config.performance.parallel_workers}")
    print(f"    Unicode正規化: {'あり (NFC)' if config.matching.normalize_unicode else 'なし'}")
    print(f"    大文字小文字: {'区別する' if config.matching.case_sensitive else '区別しない'}")


# ===========================================================================
# 実行内容の組み立て
# ===========================================================================


def build_args(config_path: str = "", hash_mode: str = "",
               output_format: str = "", retry: str = "",
               comparison: str = "") -> list:
    """本体に渡すコマンドライン引数を組み立てる。

    空文字の項目は指定しない (設定ファイルの値がそのまま使われる)。
    """
    args = []
    if config_path:
        args += ["--config", config_path]
    if comparison:
        args += ["--comparison", comparison]
    if hash_mode:
        args += ["--hash-mode", hash_mode]
    if output_format:
        args += ["--format", output_format]
    if retry and retry != "0":
        args += ["--retry", retry]
    return args


def run_checker(args: list) -> int:
    """本体を実行して終了コードを返す。"""
    if not SCRIPT.is_file():
        print(f"\n  ※ 本体が見つかりません: {SCRIPT}")
        return 127
    print()
    hr("-")
    try:
        completed = subprocess.run([sys.executable, str(SCRIPT), *args], cwd=TOOL_DIR)
    except KeyboardInterrupt:
        print("\n  中断しました。")
        return EXIT_INTERRUPTED
    return completed.returncode


# ===========================================================================
# レポートを開く
# ===========================================================================


def latest_report(config_path: str, comparison: str = ""):
    """出力先にある最新のレポートを返す。無ければ None。

    比較名を指定すると、その比較のレポートだけを対象にする。
    """
    config, _ = load_config_or_none(config_path, comparison)
    if config is None:
        return None
    out_dir = config.output.output_dir
    if not out_dir.is_dir():
        return None
    prefix = f"sync-check-{comparison}" if comparison else "sync-check"
    reports = [p for p in out_dir.iterdir()
               if p.is_file() and p.suffix in (".html", ".xlsx")
               and p.name.startswith(prefix)]
    if not reports:
        return None
    return max(reports, key=lambda p: p.stat().st_mtime)


def open_file(path: Path) -> int:
    """OS の既定のアプリでファイルを開く。"""
    try:
        if sys.platform == "win32":
            import os
            os.startfile(str(path))  # noqa: S606  Windows のみ
        elif sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=True)
        else:
            subprocess.run(["xdg-open", str(path)], check=True)
    except (OSError, subprocess.CalledProcessError) as e:
        print(f"  ※ 開けませんでした: {e}")
        return 1
    return 0


# ===========================================================================
# 各モードの対話
# ===========================================================================


def choose_hash_mode(history: dict) -> str | None:
    """ハッシュ方式を選ぶ。戻る場合は None、設定どおりなら空文字。"""
    items = [
        "設定ファイルのとおり",
        "always ― 全ファイルの内容を読む (更新日時を信用しない)",
        "smart ― サイズ・更新日時が全拠点一致なら読み飛ばす (速い)",
    ]
    choice = print_menu("ハッシュ方式", items, default=1)
    if choice == 0:
        return None
    return ["", fsc_config.HASH_MODE_ALWAYS, "smart"][choice - 1]


def choose_format() -> str | None:
    """出力形式を選ぶ。戻る場合は None、設定どおりなら空文字。"""
    items = ["設定ファイルのとおり", "excel", "html", "both (Excel と HTML の両方)"]
    choice = print_menu("出力形式", items, default=1)
    if choice == 0:
        return None
    return ["", "excel", "html", "both"][choice - 1]


def run_custom(config_path: str, history: dict, comparison: str = "") -> int | None:
    """条件を選んで実行する。戻る場合は None。"""
    hash_mode = choose_hash_mode(history)
    if hash_mode is None:
        return None

    output_format = choose_format()
    if output_format is None:
        return None

    retry = input_text(
        "読み取り失敗時の再試行回数",
        default=history.get("retry", "0"),
        validate=is_retry_count,
        hint="0 以上の整数を入力してください (0 = 再試行しない)。",
    )
    if retry is None:
        return None

    history.update({"hash_mode": hash_mode, "format": output_format, "retry": retry})
    save_history(history)

    return run_checker(build_args(
        config_path=config_path,
        comparison=comparison,
        hash_mode=hash_mode,
        output_format=output_format,
        retry=retry,
    ))


def run_mode(mode: str, label: str, config_path: str, history: dict,
             comparison: str = "") -> int | None:
    """1 つのモードを実行する。戻る場合は None。"""
    print()
    hr()
    print(f"  {label}")
    hr()

    if mode == MODE_RUN:
        return run_checker(build_args(config_path=config_path, comparison=comparison))

    if mode == MODE_CUSTOM:
        return run_custom(config_path, history, comparison)

    if mode == MODE_SHOW:
        print_config(config_path, comparison)
        return None

    if mode == MODE_OPEN:
        report = latest_report(config_path, comparison)
        if report is None:
            print("  ※ レポートが見つかりません。先にチェックを実行してください。")
            return None
        print(f"  開きます: {report}")
        open_file(report)
        return None

    return None


# ===========================================================================
# エントリポイント
# ===========================================================================


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="ファイル同期チェック 対話メニュー",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="無人実行には main.py を直接使ってください。",
    )
    parser.add_argument(
        "--config",
        default=str(TOOL_DIR / "config.yaml"),
        help="設定ファイルのパス (デフォルト: config.yaml)",
    )
    return parser


def available_modes(config_path: str) -> list:
    """この設定で選べる操作。比較が 1 組だけなら「比較対象を変更」は出さない。"""
    if len(list_comparisons(config_path)) > 1:
        return MODES
    return [m for m in MODES if m[0] != MODE_SWITCH]


def initial_comparison(config_path: str, history: dict) -> str:
    """起動時に選ばれている比較を決める。

    前回の選択が今も設定にあればそれを使う。無ければ先頭の比較。
    比較が定義されていない従来形式では空文字。
    """
    names = list_comparisons(config_path)
    if not names:
        return ""
    last = history.get("comparison", "")
    return last if last in names else names[0]


def main() -> None:
    args = build_parser().parse_args()
    history = load_history()
    comparison = initial_comparison(args.config, history)
    modes = available_modes(args.config)
    last_mode = history.get("mode")
    default_choice = next(
        (i for i, (key, _, _) in enumerate(modes, 1) if key == last_mode), None
    )

    items = [f"{label} ― {desc}" for _, label, desc in modes]

    try:
        while True:
            print()
            hr()
            print("  ファイル同期チェック")
            hr()
            print(f"  比較対象: {describe_targets(args.config, comparison)}")
            choice = print_menu("実行内容を選択", items, back_label="終了",
                                default=default_choice, default_mark="前回")
            if choice == 0:
                print("  終了します。")
                sys.exit(EXIT_OK)

            mode, label, _ = modes[choice - 1]

            if mode == MODE_SWITCH:
                print()
                hr()
                print(f"  {label}")
                hr()
                picked = choose_comparison(args.config, comparison)
                if picked:
                    comparison = picked
                    history["comparison"] = comparison
                    save_history(history)
                    print(f"  比較対象を「{comparison}」にしました。")
                default_choice = choice
                input("\n  Enter キーでメニューに戻ります...")
                continue

            rc = run_mode(mode, label, args.config, history, comparison)
            history["mode"] = mode
            save_history(history)
            default_choice = choice
            if rc is not None:
                print()
                hr("-")
                meaning = EXIT_MEANINGS.get(rc, "不明なコード")
                print(f"  終了コード: {rc} ({meaning})")
            input("\n  Enter キーでメニューに戻ります...")
    except EOFError:
        print("\n  入力が終了しました。")
        sys.exit(EXIT_EOF)
    except KeyboardInterrupt:
        print("\n  中断しました。")
        sys.exit(EXIT_INTERRUPTED)


if __name__ == "__main__":
    main()
