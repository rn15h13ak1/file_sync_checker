# File Sync Checker

3 箇所（以上）の共有フォルダの内容が一致しているかを SHA-256 ハッシュで突き合わせ、Excel / HTML レポートにまとめるツール。

## 必要環境

- Python 3.9 以上
- 各拠点に `\\server\share\...` 形式（または `/Volumes/...` 等）でアクセスできること

## セットアップ

```bash
pip install -r requirements.txt
cp config.example.yaml config.yaml
# config.yaml を編集（拠点パス・除外パターン・出力形式）
```

## 実行

```bash
python main.py                    # 既定で config.yaml を使用
python main.py -c path/to/cfg.yml # 設定ファイル指定
python main.py --no-progress      # 進捗バーを抑制
python main.py -v                 # 詳細ログ
```

CWD に依存せずフルパスで起動することもできます:

```bash
python /full/path/to/main.py -c /full/path/to/config.yaml
```

### パス解決ルール

- `-c` 省略時の既定 `config.yaml` は **CWD → main.py と同じディレクトリ** の順で探索。
- `config.yaml` 内の相対パス (`locations[].path`, `output.output_dir`) は **config ファイルのあるディレクトリ** からの相対として解決されます。絶対パスはそのまま使用されます。

レポートは `<output_dir>/sync-check-YYYYMMDD-HHMMSS.{xlsx,html}` に出力されます。

### 終了コード

| コード | 意味 |
|---|---|
| 0 | 全拠点一致・エラーなし |
| 1 | 差分またはエラーを検出 |
| 2 | 設定エラー |
| 130 | ユーザ中断 (Ctrl+C) |

## 設定ファイル

`config.example.yaml` を参照。主な項目：

- `locations`: 比較対象の拠点（2件以上）
- `exclude_patterns`: ファイル名・ディレクトリ名の glob 除外（`~$*` 等）
- `output.format`: `excel` / `html` / `both`
- `output.output_dir`: レポート出力先
- `performance.parallel_workers`: ハッシュ並列計算スレッド数
- `performance.hash_algorithm`: 現状 `sha256` のみ

## 出力レポート構成

| シート / セクション | 内容 |
|---|---|
| サマリー | 実行日時・拠点別ファイル数/サイズ・差分件数 |
| 全ファイル一覧 | 和集合。1行 = 1相対パス、各拠点のサイズ/ハッシュ/更新日時を横並び |
| ハッシュ不一致 | 同名ファイルだが内容が異なる |
| ファイル欠落 | 一部の拠点に存在しない（多数派には存在） |
| 余分なファイル | 一部の拠点のみに存在 |
| フォルダ構造差分 | 拠点ごとにフォルダの有無が異なる |
| エラー | 読み取り失敗・アクセス不可 |

### 表示の工夫

- 状態列を色分け（OK=緑 / 不一致=赤 / 欠落=オレンジ / 一部のみ=黄）
- ハッシュ値はセル幅で短縮表示、セル選択（Excel）/ ホバー（HTML）でフル値確認可能
- 欠落セルは「－」灰色背景
- 1行目固定 + 相対パス列固定（横スクロール対応）

## モジュール構成

```
file_sync_checker/
├── main.py        エントリポイント
├── config.py      YAML 読み込み・検証
├── scanner.py     再帰スキャン + 並列 SHA-256
├── comparator.py  N 拠点間の差分検出
├── reporter.py    Excel / HTML 出力
├── utils.py       ロギング・共通関数
└── config.example.yaml
```

## 設計メモ

- 本ツールは**検出のみ**。修正・反映は人間が判断して実施。
- ロック中ファイル等の読み取り失敗はエラー扱いで継続実行し、レポートに記録（リトライなし）。
- `locations` は2件以上に対応。「欠落あり」「一部のみ存在」は過半数を基準にラベル分け。
