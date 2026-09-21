# file_sync_checker

共通規約: [../ws-conventions/README.md](../ws-conventions/README.md) に従う（`~/ws` 配下の全リポジトリ共通）。

各リポジトリ固有の事情は本ファイルに追記する。

## 共通規約からの逸脱

### commit / push とテストは確認を取らずに行う

共通規約 A は「commit / push は、利用者が明示的に指示したときだけ実行する」としているが、
**本リポジトリでは修正のたびに自動でコミットし、`origin main` へ push する。**
テストと本リポジトリ内のファイル編集も、確認を取らずに実行する（2026-09-21 指示）。

- ブランチは切らず `main` に直接コミットする
- 変更 → テスト → 検査 → CHANGELOG 追記 → コミット → push の順で行う
- 他のリポジトリには適用しない

## 作業のたびに実行する

コミット前に、共通規約の検査とテストを実行する。

```bash
../ws-conventions/bin/check-markdown.sh .
../ws-conventions/bin/check-privacy.sh .
../ws-conventions/bin/check-commands.sh .
.venv/bin/python -m pytest -q
```

**`check-privacy.sh` は `LICENSE:3` の 1 件だけが出る。** MIT ライセンスの著作権者名で、
消すとライセンスが成立しない。これ以外が出たら内容を確認して直すこと。

**検査を足したら、この手順にも書き足す。** `check-commands.sh` が、ここに書いた手順と
`../ws-conventions/bin` の実体を突き合わせている。書き忘れた検査は実行されない。

**`check-terms.sh` と `gen-decision-index.py` は、本リポジトリでは使わない。**
どちらも ADR を使うリポジトリ向けで、本リポジトリは ADR を置いていない。

## 配布済みの設定ファイルを壊さない

**このツールは既に利用者へ配布されている。** 手元の `config.yaml` はトップレベルに
`locations:` を書く従来形式で、書き換えを求めずに新しい版へ入れ替えられる必要がある。

設定の形式・レポートのファイル名・固定名リンク（`sync-check.html`）を変えるときは、
[tests/test_backward_compat.py](tests/test_backward_compat.py) が守っている内容を壊していないか確かめる。
新しい設定項目は、**書かなければ従来どおりの挙動**になる既定値を持たせる。

## Windows で動かす前提

利用者の実行環境は Windows で、`menu.bat` のダブルクリックから起動する。

- `menu.bat` は CRLF・ASCII のみ。改行は `.gitattributes` で固定している
- **コンソールに出るソースは CP932 で表現できる文字だけにする。**
  em dash（U+2014）は CP932 に無いため U+2015 を使う
- `reporter.py` は対象外（出力先が HTML / Excel でコンソールに書かない）

これらは [tests/test_windows.py](tests/test_windows.py) が検査している。
