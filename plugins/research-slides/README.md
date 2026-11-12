# research-slides

研究の記録から**発表の構成案**を作り、研究室の**進捗報告 PowerPoint**（4:3）に仕上げるための Claude Code プラグインです。スキルが2つ入っています。

| スキル | 何をするか | 頼み方の例 |
|---|---|---|
| `create-doc` | 実験記録（md）・結果 CSV・図・コードを読み、スライド1枚ごとに「型・主張・本文・図・ノート」を決めた構成案（Markdown）を、使う画像と一緒に1つのフォルダに出力する。数値は出典ファイルと機械照合し、質疑用の補足スライドと想定質問も付ける | 「今回の結果で、ゼミ発表の構成案を作って」 |
| `progress-report-deck` | 過去の進捗報告 pptx（無ければ同梱のテンプレート）をコピーして中身を差し替え、新しい pptx を作る。仕上げに機械チェック（壊れ・16pt 未満・はみ出し・1〜4字だけ次行に落ちる行・下の空白・配色・消し忘れ・言い回し）と、レンダリングした画像での目視確認をする。希望すれば教授役と初見の聞き手役の2つのエージェントに添削させる | 「この構成案をスライドにして」「今月の進捗報告を作って」 |

2つを続けて使うのが基本の流れです（`create-doc` で構成案 → `progress-report-deck` でスライド化）。`progress-report-deck` は、手書きのメモや下書きの md からでも作れます。

## インストール

Claude Code の中で:

```
/plugin marketplace add yapi777/yapi-skills
/plugin install research-slides@yapi-skills
```

プラグインを使わずに入れる方法（`setup.sh` / `setup.ps1`）は、[リポジトリの README](../../README.md#インストール) を見てください。

## 組み込まれているデザインルール

- 文字は **16pt 以上**（図の注や引用文献も）
- 配色はテーマの **2色だけ**（既定は青 `#156082`＝提案・結論、オレンジ `#E97132`＝問題・従来手法）。太字には色を付けない
- スライドの**下半分を空白のまま残さない**（チップ＋パネル、KPIタイル、表＋帯などの「型」で面を作る）
- 数式は PowerPoint の正式な数式（**OMML**）で組む
- 4:3・メイリオ

自分の過去資料をベースにすると、その資料のテーマの色に自動で合わせます。ルールを変えたいときは、プロフィールの「独自ルール」に書きます（下記）。

## 機械チェック（check_deck.py）

できあがった pptx を検査するスクリプトです。スキルが自動で流しますが、単体でも使えます。

```bash
python skills/progress-report-deck/scripts/check_deck.py 進捗報告202611.pptx --base 前回の資料.pptx --skip 3,4
python skills/progress-report-deck/scripts/check_deck.py --text スライド作成ドキュメント-2026-11-12.md   # 構成案の文言だけ
```

| 判定 | 内容 |
|---|---|
| FAIL | PowerPoint で開けない・修復ダイアログが出る壊れ（切れた参照、sldIdLst から外したのに残ったスライド、不正なテーマ色名など）、16pt 未満の文字 |
| WARN | はみ出し・最終行が1〜4字だけの行（PowerPoint の折り返しを見積もる）、スライド外・隠れた文字・空の箱、下の空白、パレット外の色、太字の色、仮の文字（〇〇・TODO）や同梱テンプレートのダミー文の残り、AI 臭のする語・表記ゆれ・読点の混在 |

折り返しの見積もりは、同梱テンプレートを PowerPoint で PDF にした実際の行と照合して調整しています（メイリオ）。最後の行に必ず `DONE: <FAIL数> FAIL / <WARN数> WARN / <枚数> slides` が出ます。この行が無ければ検査は途中で落ちています。

テスト: `python -m unittest discover -s plugins/research-slides/tests`（リポジトリのルートで）

## 謝辞

機械チェックの一部の考え方（XML の健全性検査、文字量からのはみ出し・泣き別れの見積もり、隠れ・空箱の判定）、AI 臭のする語のリスト、文脈を伏せた読者役のレビューと採否表は、[carnot-tech/jinba-consulting-pptx-skill](https://github.com/carnot-tech/jinba-consulting-pptx-skill)（MIT License）を参考に、4:3・16pt 以上の研究発表向けに作り直したものです。コードは流用していません。

## 必要な環境

| もの | 用途 | 必須？ |
|---|---|---|
| Python 3.10+ と `requirements.txt` のパッケージ | pptx の編集・テキスト抽出・画像化 | 必須 |
| Node.js | 数式の埋め込み・スライド複製・文字サイズのチェック | 必須 |
| `zip` / `unzip` | pptx の展開・再圧縮（Git for Windows の Git Bash には入っています） | 必須 |
| PowerPoint（Windows） | 本番と同じ見た目でのレンダリング確認 | 任意（あると確実） |
| LibreOffice | PowerPoint が無い環境でのレンダリング確認（数式は表示されません） | 任意 |

Python パッケージは次のコマンドで入ります（手動インストールの `setup.sh` / `setup.ps1` なら自動で入ります）:

```bash
pip install -r requirements.txt
# または
pip install python-pptx Pillow defusedxml lxml "markitdown[pptx]" pymupdf
```

## はじめて使うとき

1. スライドを作りたいフォルダで Claude Code を開き、「進捗報告のスライドを作って」と頼みます（プラグインで入れた場合は `/research-slides:progress-report-deck` でも呼べます）
2. 最初に氏名・所属・研究室・学年・研究分野を聞かれます。保存を許可すると `~/.claude/yapi-skills/profile.md` に書かれ、次からは聞かれません
3. フォルダに過去の進捗報告（ファイル名に「進捗報告」を含む pptx）があれば、それをベースにします。無ければ同梱のテンプレート（匿名化したサンプル18枚）から作ります。研究背景のスライドはダミーなので、初回は中身を聞かれます

## プロフィールでできるカスタマイズ

`~/.claude/yapi-skills/profile.md` の `## research-slides` の節に書きます。

```markdown
## research-slides

- 資料のファイル名: <氏名>-進捗報告YYYYMM.pptx
- 過去の資料の置き場所: D:/lab/進捗報告
- 主な聞き手: 研究室メンバー（基礎は知っている前提）
- 独自ルール: 配色は緑 #2E7D32 と紫 #8E24AA。スピーカーノートは入れない
```

ひな形: [skills/progress-report-deck/references/profile-template.md](skills/progress-report-deck/references/profile-template.md)

## 同梱テンプレートについて

`skills/progress-report-deck/assets/template.pptx` は、論文の査読対応の回を題材にした 18 枚のサンプルです。文字はすべてダミーで、写真は灰色の枠に置き換えてあります。どの型にどんな文を入れるかの見本も兼ねています。研究室で決まったテンプレートがあれば、それを使った過去資料を作業フォルダに置いてください。そちらが優先されます。
