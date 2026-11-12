# yapi-skills

日本語で使う [Claude Code](https://code.claude.com/) のスキル集です。テーマごとに **プラグイン** に分けて配布しているので、必要なものだけ入れられます。

> **スキル**は、Claude Code に「この種類の作業はこの手順・ルールでやる」と教えるための説明書とスクリプトのセットです。入れておくと、頼み方に合わせて自動で使われるほか、`/スキル名` で呼び出せます。
> **プラグイン**は、スキルをまとめて配布・更新するための Claude Code の仕組みです。

## プラグイン一覧

| プラグイン | 入っているスキル | できること | 必要なもの |
|---|---|---|---|
| [research-slides](plugins/research-slides/) | `create-doc`、`progress-report-deck` | 研究の記録から発表の構成案を作り、研究室の進捗報告 PowerPoint（4:3）に仕上げる | Python 3.10+、Node.js（PowerPoint は任意） |

## インストール

### A. プラグインとして入れる（おすすめ）

Claude Code の中で:

```
/plugin marketplace add yapi777/yapi-skills
/plugin install research-slides@yapi-skills
```

ターミナルからなら `claude plugin marketplace add yapi777/yapi-skills` → `claude plugin install research-slides@yapi-skills`。

入れたスキルは `/research-slides:progress-report-deck` のように **`プラグイン名:スキル名`** で呼び出せます（「進捗報告のスライドを作って」のような頼み方でも自動で使われます）。プラグインごとに必要な Python パッケージなどは、各プラグインの README を見てください。

**更新**: `/plugin marketplace update yapi-skills`。`/plugin` の画面の **Marketplaces** で自動更新をオンにもできます。

### B. 手動で入れる（git clone）

```bash
git clone https://github.com/yapi777/yapi-skills.git
cd yapi-skills
bash setup.sh                      # すべてのプラグイン
bash setup.sh research-slides      # 指定したプラグインだけ
```

Windows PowerShell では `powershell -ExecutionPolicy Bypass -File setup.ps1 [プラグイン名]`。

`~/.claude/skills/` にスキルをコピーし、プラグインの `requirements.txt` があれば pip で入れます。同じ名前のスキルが既にあれば `<名前>.bak-<日時>` に退避します。この方法で入れたスキルは `/progress-report-deck` のように **スキル名だけ** で呼び出せます。更新するときは `git pull` してからもう一度 `setup` を実行します。

> A と B の両方で同じスキルを入れると、同じ依頼に2つのスキルが反応します。どちらか一方にしてください。

## プロフィール

氏名・所属・研究分野などの利用者ごとの情報は、スキルの外の **`~/.claude/yapi-skills/profile.md`**（Windows: `%USERPROFILE%\.claude\yapi-skills\profile.md`）に置きます。スキルを更新しても消えません。

初めて使うときにスキルが質問し、保存してよいか確認してから作ります。自分で書く場合のひな形は [profile-template.md](plugins/research-slides/skills/progress-report-deck/references/profile-template.md) です。項目は `## 共通` と、プラグインごとの節（`## research-slides` など）に分かれています。

## 新しいスキル・プラグインを追加する（作者向けメモ）

```
plugins/<プラグイン名>/
├── .claude-plugin/plugin.json   # name / description / author / license（version は書かない＝コミットごとに更新が届く）
├── README.md                    # 何ができるか・必要な環境・使い方
├── requirements.txt             # 必要な Python パッケージ（あれば。setup スクリプトが入れる）
└── skills/<スキル名>/SKILL.md   # スクリプトや素材は同じフォルダの scripts/ assets/ references/ に置く
```

1. 既存のテーマに合うスキルなら、そのプラグインの `skills/` に追加する。新しいテーマなら `plugins/<プラグイン名>/` を作る
2. 新しいプラグインを作ったら、`.claude-plugin/marketplace.json` の `plugins` に1件追加する（`name` は `plugin.json` と同じにする）
3. SKILL.md の中では、スキルのファイルを **`${CLAUDE_SKILL_DIR}/scripts/...`** で参照する（プラグインでも手動インストールでも同じ場所を指す）。`~/.claude/skills/...` のような固定パスは使わない
4. 利用者ごとの情報は SKILL.md に書かず、プロフィールの新しい節（`## <プラグイン名>`）から読む
5. `claude plugin validate .` が通ることを確認してから push する
6. 公開リポジトリなので、**個人名・所属・未発表の研究データ・実在の資料を入れない**（テンプレートは匿名化したものだけ）

## ライセンス

[MIT](LICENSE)

---

## English summary

A collection of [Claude Code](https://code.claude.com/) skills, written in Japanese and distributed as plugins grouped by theme.

- **research-slides** — `create-doc` turns research logs, result CSVs and figures into a slide-by-slide outline (Markdown). `progress-report-deck` builds a 4:3 lab progress-report PowerPoint from that outline. Both follow fixed design rules: 16pt minimum text, a two-color theme palette and OMML equations. Requires Python 3.10+ and Node.js. PowerPoint on Windows is optional; it is used for the final visual check.

Install it with `/plugin marketplace add yapi777/yapi-skills`, then `/plugin install research-slides@yapi-skills`. Licensed under MIT.
