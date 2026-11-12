#!/usr/bin/env bash
# yapi-skills のスキルを ~/.claude/skills/ に手動でインストールする（プラグインを使わない場合）。
#   使い方: bash setup.sh                  # すべてのプラグインのスキルを入れる
#           bash setup.sh research-slides  # 指定したプラグインのスキルだけ入れる
# 同じ名前のスキルが既にあれば、消さずに <名前>.bak-<日時> に退避してから入れ直す。
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
DEST="$HOME/.claude/skills"
STAMP="$(date +%Y%m%d-%H%M%S)"
mkdir -p "$DEST"

if [ "$#" -gt 0 ]; then
  PLUGINS=("$@")
else
  PLUGINS=()
  for p in "$HERE"/plugins/*/; do PLUGINS+=("$(basename "$p")"); done
fi

for plugin in "${PLUGINS[@]}"; do
  PDIR="$HERE/plugins/$plugin"
  if [ ! -d "$PDIR/skills" ]; then
    echo "ERROR: プラグイン '$plugin' が見つかりません（plugins/ の中を確認してください）" >&2
    exit 1
  fi
  echo "== $plugin"
  for d in "$PDIR"/skills/*/; do
    [ -f "$d/SKILL.md" ] || continue
    name="$(basename "$d")"
    if [ -e "$DEST/$name" ]; then
      mv "$DEST/$name" "$DEST/$name.bak-$STAMP"
      echo "  退避: $DEST/$name -> $name.bak-$STAMP"
    fi
    cp -r "$d" "$DEST/$name"
    echo "  installed: $DEST/$name"
  done

  if [ -f "$PDIR/requirements.txt" ]; then
    # 実際に動く Python を選ぶ（Windows の python3 は Microsoft Store の空のエイリアスのことがある）
    PY=""
    for c in python3 python py; do
      if command -v "$c" >/dev/null 2>&1 && "$c" -c "import sys; assert sys.version_info >= (3, 10)" >/dev/null 2>&1; then
        PY="$c"; break
      fi
    done
    if [ -n "$PY" ]; then
      "$PY" -m pip install --quiet -r "$PDIR/requirements.txt" && echo "  pip packages: OK"
    else
      echo "  WARN: Python 3.10+ が見つかりません。入れてから pip install -r plugins/$plugin/requirements.txt を実行してください。"
    fi
  fi
done

command -v node >/dev/null && echo "node: $(node --version)" || echo "WARN: Node.js が見つかりません（research-slides のスクリプトに必要）"
echo "完了。Claude Code を再起動すると、入れたスキルが /<スキル名> で使えます。"
