# yapi-skills のスキルを %USERPROFILE%\.claude\skills\ に手動でインストールする（プラグインを使わない場合）。
#   使い方: powershell -ExecutionPolicy Bypass -File setup.ps1                   # すべてのプラグイン
#           powershell -ExecutionPolicy Bypass -File setup.ps1 research-slides   # 指定したプラグインだけ
# 同じ名前のスキルが既にあれば、消さずに <名前>.bak-<日時> に退避してから入れ直す。
param([string[]]$Plugins)
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$dest = Join-Path $env:USERPROFILE ".claude\skills"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
New-Item -ItemType Directory -Force $dest | Out-Null

if (-not $Plugins -or $Plugins.Count -eq 0) {
  $Plugins = Get-ChildItem (Join-Path $here "plugins") -Directory | ForEach-Object { $_.Name }
}

foreach ($plugin in $Plugins) {
  $pdir = Join-Path $here "plugins\$plugin"
  if (-not (Test-Path (Join-Path $pdir "skills"))) { throw "プラグイン '$plugin' が見つかりません（plugins\ の中を確認してください）" }
  Write-Output "== $plugin"
  Get-ChildItem (Join-Path $pdir "skills") -Directory | Where-Object { Test-Path (Join-Path $_.FullName "SKILL.md") } | ForEach-Object {
    $t = Join-Path $dest $_.Name
    if (Test-Path $t) {
      Rename-Item $t "$($_.Name).bak-$stamp"
      Write-Output "  退避: $t -> $($_.Name).bak-$stamp"
    }
    Copy-Item $_.FullName $t -Recurse
    Write-Output "  installed: $t"
  }

  $req = Join-Path $pdir "requirements.txt"
  if (Test-Path $req) {
    $py = Get-Command python -ErrorAction SilentlyContinue
    if (-not $py) { $py = Get-Command py -ErrorAction SilentlyContinue }
    if ($py) { & $py.Source -m pip install --quiet -r $req; Write-Output "  pip packages: OK" }
    else { Write-Warning "Python 3.10+ が見つかりません。インストール後に pip install -r plugins\$plugin\requirements.txt を実行してください。" }
  }
}

if (-not (Get-Command node -ErrorAction SilentlyContinue)) { Write-Warning "Node.js が見つかりません（research-slides のスクリプトに必要）" }
Write-Output "完了。Claude Code を再起動すると、入れたスキルが /<スキル名> で使えます。"
