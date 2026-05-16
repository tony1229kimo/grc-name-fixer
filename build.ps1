<#
.SYNOPSIS
    打包 GRC 團名修正工具 + 可選發佈 GitHub Release。

.DESCRIPTION
    步驟:
      1. 讀 _version.py 取得 VERSION
      2. pyinstaller grc-tool.spec --clean → dist\GRC-團名修正工具\
      3. 壓成 dist\GRC-團名修正工具-v{VERSION}.zip
      4. (可選 -Release) gh release create v{VERSION} 上傳 zip,同事就會在 web UI 看到更新通知

.EXAMPLE
    .\build.ps1
        只打包 + 壓 zip,不發 release。用來自己測試。

.EXAMPLE
    .\build.ps1 -Release
        打包 + 壓 zip + 發 GitHub Release v{VERSION}。
        ⚠ 發完所有同事下次開 app 就會看到更新通知。

.EXAMPLE
    .\build.ps1 -Release -ReleaseNotes "修了下載 bug + 加新版通知"
        同上,但帶自訂的 release 說明。

.NOTES
    需要:
      - pip install pyinstaller
      - gh (GitHub CLI) + gh auth login  (僅 -Release 模式需要)
#>
[CmdletBinding()]
param(
    [switch]$Release,
    [string]$ReleaseNotes = "",
    [switch]$SkipBuild   # 已經有 dist 了不想重 build,只想重新打 zip / 發 release
)

$ErrorActionPreference = 'Stop'
Set-Location -Path $PSScriptRoot

# ---------- 1. 讀 VERSION ----------
$versionLine = Select-String -Path '_version.py' -Pattern '^VERSION\s*=\s*"([^"]+)"' | Select-Object -First 1
if (-not $versionLine) {
    Write-Error "找不到 _version.py 裡的 VERSION 設定"
}
$version = $versionLine.Matches[0].Groups[1].Value
$tag = "v$version"
Write-Host ""
Write-Host "===== GRC 團名修正工具 build =====" -ForegroundColor Cyan
Write-Host "  Version : $version"
Write-Host "  Tag     : $tag"
Write-Host "  Release : $($Release.IsPresent)"
Write-Host "=================================" -ForegroundColor Cyan
Write-Host ""

# ---------- 2. PyInstaller ----------
# dist 資料夾名稱保留中文(同事解壓後看到的資料夾)
# zip 檔名用 ASCII(避免 GitHub 把中文 sanitize 成 . 變成「GRC-.-v1.1.0.zip」)
$distDir = 'dist\GRC-團名修正工具'
$zipPath = "dist\grc-name-fixer-$tag.zip"

if (-not $SkipBuild) {
    # 確認 pyinstaller 可用
    $pyi = Get-Command pyinstaller -ErrorAction SilentlyContinue
    if (-not $pyi) {
        Write-Error "找不到 pyinstaller。請先跑: pip install pyinstaller"
    }

    Write-Host "[1/3] PyInstaller 打包中..." -ForegroundColor Yellow
    & pyinstaller grc-tool.spec --clean --noconfirm
    if ($LASTEXITCODE -ne 0) {
        Write-Error "PyInstaller 失敗 (exit $LASTEXITCODE)"
    }
    Write-Host "      ✓ 完成 → $distDir" -ForegroundColor Green
} else {
    Write-Host "[1/3] -SkipBuild 跳過 PyInstaller (用既有的 $distDir)" -ForegroundColor DarkGray
}

if (-not (Test-Path $distDir)) {
    Write-Error "找不到 $distDir,無法繼續"
}

# ---------- 3. 壓 zip ----------
Write-Host "[2/3] 壓 zip..." -ForegroundColor Yellow
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
Compress-Archive -Path $distDir -DestinationPath $zipPath -CompressionLevel Optimal
$zipSize = [math]::Round((Get-Item $zipPath).Length / 1MB, 1)
Write-Host "      ✓ 完成 → $zipPath ($zipSize MB)" -ForegroundColor Green

# ---------- 4. (可選) 發 GitHub Release ----------
if (-not $Release) {
    Write-Host ""
    Write-Host "[3/3] 略過 (沒加 -Release)" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "👉 確認沒問題後,跑 .\build.ps1 -Release 發 GitHub Release" -ForegroundColor Cyan
    return
}

# 確認 gh 可用
$gh = Get-Command gh -ErrorAction SilentlyContinue
if (-not $gh) {
    Write-Error "找不到 gh CLI。請先裝: winget install GitHub.cli"
}

# 確認 gh 已登入(否則 release create 會吐 exit 4 但不講清楚)
& gh auth status 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Error @"
gh CLI 還沒登入 GitHub。請先跑:
    gh auth login
然後選 GitHub.com → HTTPS → Login with a web browser,
照指示完成後再重跑本腳本。
"@
}

# 確認 git 是乾淨的(避免發了 release 但 code 沒同步)
$gitStatus = git status --porcelain
if ($gitStatus) {
    Write-Warning "Git 有未 commit 的改動:"
    git status --short
    $cont = Read-Host "仍要發 release? (y/N)"
    if ($cont -ne 'y' -and $cont -ne 'Y') {
        Write-Host "取消" -ForegroundColor Yellow
        return
    }
}

# 確認 tag 還沒被用過
$existingTag = git tag --list $tag
if ($existingTag) {
    Write-Error "Tag $tag 已存在。請先更新 _version.py 的 VERSION,再重跑。"
}

# Release notes:預設用最近一個 commit 訊息
if (-not $ReleaseNotes) {
    $ReleaseNotes = git log -1 --pretty=format:"%s%n%n%b"
}

Write-Host ""
Write-Host "[3/3] 發 GitHub Release $tag..." -ForegroundColor Yellow
& gh release create $tag $zipPath `
    --title "GRC 團名修正工具 $tag" `
    --notes $ReleaseNotes
if ($LASTEXITCODE -ne 0) {
    Write-Error "gh release create 失敗 (exit $LASTEXITCODE)"
}

Write-Host ""
Write-Host "      ✓ Release $tag 已發佈!" -ForegroundColor Green
Write-Host ""
Write-Host "===== 下一步 =====" -ForegroundColor Cyan
Write-Host "  • 所有同事下次開 app,web UI 會自動跳新版通知"
Write-Host "  • 通知有快取 4 小時,如果同事的 app 開著沒關,重整網頁可以提早看到"
Write-Host "  • Release 頁面: https://github.com/tony1229kimo/grc-name-fixer/releases/tag/$tag" -ForegroundColor Blue
Write-Host ""
