#Requires -Version 5.1
<#
.SYNOPSIS
  首次将「2026 语文教案」从 GitHub 克隆到本机 Documents 目录。

.DESCRIPTION
  目标路径：C:\Users\Bechodan1209\Documents\2026 语文教案
  云端仓库：https://github.com/decemberlin54-del/declin
#>

$ErrorActionPreference = "Stop"

$RepoUrl   = "https://github.com/decemberlin54-del/declin.git"
$ParentDir = Join-Path $env:USERPROFILE "Documents"
$TargetDir = Join-Path $ParentDir "2026 语文教案"

Write-Host "=== 2026 语文教案 · 首次安装 ===" -ForegroundColor Cyan
Write-Host "目标路径: $TargetDir"

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    Write-Host "未检测到 Git。请先安装: https://git-scm.com/download/win" -ForegroundColor Red
    Read-Host "按 Enter 退出"
    exit 1
}

if (-not (Test-Path $ParentDir)) {
    New-Item -ItemType Directory -Path $ParentDir -Force | Out-Null
}

if (Test-Path $TargetDir) {
    if (Test-Path (Join-Path $TargetDir ".git")) {
        Write-Host "目录已存在且为 Git 仓库，正在拉取最新内容..." -ForegroundColor Yellow
        Set-Location $TargetDir
        git pull
    } else {
        Write-Host "错误: $TargetDir 已存在但不是 Git 仓库。请手动处理后再运行。" -ForegroundColor Red
        Read-Host "按 Enter 退出"
        exit 1
    }
} else {
    Write-Host "正在从 GitHub 克隆..." -ForegroundColor Green
    Set-Location $ParentDir
    git clone $RepoUrl "2026 语文教案"
    Set-Location $TargetDir
}

Write-Host ""
Write-Host "安装完成！" -ForegroundColor Green
Write-Host "下一步: 用 Cursor 打开文件夹 -> $TargetDir" -ForegroundColor Green
Write-Host "       File -> Open Folder -> 选择「2026 语文教案」" -ForegroundColor Green

$open = Read-Host "是否尝试用 Cursor 打开？(Y/n)"
if ($open -ne "n" -and $open -ne "N") {
    $cursor = Get-Command cursor -ErrorAction SilentlyContinue
    if ($cursor) {
        cursor $TargetDir
    } else {
        Write-Host "未找到 cursor 命令。请手动在 Cursor 中打开上述路径。" -ForegroundColor Yellow
    }
}

Read-Host "按 Enter 退出"
