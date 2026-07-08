#Requires -Version 5.1
<#
.SYNOPSIS
  将本机修改保存并推送到 GitHub 云端。

.DESCRIPTION
  在 C:\Users\Bechodan1209\Documents\2026 语文教案 目录执行 git add / commit / push。
#>

$ErrorActionPreference = "Stop"

$TargetDir = Join-Path $env:USERPROFILE "Documents\2026 语文教案"

Write-Host "=== 保存并上传云端 ===" -ForegroundColor Cyan
Write-Host "工作目录: $TargetDir"

if (-not (Test-Path (Join-Path $TargetDir ".git"))) {
    Write-Host "未找到 Git 仓库。请先运行「首次安装到本机.ps1」" -ForegroundColor Red
    Read-Host "按 Enter 退出"
    exit 1
}

Set-Location $TargetDir

$status = git status --porcelain
if (-not $status) {
    Write-Host "没有需要保存的修改。" -ForegroundColor Yellow
    Read-Host "按 Enter 退出"
    exit 0
}

Write-Host ""
Write-Host "以下文件将被保存:" -ForegroundColor Yellow
git status -s
Write-Host ""

$msg = Read-Host "请输入本次更新说明（直接回车则用默认说明）"
if ([string]::IsNullOrWhiteSpace($msg)) {
    $msg = "更新教案 $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
}

git add .
git commit -m $msg
git push

Write-Host ""
Write-Host "已上传到 GitHub 云端！" -ForegroundColor Green
Write-Host "换电脑后运行「首次安装到本机.ps1」即可恢复全部文件。" -ForegroundColor Green

Read-Host "按 Enter 退出"
