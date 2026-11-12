<#
  verify_hashes.ps1 —— 校验本目录所有文件的 SHA256 是否与清单一致
  用法：powershell -NoProfile -ExecutionPolicy Bypass -File verify_hashes.ps1
  作用：把“下载的 zip 解压后，确认文件没被中间人/磁盘错误篡改”这件事变成一键操作。
#>
$ErrorActionPreference = "Continue"
$dir = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($dir)) { $dir = Split-Path -Parent $MyInvocation.MyCommand.Path }
$manifest = Join-Path $dir "SHA256校验清单.txt"

if (-not (Test-Path -LiteralPath $manifest)) {
  Write-Host ("找不到清单文件：" + $manifest) -ForegroundColor Red
  Write-Host "请确认 SHA256校验清单.txt 与本脚本在同一目录。" -ForegroundColor Yellow
  exit 1
}

$okCount = 0
$failCount = 0
$lines = Get-Content -LiteralPath $manifest -Encoding UTF8

foreach ($raw in $lines) {
  $line = $raw.Trim()
  if (-not $line -or $line.StartsWith("#")) { continue }
  $parts = @($line -split "\s+", 2)
  if ($parts.Count -ne 2) { continue }
  $expected = $parts[0].ToLower().Trim()
  $name = $parts[1].Trim()
  $path = Join-Path $dir $name

  if (-not (Test-Path -LiteralPath $path)) {
    Write-Host ("[缺失] " + $name) -ForegroundColor Red
    $failCount++
    continue
  }
  $actual = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLower()
  if ($actual -eq $expected) {
    Write-Host ("[ OK ] " + $name) -ForegroundColor Green
    $okCount++
  } else {
    Write-Host ("[篡改] " + $name) -ForegroundColor Red
    Write-Host ("        期望 " + $expected) -ForegroundColor Gray
    Write-Host ("        实际 " + $actual) -ForegroundColor Red
    $failCount++
  }
}

Write-Host ""
if ($failCount -eq 0) {
  Write-Host ("校验通过：共 " + $okCount + " 个文件，全部与清单一致。") -ForegroundColor Green
  exit 0
} else {
  Write-Host ("校验失败：" + $failCount + " 个文件不匹配或缺失。请勿运行，重新下载！") -ForegroundColor Red
  exit 1
}
