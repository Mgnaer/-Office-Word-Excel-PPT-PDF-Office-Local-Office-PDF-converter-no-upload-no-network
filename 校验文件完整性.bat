@echo off
setlocal
title 文件完整性校验
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0verify_hashes.ps1"
if errorlevel 1 (
  echo.
  echo [失败] 文件校验未通过，请勿运行！请重新下载。
  pause
) else (
  echo.
  echo [通过] 全部文件校验一致，可以放心使用。
  timeout /t 3 >nul 2>&1
)
