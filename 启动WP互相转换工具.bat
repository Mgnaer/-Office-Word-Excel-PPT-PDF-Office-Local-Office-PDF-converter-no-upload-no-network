@echo off
setlocal
title WP Converter Launcher
REM ============================================================
REM  启动前先做文件完整性校验（防篡改）
REM  校验不通过则拒绝启动，避免运行被替换过的脚本
REM ============================================================
if not exist "%~dp0SHA256校验清单.txt" goto missing_manifest
if not exist "%~dp0verify_hashes.ps1" goto missing_manifest

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0verify_hashes.ps1"
if errorlevel 1 (
  echo.
  echo [失败] 文件完整性校验未通过，请勿运行！
  echo        请重新下载发布包，或联系提供者核对 SHA256。
  pause
  exit /b 1
)

REM 校验通过，交给启动引导：检测 Python（含 tkinter），缺失时可自动安装
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0bootstrap.ps1" %*
if errorlevel 1 (
  echo.
  echo [启动失败] 请查看上面的提示信息。
  pause
  exit /b 1
)
exit /b 0

:missing_manifest
echo.
echo [失败] 缺少完整性校验文件（SHA256校验清单.txt 或 verify_hashes.ps1）。
echo        为安全起见已中止启动，请重新下载完整发布包。
pause
exit /b 1
