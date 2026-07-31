@echo off
chcp 65001 >nul
title Dashboard refresh
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0refresh.ps1"
echo.
echo Press any key to close this window.
pause >nul
