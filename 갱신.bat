@echo off
chcp 65001 >nul
title Dashboard refresh
rem 시작 전에 자동 배치가 올린 최신 커밋을 받아 충돌 가능성을 줄인다 (실패해도 무시).
git -C "%~dp0." pull --ff-only >nul 2>nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0refresh.ps1"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0publish.ps1"
echo.
echo Press any key to close this window.
pause >nul
