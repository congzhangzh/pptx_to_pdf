@echo off
cd /d %~dp0
chcp 65001
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
:loop
echo [supervisor] starting fake_ocr.py...
uv run python fake_ocr.py
echo [supervisor] worker exited (code %ERRORLEVEL%), restarting in 2s...
timeout /t 2 /nobreak >nul
goto loop
