@echo off
chcp 65001
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
set OCR_URL=http://localhost:9000/ocr
:loop
echo [supervisor] starting service.py...
uv run python service.py
echo [supervisor] worker exited (code %ERRORLEVEL%), restarting in 2s...
timeout /t 2 /nobreak >nul
goto loop
