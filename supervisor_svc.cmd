@echo off
set OCR_URL=http://localhost:9000/ocr
:loop
echo [supervisor] starting service.py...
uv run python service.py
echo [supervisor] worker exited (code %ERRORLEVEL%), restarting in 2s...
timeout /t 2 /nobreak >nul
goto loop
