@echo off
:loop
echo [supervisor] starting worker...
uv run python main.py
echo [supervisor] worker exited (code %ERRORLEVEL%), restarting in 2s...
timeout /t 2 /nobreak >nul
goto loop
