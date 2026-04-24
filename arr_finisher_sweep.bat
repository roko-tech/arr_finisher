@echo off
REM Runs arr_finisher --sweep silently. Intended for Windows Task Scheduler.
python "%~dp0arr_finisher.py" --sweep >nul 2>&1
exit /b 0
