@echo off
REM Run the arr_finisher test suite.
REM Usage:
REM   run_tests.bat          - fast tests only (no network)
REM   run_tests.bat all      - includes integration tests (hits live APIs)

setlocal
cd /d "%~dp0\.."
if /I "%1"=="all" (
    set NETWORK_TESTS=1
)
python -m pytest tests -v
endlocal
