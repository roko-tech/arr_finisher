@echo off
REM Sonarr/Radarr custom script launcher. Runs silently.
REM Use this .bat as the custom-script path in Sonarr and Radarr.

python "%~dp0arr_finisher.py" >nul 2>&1
exit /b 0
