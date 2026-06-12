@echo off
REM AG FlexPoints - pull latest script from GitHub, then run.
REM Never run a stale local copy.
cd /d "%~dp0"
git fetch origin
git checkout origin/main -- fetch_flexpoints.py
python fetch_flexpoints.py
if errorlevel 1 pause
