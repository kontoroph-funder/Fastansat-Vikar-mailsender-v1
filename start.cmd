@echo off
cd /d "%~dp0"
if exist "FastansatVikar\FastansatVikar.exe" (
  start "" "FastansatVikar\FastansatVikar.exe"
) else (
  ".venv\Scripts\python.exe" -m app
)
