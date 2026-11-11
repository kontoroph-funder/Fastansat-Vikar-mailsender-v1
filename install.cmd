@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"
set "installResult=%errorlevel%"
if "%installResult%"=="1" echo Installationen fejlede. Se fejlen ovenfor.
pause
exit /b %installResult%
