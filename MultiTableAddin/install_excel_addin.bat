@echo off
setlocal
title Add-in Script Launcher
set "SCRIPT_PATH=%~dp0scripts\install_excel_addin.ps1"
if not exist "%SCRIPT_PATH%" set "SCRIPT_PATH=%~dp0files\scripts\install_excel_addin.ps1"
echo ========================================
echo Running script:
echo   scripts\install_excel_addin.ps1
echo ========================================
echo.
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_PATH%" %*
set EXITCODE=%ERRORLEVEL%
echo.
if "%EXITCODE%"=="0" (
  echo Operation completed.
) else (
  echo Operation failed. Exit code: %EXITCODE%
)
echo.
pause
exit /b %EXITCODE%
