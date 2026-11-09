@echo off
setlocal EnableExtensions
set "REPO_ROOT=%~dp0.."
cd /d "%REPO_ROOT%"

set "SCRIPT=%~dp0Certify-Excel.ps1"

if not exist "%SCRIPT%" (
    echo ERROR: PowerShell script not found:
    echo   %SCRIPT%
    echo.
    pause
    exit /b 1
)

rem Certifies the checked-out commit in Excel: imports the configured VBA into
rem a new workbook, compiles it, runs every suite and example, and writes a
rem validated evidence bundle next to the checkout. See docs\EXCEL_EVIDENCE.md.
rem Extra arguments are passed to the script, for example -RepoRoot C:\project.
echo Excel certification
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" %*
set "CODE=%ERRORLEVEL%"
echo.
pause
exit /b %CODE%
