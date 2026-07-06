@echo off
setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0"

set "PS1FILE=%~dp0RepoGrandmaCH.ps1"

if not exist "%PS1FILE%" (
    echo Error: PowerShell script not found.
    echo Please make sure the .bat file and RepoGrandmaCH.ps1 are in the same folder.
    echo.
    powershell.exe -NoLogo -NoProfile -Command "Write-Host '執行完畢，按Enter鍵以關閉本視窗' -ForegroundColor Green"
    pause >nul
    exit /b 1
)

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%PS1FILE%" -OpenOutputFolder

set "ERR=%ERRORLEVEL%"
echo.
powershell.exe -NoLogo -NoProfile -Command "Write-Host '執行完畢，按Enter鍵以關閉本視窗' -ForegroundColor Green"
pause >nul
exit /b %ERR%