@echo off
:: ─────────────────────────────────────────────────────────────────────────────
::  Medicus UWP Installer Launcher
::  Double-click this file to install Medicus.
::  It will automatically request Administrator privileges.
:: ─────────────────────────────────────────────────────────────────────────────

:: Check if already elevated
net session >nul 2>&1
if %errorLevel% == 0 goto :run

:: Not elevated — re-launch self as admin via PowerShell
echo Requesting Administrator privileges...
powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
exit /b

:run
:: Run the PowerShell installer, bypassing execution policy for this session only
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-Medicus-v3.ps1"

:: Keep window open if PowerShell exits with an error
if %errorLevel% neq 0 (
    echo.
    echo Something went wrong. Error code: %errorLevel%
    echo Please take a photo of this screen and share it for support.
    pause
)
