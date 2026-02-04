@echo off
:: This script will run your PowerShell script with bypass and keep the window open
:: Check if PowerShell exists
where powershell >nul 2>&1
if %errorlevel% neq 0 (
    echo PowerShell not found!
    pause
    exit /b
)

:: Run the PowerShell script with bypass and elevation
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "& {Start-Process powershell.exe -ArgumentList '-NoProfile -ExecutionPolicy Bypass -NoExit -File ""%~dp0Windows-Update.ps1""' -Verb RunAs -WindowStyle Normal}"
pause