@echo off
setlocal

rem Download the latest Bulk-Print.ps1 from GitHub, then install the user-level SendTo shortcut (no admin rights).
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-SendTo.ps1"

if errorlevel 1 (
    echo.
    echo Installation failed. Review the PowerShell error above.
)

pause
endlocal
