@echo off
setlocal

rem Remove the user-level SendTo shortcut without requiring admin rights.
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$shortcutPath=Join-Path $env:APPDATA 'Microsoft\Windows\SendTo\Bulk Print.lnk'; if(Test-Path -LiteralPath $shortcutPath){Remove-Item -LiteralPath $shortcutPath -Force; Write-Host 'Removed: Send to -> Bulk Print' -ForegroundColor Green}else{Write-Host 'The Bulk Print SendTo shortcut was not found.' -ForegroundColor Yellow}"

if errorlevel 1 (
    echo.
    echo Uninstallation failed. Review the PowerShell error above.
)

pause
endlocal
