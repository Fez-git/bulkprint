@echo off
setlocal

set "BULK_PRINT_DIR=%~dp0"

rem Install the user-level SendTo shortcut without requiring admin rights.
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; $scriptPath=Join-Path $env:BULK_PRINT_DIR 'Bulk-Print.ps1'; if(-not (Test-Path -LiteralPath $scriptPath)){throw 'Bulk-Print.ps1 was not found.'}; $sendToPath=Join-Path $env:APPDATA 'Microsoft\Windows\SendTo'; $shortcutPath=Join-Path $sendToPath 'Bulk Print.lnk'; $powershellPath=Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'; $shell=New-Object -ComObject WScript.Shell; $shortcut=$shell.CreateShortcut($shortcutPath); $q=[char]34; $shortcut.TargetPath=$powershellPath; $shortcut.Arguments='-NoProfile -ExecutionPolicy Bypass -File '+$q+$scriptPath+$q; $shortcut.WorkingDirectory=$env:BULK_PRINT_DIR; $shortcut.Description='Print supported files in the selected folder'; $shortcut.IconLocation=$powershellPath+',0'; $shortcut.Save(); Write-Host 'Installed: Send to -> Bulk Print' -ForegroundColor Green"

if errorlevel 1 (
    echo.
    echo Installation failed. Review the PowerShell error above.
)

pause
endlocal
