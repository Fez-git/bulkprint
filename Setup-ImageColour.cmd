@echo off
setlocal

rem One-time per PC: open the default printer's own settings, choose Colour, and save them for image printing.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Bulk-Print.ps1" -SetupImageColour

endlocal
