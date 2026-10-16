# Builds the release package into dist\:
#   dist\BulkPrint.zip     the download for new users (stable name: the Releases "latest" link depends on it)
#   dist\Bulk-Print.ps1    the same script as inside the ZIP, attached to the release so the installer can update it
# Run from the repo root:  powershell -NoProfile -ExecutionPolicy Bypass -File build-release.ps1
param([string]$OutDir = (Join-Path $PSScriptRoot 'dist'))

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$root = $PSScriptRoot
$scriptFile = Join-Path $root 'Bulk-Print.ps1'
$match = Select-String -LiteralPath $scriptFile -Pattern '^\$BulkPrintVersion\s*=\s*''([0-9.]+)''' | Select-Object -First 1
if ($null -eq $match) { throw 'Could not find $BulkPrintVersion in Bulk-Print.ps1' }
$version = $match.Matches[0].Groups[1].Value

# Plain-text files for the ZIP are written with Windows line endings so Notepad and cmd.exe read them properly.
function ConvertTo-Crlf([string]$Text) { return (($Text -replace "`r`n", "`n") -replace "`n", "`r`n") }
$installCmd = ConvertTo-Crlf @'
@echo off
setlocal

rem Install Bulk Print: adds Send to -> Bulk Print to the right-click menu (no admin rights needed).
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0files\Install-SendTo.ps1"

if errorlevel 1 (
    echo.
    echo Installation failed. Review the message above.
)

pause
endlocal
'@
$readme = ConvertTo-Crlf ((Get-Content -LiteralPath (Join-Path $root 'release\README.txt') -Raw) -replace '\{\{VERSION\}\}', $version)

# Entry name in the ZIP -> bytes. Scripts are taken as they are in the working tree, so the loose
# Bulk-Print.ps1 and the copy in the ZIP are byte-identical.
$utf8 = New-Object System.Text.UTF8Encoding($false)
$entries = [ordered]@{
    'Install.cmd'                = $utf8.GetBytes($installCmd)
    'README.txt'                 = $utf8.GetBytes($readme)
    'LICENSE'                    = [System.IO.File]::ReadAllBytes((Join-Path $root 'LICENSE'))
    'files/Bulk-Print.ps1'       = [System.IO.File]::ReadAllBytes($scriptFile)
    'files/Install-SendTo.ps1'   = [System.IO.File]::ReadAllBytes((Join-Path $root 'Install-SendTo.ps1'))
    'files/Setup-ImageColour.cmd' = [System.IO.File]::ReadAllBytes((Join-Path $root 'Setup-ImageColour.cmd'))
    'files/Uninstall-SendTo.cmd' = [System.IO.File]::ReadAllBytes((Join-Path $root 'Uninstall-SendTo.cmd'))
}

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$zipPath = Join-Path $OutDir 'BulkPrint.zip'
if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath }
$zip = [System.IO.Compression.ZipFile]::Open($zipPath, 'Create')
try {
    foreach ($name in $entries.Keys) {
        # Forward slashes: PowerShell 5.1's Compress-Archive writes backslashes, which some unzip tools reject.
        $entry = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
        $stream = $entry.Open()
        try { $stream.Write($entries[$name], 0, $entries[$name].Length) } finally { $stream.Dispose() }
    }
}
finally { $zip.Dispose() }

Copy-Item -LiteralPath $scriptFile -Destination (Join-Path $OutDir 'Bulk-Print.ps1') -Force

# Check what was built, by reading it back.
$check = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
try { $names = @($check.Entries | ForEach-Object FullName) } finally { $check.Dispose() }
if (@($names | Where-Object { $_ -like '*\*' }).Count -gt 0) { throw 'The ZIP contains backslash entry names' }
$scriptHash = (Get-FileHash -LiteralPath $scriptFile).Hash
Write-Host "Built Bulk Print $version" -ForegroundColor Green
Write-Host "  $zipPath"
$names | ForEach-Object { Write-Host "    $_" }
Write-Host "  $(Join-Path $OutDir 'Bulk-Print.ps1')"
Write-Host "  Bulk-Print.ps1 SHA256 $scriptHash"
Write-Host "  BulkPrint.zip  SHA256 $((Get-FileHash -LiteralPath $zipPath).Hash)"
