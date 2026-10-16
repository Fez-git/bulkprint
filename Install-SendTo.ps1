param(
    # Always the newest published release. The asset name stays the same across releases, so this link never changes.
    [string]$SourceUrl = 'https://github.com/Fez-git/bulkprint/releases/latest/download/Bulk-Print.ps1'
)

$ErrorActionPreference = 'Stop'
$installDir = $PSScriptRoot
$scriptPath = Join-Path $installDir 'Bulk-Print.ps1'

function Get-ScriptVersion {
    param([string]$Path)

    # Reads the line  $BulkPrintVersion = 'x.y.z'  without running the script.
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    $match = Select-String -LiteralPath $Path -Pattern '^\$BulkPrintVersion\s*=\s*''([0-9]+(\.[0-9]+){1,3})''' | Select-Object -First 1
    if ($null -eq $match) { return $null }
    return [version]$match.Matches[0].Groups[1].Value
}

function Update-BulkPrint {
    if (Test-Path -LiteralPath (Join-Path $installDir '.git')) {
        Write-Host 'This folder is a git clone: update it with "git pull" instead. Skipping download.' -ForegroundColor Yellow
        return
    }

    # Download to a temp file first: a failed or partial download must never replace a working script.
    $tempPath = Join-Path ([System.IO.Path]::GetTempPath()) ("Bulk-Print-" + [guid]::NewGuid().ToString('N') + '.ps1')
    try {
        # Windows PowerShell 5.1 on older Windows 10 builds does not offer TLS 1.2 by default; GitHub requires it.
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $SourceUrl -OutFile $tempPath -UseBasicParsing

        # Refuse anything that is not a complete Bulk-Print.ps1 (error page, cut-off download).
        $parseErrors = $null
        [void][System.Management.Automation.Language.Parser]::ParseFile($tempPath, [ref]$null, [ref]$parseErrors)
        if ($parseErrors.Count -gt 0) { throw 'the downloaded file is not valid PowerShell' }
        if (-not (Select-String -LiteralPath $tempPath -Pattern 'function Print-OfficeFile' -SimpleMatch -Quiet)) {
            throw 'the downloaded file is not Bulk-Print.ps1'
        }

        $remoteVersion = Get-ScriptVersion $tempPath
        $localVersion = Get-ScriptVersion $scriptPath
        $versionText = if ($remoteVersion) { " $remoteVersion" } else { '' }

        if ((Test-Path -LiteralPath $scriptPath) -and
            (Get-FileHash -LiteralPath $tempPath).Hash -eq (Get-FileHash -LiteralPath $scriptPath).Hash) {
            Write-Host "Bulk-Print.ps1 is already the latest version$versionText." -ForegroundColor Green
            return
        }

        # Never go backwards: a copy newer than the published release is kept (for example one built from source).
        if ($localVersion -and $remoteVersion -and $localVersion -gt $remoteVersion) {
            Write-Host "This copy of Bulk-Print.ps1 ($localVersion) is newer than the latest release ($remoteVersion). Keeping it." -ForegroundColor Yellow
            return
        }

        if (Test-Path -LiteralPath $scriptPath) {
            Copy-Item -LiteralPath $scriptPath -Destination "$scriptPath.bak" -Force
        }
        Move-Item -LiteralPath $tempPath -Destination $scriptPath -Force
        $fromText = if ($localVersion) { "$localVersion" } else { 'unversioned' }
        $toText = if ($remoteVersion) { "$remoteVersion" } else { 'unversioned' }
        Write-Host "Bulk-Print.ps1 updated to the latest release ($fromText -> $toText). The previous copy is kept as Bulk-Print.ps1.bak." -ForegroundColor Green
    }
    catch {
        Write-Host "Could not download the latest Bulk-Print.ps1: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host 'Keeping the copy already in this folder.' -ForegroundColor Yellow
    }
    finally {
        if (Test-Path -LiteralPath $tempPath) { Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue }
    }
}

function Write-Check {
    param([bool]$Ok, [string]$Message)
    if ($Ok) { Write-Host "  [ok]  $Message" -ForegroundColor Green }
    else { Write-Host "  [!!]  $Message" -ForegroundColor Yellow }
}

function Get-FileTypeProgId {
    param([string]$Extension)
    $choice = (Get-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$Extension\UserChoice" -ErrorAction SilentlyContinue).ProgId
    if ($choice) { return $choice }
    return (Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\$Extension" -ErrorAction SilentlyContinue).'(default)'
}

function Test-Prerequisites {
    # Reports only. Nothing here stops the install: a missing program only affects its own file type.
    Write-Host "`nChecking this PC (nothing here blocks the install):"

    foreach ($app in @(@('Word', 'Word.Application', '.doc/.docx'), @('Excel', 'Excel.Application', '.xls/.xlsx'), @('PowerPoint', 'PowerPoint.Application', '.ppt/.pptx'))) {
        $found = Test-Path -LiteralPath "Registry::HKEY_CLASSES_ROOT\$($app[1])\CLSID"
        Write-Check $found $(if ($found) { "$($app[0]) (desktop) found" } else { "$($app[0]) (desktop) not found - $($app[2]) files will fail. Office on the web and the Store edition cannot be used." })
    }

    $emlProgId = Get-FileTypeProgId '.eml'
    $emlCommand = if ($emlProgId) { (Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\$emlProgId\shell\open\command" -ErrorAction SilentlyContinue).'(default)' } else { $null }
    $classicOutlook = Test-Path -LiteralPath 'Registry::HKEY_CLASSES_ROOT\Outlook.Application\CLSID'
    $emlOk = $classicOutlook -and $emlCommand -and ($emlCommand -match 'OUTLOOK\.EXE')
    Write-Check $emlOk $(if ($emlOk) { '.eml files open in classic Outlook' } else { '.eml files do not open in classic Outlook - emails will fail. Needs classic Outlook as the program that opens .eml (the new Outlook app cannot be scripted).' })

    $pdfProgId = Get-FileTypeProgId '.pdf'
    $pdfPrint = $pdfProgId -and (Test-Path -LiteralPath "Registry::HKEY_CLASSES_ROOT\$pdfProgId\shell\print")
    Write-Check ([bool]$pdfPrint) $(if ($pdfPrint) { 'The PDF program registers a Print action' } else { 'The PDF program does not register a Print action - PDFs may only open on screen. Adobe Reader offers one.' })

    $printer = $null
    try { $printer = (Get-CimInstance Win32_Printer -Filter 'Default=TRUE' -ErrorAction Stop | Select-Object -First 1).Name } catch {}
    Write-Check ([bool]$printer) $(if ($printer) { "Default printer: $printer" } else { 'No default printer set - set one in Settings > Printers & scanners.' })
    Write-Host '  Documents print with that printer''s own colour setting; images request colour.'
}

# Files downloaded from the internet carry a "from the internet" mark; clear it so later double-clicks
# (Setup-ImageColour.cmd, Uninstall-SendTo.cmd) do not show a security warning.
try { Get-ChildItem -LiteralPath $installDir -Recurse -File -ErrorAction Stop | Unblock-File -ErrorAction SilentlyContinue } catch {}

Update-BulkPrint

if (-not (Test-Path -LiteralPath $scriptPath)) {
    throw 'Bulk-Print.ps1 was not found and could not be downloaded.'
}

# Install the user-level SendTo shortcut without requiring admin rights.
$sendToPath = Join-Path $env:APPDATA 'Microsoft\Windows\SendTo'
$shortcutPath = Join-Path $sendToPath 'Bulk Print.lnk'
New-Item -ItemType Directory -Force -Path $sendToPath | Out-Null
$powershellPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $powershellPath
$shortcut.Arguments = '-NoProfile -ExecutionPolicy Bypass -File "' + $scriptPath + '"'
$shortcut.WorkingDirectory = $installDir
$shortcut.Description = 'Print supported files in the selected folder'
$shortcut.IconLocation = $powershellPath + ',0'
$shortcut.Save()
Write-Host 'Installed: Send to -> Bulk Print' -ForegroundColor Green
Write-Host "Keep this folder where it is: the right-click shortcut points to $installDir" -ForegroundColor Yellow

Test-Prerequisites

Write-Host "`nIf images print in black and white, the printer's own driver is overriding the colour request."
Write-Host 'The colour setup fixes that: it opens the printer''s settings once, you choose Colour, and it is used for images only.'
$answer = Read-Host 'Run the colour setup now? (y/N)'
if ($answer.Trim().ToLowerInvariant() -eq 'y') {
    & $powershellPath -NoProfile -ExecutionPolicy Bypass -File $scriptPath -SetupImageColour
}
else {
    Write-Host 'Skipped. Run Setup-ImageColour.cmd any time if images come out in black and white.'
}
