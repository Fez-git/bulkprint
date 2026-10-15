param(
    [string]$SourceUrl = 'https://raw.githubusercontent.com/Fez-git/bulkprint/main/Bulk-Print.ps1'
)

$ErrorActionPreference = 'Stop'
$installDir = $PSScriptRoot
$scriptPath = Join-Path $installDir 'Bulk-Print.ps1'

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

        if ((Test-Path -LiteralPath $scriptPath) -and
            (Get-FileHash -LiteralPath $tempPath).Hash -eq (Get-FileHash -LiteralPath $scriptPath).Hash) {
            Write-Host 'Bulk-Print.ps1 is already the latest version.' -ForegroundColor Green
            return
        }

        if (Test-Path -LiteralPath $scriptPath) {
            Copy-Item -LiteralPath $scriptPath -Destination "$scriptPath.bak" -Force
        }
        Move-Item -LiteralPath $tempPath -Destination $scriptPath -Force
        Write-Host 'Bulk-Print.ps1 updated to the latest version from GitHub (previous copy kept as Bulk-Print.ps1.bak).' -ForegroundColor Green
    }
    catch {
        Write-Host "Could not download the latest Bulk-Print.ps1: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host 'Keeping the copy already in this folder.' -ForegroundColor Yellow
    }
    finally {
        if (Test-Path -LiteralPath $tempPath) { Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue }
    }
}

Update-BulkPrint

if (-not (Test-Path -LiteralPath $scriptPath)) {
    throw 'Bulk-Print.ps1 was not found and could not be downloaded.'
}

# Install the user-level SendTo shortcut without requiring admin rights.
$sendToPath = Join-Path $env:APPDATA 'Microsoft\Windows\SendTo'
$shortcutPath = Join-Path $sendToPath 'Bulk Print.lnk'
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
