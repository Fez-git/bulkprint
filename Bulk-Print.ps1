param(
    # Send To passes every selected folder as a separate argument.
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$FolderPath
)

$ErrorActionPreference = 'Stop'
$supportedExtensions = @('.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.pdf')
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$script:SkippedForPrompt = New-Object 'System.Collections.Generic.List[string]'
$script:FailedForPrompt = New-Object 'System.Collections.Generic.List[string]'
$script:TempFolders = New-Object 'System.Collections.Generic.List[string]'
$script:LogPaths = New-Object 'System.Collections.Generic.List[string]'
$script:CurrentFolderName = ''

function Write-Log {
    param([string]$Message)
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $Message"
    Add-Content -LiteralPath $script:LogPath -Value $line
    Write-Host $line
}

function Add-SkippedFile {
    param(
        [string]$Name,
        [string]$Reason,
        [string]$Extension
    )

    Write-Log "SKIPPED  $Name -- $Reason"
    if ($Extension.ToLowerInvariant() -notin @('.eml', '.md')) {
        [void]$script:SkippedForPrompt.Add("[$script:CurrentFolderName] $Name -- $Reason")
    }
}

function Add-FailedFile {
    param(
        [string]$Name,
        [string]$Reason
    )

    Write-Log "FAILED   $Name -- $Reason"
    [void]$script:FailedForPrompt.Add("[$script:CurrentFolderName] $Name -- $Reason")
}

function Release-ComObject {
    param($Object)
    if ($null -ne $Object) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Object) } catch {}
    }
}

function Print-OfficeFile {
    param([System.IO.FileInfo]$File)

    $extension = $File.Extension.ToLowerInvariant()
    $app = $null
    $document = $null

    try {
        switch ($extension) {
            '.doc' {
                $app = New-Object -ComObject Word.Application
                $app.Visible = $false
                try { $app.AutomationSecurity = 3 } catch {}
                $document = $app.Documents.Open($File.FullName, $false, $true)
                # Background = $false: return only once spooled, so Quit cannot cancel the job.
                $document.PrintOut($false)
                $document.Close(0)
            }
            '.docx' {
                $app = New-Object -ComObject Word.Application
                $app.Visible = $false
                try { $app.AutomationSecurity = 3 } catch {}
                $document = $app.Documents.Open($File.FullName, $false, $true)
                $document.PrintOut($false)
                $document.Close(0)
            }
            '.xls' {
                $app = New-Object -ComObject Excel.Application
                $app.Visible = $false
                $app.DisplayAlerts = $false
                try { $app.AutomationSecurity = 3 } catch {}
                $document = $app.Workbooks.Open($File.FullName, 0, $true)
                $document.PrintOut()
                $document.Close($false)
            }
            '.xlsx' {
                $app = New-Object -ComObject Excel.Application
                $app.Visible = $false
                $app.DisplayAlerts = $false
                try { $app.AutomationSecurity = 3 } catch {}
                $document = $app.Workbooks.Open($File.FullName, 0, $true)
                $document.PrintOut()
                $document.Close($false)
            }
            '.ppt' {
                $app = New-Object -ComObject PowerPoint.Application
                $document = $app.Presentations.Open($File.FullName, $true, $false, $false)
                $document.PrintOptions.PrintInBackground = 0
                $document.PrintOut()
                $document.Close()
            }
            '.pptx' {
                $app = New-Object -ComObject PowerPoint.Application
                $document = $app.Presentations.Open($File.FullName, $true, $false, $false)
                $document.PrintOptions.PrintInBackground = 0
                $document.PrintOut()
                $document.Close()
            }
        }
        Write-Log "PRINTED  $($File.Name)"
    }
    catch {
        Add-FailedFile $File.Name $_.Exception.Message
    }
    finally {
        if ($null -ne $document) { Release-ComObject $document }
        if ($null -ne $app) {
            try { $app.Quit() } catch {}
            Release-ComObject $app
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Print-PdfFile {
    param([System.IO.FileInfo]$File)
    try {
        Start-Process -FilePath $File.FullName -Verb Print | Out-Null
        Write-Log "SUBMITTED $($File.Name) to the default printer"
    }
    catch {
        Add-FailedFile $File.Name $_.Exception.Message
    }
}

function Invoke-BulkPrintFolder {
    param([string]$Path)

    $folder = Get-Item -LiteralPath $Path -ErrorAction SilentlyContinue
    if ($null -eq $folder -or -not $folder.PSIsContainer) {
        Write-Host "Not a folder, skipped: $Path" -ForegroundColor Red
        [void]$script:FailedForPrompt.Add("$Path -- not a folder")
        return
    }

    $script:CurrentFolderName = $folder.Name
    Write-Host "`n=== $($folder.FullName) ===" -ForegroundColor Cyan

    # List files before the log is created, and leave out logs from earlier runs.
    $files = @(Get-ChildItem -LiteralPath $folder.FullName -File |
        Where-Object { $_.Name -notlike 'Bulk-Print-*.log' } |
        Sort-Object Name)
    $printable = @($files | Where-Object { $supportedExtensions -contains $_.Extension.ToLowerInvariant() })
    $zips = @($files | Where-Object { $_.Extension.ToLowerInvariant() -eq '.zip' })
    $other = @($files | Where-Object {
        ($supportedExtensions -notcontains $_.Extension.ToLowerInvariant()) -and ($_.Extension.ToLowerInvariant() -ne '.zip')
    })

    $script:LogPath = Join-Path $folder.FullName "Bulk-Print-$timestamp.log"
    New-Item -ItemType File -Path $script:LogPath -Force | Out-Null
    [void]$script:LogPaths.Add($script:LogPath)
    Write-Log "START    Folder: $($folder.FullName)"
    Write-Log 'PRINTER  Windows default printer'

    if ($zips.Count -gt 0) {
        Write-Host "Found $($zips.Count) ZIP file(s) in '$($folder.Name)'. Extract and print supported contents temporarily? (y/N)" -ForegroundColor Yellow
        $includeZipContents = (Read-Host).Trim().ToLowerInvariant() -eq 'y'
        if ($includeZipContents) {
            foreach ($zip in $zips) {
                $temp = Join-Path ([System.IO.Path]::GetTempPath()) ("BulkPrint-" + [guid]::NewGuid().ToString('N'))
                # Deleted only after the final prompt: PDF readers open the file after Start-Process returns.
                [void]$script:TempFolders.Add($temp)
                try {
                    Expand-Archive -LiteralPath $zip.FullName -DestinationPath $temp -Force
                    $archiveFiles = @(Get-ChildItem -LiteralPath $temp -File -Recurse | Where-Object {
                        $supportedExtensions -contains $_.Extension.ToLowerInvariant()
                    } | Sort-Object FullName)
                    foreach ($file in $archiveFiles) {
                        if ($file.Extension.ToLowerInvariant() -eq '.pdf') { Print-PdfFile $file } else { Print-OfficeFile $file }
                    }
                    $archiveOther = @(Get-ChildItem -LiteralPath $temp -File -Recurse | Where-Object {
                        ($supportedExtensions -notcontains $_.Extension.ToLowerInvariant()) -and ($_.Extension.ToLowerInvariant() -ne '.zip')
                    })
                    foreach ($file in $archiveOther) {
                        Add-SkippedFile "$($zip.Name)\$($file.Name)" 'unsupported file type inside ZIP' $file.Extension
                    }
                    Write-Log "ZIP DONE $($zip.Name)"
                }
                catch {
                    Add-FailedFile $zip.Name "ZIP could not be processed: $($_.Exception.Message)"
                }
            }
        }
        else {
            foreach ($zip in $zips) {
                Add-SkippedFile $zip.Name 'ZIP not selected for extraction' $zip.Extension
            }
        }
    }

    foreach ($file in $printable) {
        if ($file.Extension.ToLowerInvariant() -eq '.pdf') { Print-PdfFile $file } else { Print-OfficeFile $file }
    }

    foreach ($file in $other) {
        Add-SkippedFile $file.Name 'unsupported file type' $file.Extension
    }

    Write-Log 'END      Folder finished.'
}

$folders = @($FolderPath | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
if ($folders.Count -eq 0) {
    Write-Host 'No folder was supplied.' -ForegroundColor Yellow
    Write-Host 'Use this through Send to -> Bulk Print, or pass one or more folder paths.'
    Read-Host 'Press Enter to close'
    exit 1
}

foreach ($path in $folders) {
    # One folder failing (e.g. read-only, so no log) must not stop the others.
    try {
        Invoke-BulkPrintFolder $path
    }
    catch {
        Write-Host "Folder failed: $path -- $($_.Exception.Message)" -ForegroundColor Red
        [void]$script:FailedForPrompt.Add("$path -- folder failed: $($_.Exception.Message)")
    }
}

if ($script:FailedForPrompt.Count -gt 0) {
    Write-Host "`nFAILED ($($script:FailedForPrompt.Count)):" -ForegroundColor Red
    foreach ($failed in $script:FailedForPrompt) {
        Write-Host " - $failed" -ForegroundColor Red
    }
}

if ($script:SkippedForPrompt.Count -gt 0) {
    Write-Host "`nSkipped files (excluding .eml and .md):" -ForegroundColor Yellow
    foreach ($skipped in $script:SkippedForPrompt) {
        Write-Host " - $skipped" -ForegroundColor Yellow
    }
}
else {
    Write-Host "`nNo files were skipped, excluding .eml and .md." -ForegroundColor Green
}

if ($script:LogPaths.Count -gt 0) {
    Write-Host "`nFinished. Log(s) saved to:" -ForegroundColor Green
    foreach ($log in $script:LogPaths) { Write-Host " $log" -ForegroundColor Green }
}

if ($script:TempFolders.Count -gt 0) {
    Write-Host "`nExtracted ZIP files are deleted when you press Enter. Wait until the PDFs have printed first." -ForegroundColor Yellow
}
Read-Host 'Press Enter to close'

foreach ($temp in $script:TempFolders) {
    if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Recurse -Force -ErrorAction SilentlyContinue }
}
