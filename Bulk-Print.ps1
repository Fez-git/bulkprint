param(
    [Parameter(Position = 0)]
    [string]$FolderPath
)

$ErrorActionPreference = 'Stop'
$supportedExtensions = @('.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.pdf')
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$script:SkippedForPrompt = New-Object 'System.Collections.Generic.List[string]'

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
        [void]$script:SkippedForPrompt.Add("$Name -- $Reason")
    }
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
                $document.PrintOut()
                $document.Close(0)
            }
            '.docx' {
                $app = New-Object -ComObject Word.Application
                $app.Visible = $false
                try { $app.AutomationSecurity = 3 } catch {}
                $document = $app.Documents.Open($File.FullName, $false, $true)
                $document.PrintOut()
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
                $document.PrintOut()
                $document.Close()
            }
            '.pptx' {
                $app = New-Object -ComObject PowerPoint.Application
                $document = $app.Presentations.Open($File.FullName, $true, $false, $false)
                $document.PrintOut()
                $document.Close()
            }
        }
        Write-Log "PRINTED  $($File.Name)"
    }
    catch {
        Write-Log "FAILED   $($File.Name) -- $($_.Exception.Message)"
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
        Write-Log "FAILED    $($File.Name) -- $($_.Exception.Message)"
    }
}

if ([string]::IsNullOrWhiteSpace($FolderPath)) {
    Write-Host 'No folder was supplied.' -ForegroundColor Yellow
    Write-Host 'Use this through Send to -> Bulk Print, or pass a folder path.'
    Read-Host 'Press Enter to close'
    exit 1
}

$folder = Get-Item -LiteralPath $FolderPath -ErrorAction SilentlyContinue
if ($null -eq $folder -or -not $folder.PSIsContainer) {
    Write-Host "Folder not found: $FolderPath" -ForegroundColor Red
    Read-Host 'Press Enter to close'
    exit 1
}

$script:LogPath = Join-Path $folder.FullName "Bulk-Print-$timestamp.log"
New-Item -ItemType File -Path $script:LogPath -Force | Out-Null
Write-Log "START    Folder: $($folder.FullName)"
Write-Log 'PRINTER  Windows default printer'

$files = @(Get-ChildItem -LiteralPath $folder.FullName -File | Sort-Object Name)
$printable = @($files | Where-Object { $supportedExtensions -contains $_.Extension.ToLowerInvariant() })
$zips = @($files | Where-Object { $_.Extension.ToLowerInvariant() -eq '.zip' })
$other = @($files | Where-Object {
    ($supportedExtensions -notcontains $_.Extension.ToLowerInvariant()) -and ($_.Extension.ToLowerInvariant() -ne '.zip')
})

if ($zips.Count -gt 0) {
    Write-Host "Found $($zips.Count) ZIP file(s). Extract and print supported contents temporarily? (y/N)" -ForegroundColor Yellow
    $includeZipContents = (Read-Host).Trim().ToLowerInvariant() -eq 'y'
    if ($includeZipContents) {
        foreach ($zip in $zips) {
            $temp = Join-Path ([System.IO.Path]::GetTempPath()) ("BulkPrint-" + [guid]::NewGuid().ToString('N'))
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
                    Add-SkippedFile $file.FullName 'unsupported file type inside ZIP' $file.Extension
                }
                Write-Log "ZIP DONE $($zip.Name)"
            }
            catch {
                Write-Log "ZIP FAIL $($zip.Name) -- $($_.Exception.Message)"
            }
            finally {
                if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Recurse -Force -ErrorAction SilentlyContinue }
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

Write-Log 'END      Print job submitted.'
if ($script:SkippedForPrompt.Count -gt 0) {
    Write-Host "`nSkipped files (excluding .eml and .md):" -ForegroundColor Yellow
    foreach ($skipped in $script:SkippedForPrompt) {
        Write-Host " - $skipped" -ForegroundColor Yellow
    }
}
else {
    Write-Host "`nNo files were skipped, excluding .eml and .md." -ForegroundColor Green
}
Write-Host "`nFinished. Log saved to:`n$script:LogPath" -ForegroundColor Green
Read-Host 'Press Enter to close'
