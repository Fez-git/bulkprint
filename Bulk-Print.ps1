param(
    # Send To passes every selected folder as a separate argument.
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$FolderPath
)

$ErrorActionPreference = 'Stop'
$supportedExtensions = @('.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.pdf', '.eml')
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$script:SkippedForPrompt = New-Object 'System.Collections.Generic.List[string]'
$script:FailedForPrompt = New-Object 'System.Collections.Generic.List[string]'
$script:TempFolders = New-Object 'System.Collections.Generic.List[string]'
$script:LogPaths = New-Object 'System.Collections.Generic.List[string]'
$script:NoLogFolders = New-Object 'System.Collections.Generic.List[string]'
$script:LogPath = $null
$script:CurrentFolderName = ''
$script:Outlook = $null

function Write-Log {
    param([string]$Message)
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $Message"
    # No log path when the folder is read-only: screen output only.
    if ($script:LogPath) {
        # The log was once found locked by another process mid-run (cause not identified).
        # Retry briefly, and never let a log write stop the printing.
        for ($attempt = 1; $attempt -le 5; $attempt++) {
            try {
                Add-Content -LiteralPath $script:LogPath -Value $line -ErrorAction Stop
                break
            }
            catch {
                if ($attempt -lt 5) { Start-Sleep -Milliseconds 200 }
                else { Write-Host "WARNING: could not write to the log: $($_.Exception.Message)" -ForegroundColor Yellow }
            }
        }
    }
    Write-Host $line
}

function Add-SkippedFile {
    param(
        [string]$Name,
        [string]$Reason,
        [string]$Extension
    )

    Write-Log "SKIPPED  $Name -- $Reason"
    if ($Extension.ToLowerInvariant() -notin @('.md')) {
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

function Test-EncryptedOpenXml {
    param([System.IO.FileInfo]$File)

    # A normal .docx/.xlsx/.pptx is a ZIP and starts with "PK". A password-protected
    # one is wrapped in an OLE container instead and starts with D0 CF 11 E0.
    if ($File.Extension.ToLowerInvariant() -notin @('.docx', '.xlsx', '.pptx')) { return $false }
    $stream = [System.IO.File]::OpenRead($File.FullName)
    try {
        $header = New-Object byte[] 4
        $read = $stream.Read($header, 0, 4)
        return ($read -eq 4 -and $header[0] -eq 0xD0 -and $header[1] -eq 0xCF -and $header[2] -eq 0x11 -and $header[3] -eq 0xE0)
    }
    finally {
        $stream.Dispose()
    }
}

function Print-OfficeFile {
    param([System.IO.FileInfo]$File)

    $extension = $File.Extension.ToLowerInvariant()
    $app = $null
    $document = $null
    $keepAppOpen = $false
    # A wrong password makes Office raise an error instead of showing a password box
    # in a hidden window (which hangs the script). Unprotected files ignore it.
    $dummyPassword = 'BulkPrint-NoPassword'

    try {
        if (Test-EncryptedOpenXml $File) {
            Add-SkippedFile $File.Name 'password-protected, print it by hand' $File.Extension
            return
        }

        switch ($extension) {
            { $_ -in '.doc', '.docx' } {
                $app = New-Object -ComObject Word.Application
                $app.Visible = $false
                $app.DisplayAlerts = 0
                try { $app.AutomationSecurity = 3 } catch {}
                # FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument,
                # PasswordTemplate, Revert, WritePasswordDocument
                $document = $app.Documents.Open($File.FullName, $false, $true, $false, $dummyPassword, $dummyPassword, $false, $dummyPassword)
                # Background = $false: return only once spooled, so Quit cannot cancel the job.
                $document.PrintOut($false)
                $document.Close(0)
            }
            { $_ -in '.xls', '.xlsx' } {
                $app = New-Object -ComObject Excel.Application
                $app.Visible = $false
                $app.DisplayAlerts = $false
                try { $app.AutomationSecurity = 3 } catch {}
                # FileName, UpdateLinks, ReadOnly, Format, Password, WriteResPassword, IgnoreReadOnlyRecommended
                # WriteResPassword stays Missing: a dummy value there makes every unprotected workbook fail to open.
                $document = $app.Workbooks.Open($File.FullName, 0, $true, 5, $dummyPassword, [Type]::Missing, $true)
                $document.PrintOut()
                $document.Close($false)
            }
            { $_ -in '.ppt', '.pptx' } {
                # PowerPoint runs only once per user: New-Object attaches to an open copy,
                # and Quit would close the user's own presentations.
                $keepAppOpen = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue).Count -gt 0
                $app = New-Object -ComObject PowerPoint.Application
                # PowerPoint has no password argument; "file::password::" fails instead of prompting.
                # .pptx needs no trick: an encrypted one was already caught above.
                $openPath = if ($extension -eq '.ppt') { "$($File.FullName)::$($dummyPassword)::" } else { $File.FullName }
                $document = $app.Presentations.Open($openPath, $true, $false, $false)
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
        if ($null -ne $document) {
            # Inside the user's own PowerPoint, close a presentation left open by a failed print.
            if ($keepAppOpen) { try { $document.Close() } catch {} }
            Release-ComObject $document
        }
        if ($null -ne $app) {
            if (-not $keepAppOpen) { try { $app.Quit() } catch {} }
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

function ConvertFrom-EncodedWord {
    param([string]$Text)

    # RFC 2047 encoded words, e.g. =?UTF-8?B?...?= or =?iso-8859-1?Q?...?=
    # Whitespace between two encoded words is not part of the text.
    $Text = $Text -replace '(\?=)\s+(=\?)', '$1$2'
    return [regex]::Replace($Text, '=\?([^?]+)\?([BbQq])\?([^?]*)\?=', [System.Text.RegularExpressions.MatchEvaluator]{
        param($m)
        try {
            $encoding = [System.Text.Encoding]::GetEncoding($m.Groups[1].Value)
            if ($m.Groups[2].Value -in 'B', 'b') {
                $bytes = [Convert]::FromBase64String($m.Groups[3].Value)
            }
            else {
                $q = $m.Groups[3].Value -replace '_', ' '
                $list = New-Object 'System.Collections.Generic.List[byte]'
                for ($i = 0; $i -lt $q.Length; $i++) {
                    if ($q[$i] -eq '=' -and $i + 2 -lt $q.Length) {
                        $list.Add([Convert]::ToByte($q.Substring($i + 1, 2), 16)); $i += 2
                    }
                    else { $list.Add([byte][char]$q[$i]) }
                }
                $bytes = $list.ToArray()
            }
            $encoding.GetString($bytes)
        }
        catch { $m.Value }
    })
}

function Get-EmlSubject {
    param([System.IO.FileInfo]$File)

    # Headers sit at the top of the file; 64 KB is plenty and avoids reading attachments.
    $stream = [System.IO.File]::OpenRead($File.FullName)
    try {
        $buffer = New-Object byte[] ([Math]::Min($stream.Length, 65536))
        $read = $stream.Read($buffer, 0, $buffer.Length)
    }
    finally {
        $stream.Dispose()
    }
    $text = [System.Text.Encoding]::UTF8.GetString($buffer, 0, $read)
    $headerEnd = [regex]::Match($text, '\r?\n\r?\n')
    if ($headerEnd.Success) { $text = $text.Substring(0, $headerEnd.Index) }
    # Long headers are folded onto continuation lines that start with whitespace.
    $text = $text -replace '\r?\n[ \t]+', ' '
    $match = [regex]::Match($text, '(?im)^Subject:[ \t]*(.*?)\r?$')
    if (-not $match.Success) { return '' }
    return (ConvertFrom-EncodedWord $match.Groups[1].Value).Trim()
}

function Get-Outlook {
    # New-Object, not GetActiveObject: a freshly started Outlook stays invisible to GetActiveObject
    # for minutes (Office registers itself late), and Outlook runs once per user, so New-Object
    # returns the running copy or starts one.
    # An Outlook with no main window exits when its last email window closes, even while this
    # script holds it, so check the connection is alive and reconnect if not.
    # The check must be a method call: PowerShell returns $null instead of throwing when a
    # property is read from a disconnected COM object.
    if ($null -ne $script:Outlook) {
        try { [void]$script:Outlook.GetNamespace('MAPI') }
        catch {
            Release-ComObject $script:Outlook
            $script:Outlook = $null
        }
    }
    if ($null -eq $script:Outlook) {
        $script:Outlook = New-Object -ComObject Outlook.Application
    }
    return $script:Outlook
}

function Print-EmlFile {
    param([System.IO.FileInfo]$File)

    $inspector = $null
    $item = $null

    try {
        $subject = Get-EmlSubject $File
        $outlook = Get-Outlook
        $initialCount = $outlook.Inspectors.Count

        # Outlook's object model cannot open .eml (OpenSharedItem takes .msg only),
        # so open it the way a double-click does and pick up the window Outlook shows.
        Start-Process -FilePath $File.FullName | Out-Null

        $timeoutSeconds = 60
        $deadline = (Get-Date).AddSeconds($timeoutSeconds)
        while ($null -eq $inspector -and (Get-Date) -lt $deadline) {
            Start-Sleep -Milliseconds 500
            try {
                for ($i = $outlook.Inspectors.Count; $i -gt $initialCount; $i--) {
                    $candidate = $outlook.Inspectors.Item($i)
                    # Match on Subject so an email the user opens meanwhile is never printed by mistake.
                    if ("$($candidate.CurrentItem.Subject)".Trim() -eq $subject) {
                        $inspector = $candidate
                        break
                    }
                    Release-ComObject $candidate
                }
            }
            catch {}
        }
        if ($null -eq $inspector) {
            throw "Outlook did not open the email within $timeoutSeconds seconds (looked for subject '$subject')"
        }

        try { $inspector.WindowState = 1 } catch {}   # olMinimized
        $item = $inspector.CurrentItem
        # Same as Outlook's Quick Print: default printer, Memo Style, whole email.
        $item.PrintOut()
        Start-Sleep -Seconds 2
        Write-Log "PRINTED  $($File.Name)"
    }
    catch {
        Add-FailedFile $File.Name $_.Exception.Message
    }
    finally {
        # Close only the email window (olDiscard = 1). Never Quit: Outlook is the user's mailbox.
        if ($null -ne $inspector) {
            try { $inspector.Close(1) } catch {}
        }
        Release-ComObject $item
        Release-ComObject $inspector
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Print-File {
    param([System.IO.FileInfo]$File)

    switch ($File.Extension.ToLowerInvariant()) {
        '.pdf' { Print-PdfFile $File }
        '.eml' { Print-EmlFile $File }
        default { Print-OfficeFile $File }
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
    try {
        New-Item -ItemType File -Path $script:LogPath -Force | Out-Null
        [void]$script:LogPaths.Add($script:LogPath)
    }
    catch {
        # Read-only folder (e.g. a network share): print anyway, without a log.
        $script:LogPath = $null
        [void]$script:NoLogFolders.Add($folder.FullName)
        Write-Host 'WARNING: this folder is read-only, so no log was created. Printing anyway.' -ForegroundColor Yellow
    }
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
                        Print-File $file
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
        Print-File $file
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

# Let go of Outlook without closing it: an Outlook the user had open stays open,
# one started for this run exits by itself once nothing holds it.
Release-ComObject $script:Outlook
$script:Outlook = $null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

if ($script:FailedForPrompt.Count -gt 0) {
    Write-Host "`nFAILED ($($script:FailedForPrompt.Count)):" -ForegroundColor Red
    foreach ($failed in $script:FailedForPrompt) {
        Write-Host " - $failed" -ForegroundColor Red
    }
}

if ($script:SkippedForPrompt.Count -gt 0) {
    Write-Host "`nSkipped files (excluding .md):" -ForegroundColor Yellow
    foreach ($skipped in $script:SkippedForPrompt) {
        Write-Host " - $skipped" -ForegroundColor Yellow
    }
}
else {
    Write-Host "`nNo files were skipped, excluding .md." -ForegroundColor Green
}

if ($script:LogPaths.Count -gt 0) {
    Write-Host "`nFinished. Log(s) saved to:" -ForegroundColor Green
    foreach ($log in $script:LogPaths) { Write-Host " $log" -ForegroundColor Green }
}

if ($script:NoLogFolders.Count -gt 0) {
    Write-Host "`nWARNING: no log was created for these read-only folder(s). Their files were still sent to print:" -ForegroundColor Yellow
    foreach ($noLog in $script:NoLogFolders) { Write-Host " $noLog" -ForegroundColor Yellow }
}

if ($script:TempFolders.Count -gt 0) {
    Write-Host "`nExtracted ZIP files are deleted when you press Enter. Wait until the PDFs have printed first." -ForegroundColor Yellow
}
Read-Host 'Press Enter to close'

foreach ($temp in $script:TempFolders) {
    if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Recurse -Force -ErrorAction SilentlyContinue }
}
