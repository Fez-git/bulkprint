param(
    # Send To passes every selected folder as a separate argument.
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$FolderPath
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
# Images are not printed one by one: each folder's images go into one job, 4 per page.
$imageExtensions = @('.jpg', '.jpeg', '.png')
# Page orientation for image pages. Portrait suits portrait photos; set to $true for landscape.
$imagePageLandscape = $false
$supportedExtensions = @('.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.pdf', '.eml') + $imageExtensions
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

function Get-PrintImage {
    param(
        [string]$Path,
        [double]$BoxWidth,
        [double]$BoxHeight
    )

    # Read via memory so the file is not locked, then copy to a new Bitmap so the stream can close.
    $stream = New-Object System.IO.MemoryStream(, [System.IO.File]::ReadAllBytes($Path))
    try {
        $source = [System.Drawing.Image]::FromStream($stream)
        try {
            # Phone photos are saved unrotated plus an EXIF orientation tag (0x0112) that GDI+ ignores.
            if ($source.PropertyIdList -contains 0x0112) {
                $rotate = switch ([BitConverter]::ToUInt16($source.GetPropertyItem(0x0112).Value, 0)) {
                    2 { 'RotateNoneFlipX' }
                    3 { 'Rotate180FlipNone' }
                    4 { 'Rotate180FlipX' }
                    5 { 'Rotate90FlipX' }
                    6 { 'Rotate90FlipNone' }
                    7 { 'Rotate270FlipX' }
                    8 { 'Rotate270FlipNone' }
                    default { $null }
                }
                if ($rotate) { $source.RotateFlip([System.Drawing.RotateFlipType]$rotate) }
            }

            # Shrink to the cell size at 300 dpi (box is in 1/100 inch): full-size photos make a huge print job.
            # 1.0, not 1: with an int first argument PowerShell picks Math.Min(int, int) and rounds the scale to 0 or 1.
            $scale = [Math]::Min(1.0, [Math]::Min(($BoxWidth * 3) / $source.Width, ($BoxHeight * 3) / $source.Height))
            $width = [Math]::Max(1, [int]($source.Width * $scale))
            $height = [Math]::Max(1, [int]($source.Height * $scale))
            $bitmap = New-Object System.Drawing.Bitmap($width, $height)
            $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
            try {
                $graphics.Clear([System.Drawing.Color]::White)   # transparent PNG areas print white
                $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
                $graphics.DrawImage($source, 0, 0, $width, $height)
            }
            finally {
                $graphics.Dispose()
            }
            return $bitmap
        }
        finally {
            $source.Dispose()
        }
    }
    finally {
        $stream.Dispose()
    }
}

function Print-ImageFiles {
    param(
        [System.IO.FileInfo[]]$Files,
        [string]$JobName,
        # Tests pass a PreviewPrintController to render pages without printing.
        [System.Drawing.Printing.PrintController]$Controller = $null
    )

    $document = New-Object System.Drawing.Printing.PrintDocument
    $script:ImageJob = @{
        Files   = $Files
        Index   = 0
        Pages   = 0
        Printed = 0
        Font    = New-Object System.Drawing.Font('Segoe UI', 8)
        Format  = New-Object System.Drawing.StringFormat
    }
    $script:ImageJob.Format.Alignment = [System.Drawing.StringAlignment]::Center
    $script:ImageJob.Format.Trimming = [System.Drawing.StringTrimming]::EllipsisCharacter
    $script:ImageJob.Format.FormatFlags = [System.Drawing.StringFormatFlags]::NoWrap

    try {
        $document.DocumentName = "Bulk Print - images - $JobName"
        $document.DefaultPageSettings.Landscape = $imagePageLandscape
        $document.DefaultPageSettings.Color = $true
        $document.DefaultPageSettings.Margins = New-Object System.Drawing.Printing.Margins(40, 40, 40, 40)
        if ($null -ne $Controller) { $document.PrintController = $Controller }

        # Logged every time: a driver can ignore the colour request (manufacturer drivers often have their
        # own colour-mode setting), and this line is what tells the cases apart afterwards.
        $printerName = $document.PrinterSettings.PrinterName
        $driverName = try { (Get-CimInstance Win32_Printer | Where-Object Name -eq $printerName | Select-Object -First 1).DriverName } catch { 'unknown' }
        Write-Log "IMAGES   printer '$printerName', driver '$driverName', reports colour support: $($document.PrinterSettings.SupportsColor), colour requested: $($document.DefaultPageSettings.Color)"
        if (-not $document.PrinterSettings.SupportsColor) {
            Write-Log 'WARNING  the default printer does not report colour support; images may print in black and white'
        }

        $document.add_PrintPage({
            param($sender, $e)
            $job = $script:ImageJob
            $job.Pages++
            $g = $e.Graphics
            # On a real printer the drawing origin is the corner of the printable area, not the paper edge.
            if (-not $sender.PrintController.IsPreview) {
                $g.TranslateTransform(-$e.PageSettings.HardMarginX, -$e.PageSettings.HardMarginY)
            }

            # 2 x 2 grid inside the margins; units are 1/100 inch.
            $m = $e.MarginBounds
            $gap = 20
            $captionHeight = 18
            $cellWidth = ($m.Width - $gap) / 2
            $cellHeight = ($m.Height - $gap) / 2
            $boxHeight = $cellHeight - $captionHeight

            for ($slot = 0; $slot -lt 4 -and $job.Index -lt $job.Files.Count; $slot++) {
                $file = $job.Files[$job.Index]
                $job.Index++
                $x = $m.Left + ($slot % 2) * ($cellWidth + $gap)
                $y = $m.Top + [Math]::Floor($slot / 2) * ($cellHeight + $gap)
                try {
                    $bitmap = Get-PrintImage $file.FullName $cellWidth $boxHeight
                    try {
                        $fit = [Math]::Min($cellWidth / $bitmap.Width, $boxHeight / $bitmap.Height)
                        $w = $bitmap.Width * $fit
                        $h = $bitmap.Height * $fit
                        $g.DrawImage($bitmap, [single]($x + ($cellWidth - $w) / 2), [single]($y + ($boxHeight - $h) / 2), [single]$w, [single]$h)
                    }
                    finally {
                        $bitmap.Dispose()
                    }
                    $job.Printed++
                }
                catch {
                    $g.DrawString('(image could not be read)', $job.Font, [System.Drawing.Brushes]::Gray,
                        (New-Object System.Drawing.RectangleF([single]$x, [single]($y + $boxHeight / 2), [single]$cellWidth, [single]$captionHeight)), $job.Format)
                    Add-FailedFile $file.Name "image could not be read: $($_.Exception.Message)"
                }
                $g.DrawString($file.Name, $job.Font, [System.Drawing.Brushes]::Black,
                    (New-Object System.Drawing.RectangleF([single]$x, [single]($y + $boxHeight + 2), [single]$cellWidth, [single]$captionHeight)), $job.Format)
            }
            $e.HasMorePages = $job.Index -lt $job.Files.Count
        })

        $document.Print()
        Write-Log "PRINTED  $($script:ImageJob.Printed) of $($Files.Count) image(s) on $($script:ImageJob.Pages) page(s), colour, 4 per page"
    }
    catch {
        Add-FailedFile "$($Files.Count) image(s)" "image print job failed: $($_.Exception.Message)"
    }
    finally {
        $script:ImageJob.Font.Dispose()
        $script:ImageJob.Format.Dispose()
        $document.Dispose()
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

    # Images (also from ZIPs) are collected and printed together at the end of the folder.
    $folderImages = New-Object 'System.Collections.Generic.List[System.IO.FileInfo]'

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
                        if ($imageExtensions -contains $file.Extension.ToLowerInvariant()) { $folderImages.Add($file) }
                        else { Print-File $file }
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
        if ($imageExtensions -contains $file.Extension.ToLowerInvariant()) { $folderImages.Add($file) }
        else { Print-File $file }
    }

    if ($folderImages.Count -gt 0) {
        Print-ImageFiles $folderImages.ToArray() $folder.Name
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
