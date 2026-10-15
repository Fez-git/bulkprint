param(
    [Parameter(Position = 0)]
    [string]$EmlPath
)

$ErrorActionPreference = 'Stop'
$outlook = $null
$inspector = $null

try {
    if ([string]::IsNullOrWhiteSpace($EmlPath) -or -not (Test-Path -LiteralPath $EmlPath -PathType Leaf)) {
        throw "The selected .eml file was not found: $EmlPath"
    }

    if ([System.IO.Path]::GetExtension($EmlPath).ToLowerInvariant() -ne '.eml') {
        throw 'This shortcut accepts .eml files only.'
    }

    $resolvedPath = (Resolve-Path -LiteralPath $EmlPath).Path
    $initialInspectorCount = 0

    try {
        $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
        $initialInspectorCount = $outlook.Inspectors.Count
    }
    catch {}

    # Open through the normal Windows association so Outlook renders the real email header.
    Start-Process -FilePath $resolvedPath | Out-Null

    for ($attempt = 1; $attempt -le 40 -and $null -eq $inspector; $attempt++) {
        Start-Sleep -Milliseconds 500
        try {
            if ($null -eq $outlook) {
                $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
            }
            if ($outlook.Inspectors.Count -gt $initialInspectorCount) {
                $inspector = $outlook.Inspectors.Item($outlook.Inspectors.Count)
            }
        }
        catch {}
    }

    if ($null -eq $inspector) { throw 'Outlook did not open the selected .eml file.' }

    $inspector.Activate()
    $inspector.CommandBars.ExecuteMso('FilePrint')
    Write-Host 'Outlook print dialog opened. Select pages 1-2, then click Print.' -ForegroundColor Green
}
catch {
    Write-Host "Email printing failed: $($_.Exception.Message)" -ForegroundColor Red
    Read-Host 'Press Enter to close'
    exit 1
}
finally {
    if ($null -ne $inspector) {
        try { $inspector.Close() } catch {}
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($inspector) } catch {}
    }
    if ($null -ne $outlook) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Start-Sleep -Seconds 1
