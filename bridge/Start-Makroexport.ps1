#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
$picker = New-Object Windows.Forms.OpenFileDialog
$picker.Title = 'Die aktuelle Kassen-MDB auswaehlen'
$picker.Filter = 'Access-Datenbank (*.mdb;*.accdb)|*.mdb;*.accdb'
$picker.FileName = 'Kassa_CSE_V2.0.mdb'
if ($picker.ShowDialog() -ne [Windows.Forms.DialogResult]::OK) { Write-Host 'Abgebrochen.'; exit 0 }
$source = $picker.FileName
$picker.Dispose()
$original = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
$resultRoot = Join-Path $PSScriptRoot 'Ergebnisse'
New-Item -ItemType Directory -Path $resultRoot -Force | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$nonce = [guid]::NewGuid().ToString('N').Substring(0, 8)
$output = Join-Path $resultRoot ('Makros-{0}-{1}' -f $stamp, $nonce)
$failure = $null
try {
    & (Join-Path $PSScriptRoot 'Export-LinAccessBridge.ps1') -DatabasePath $source `
        -OutputDirectory $output -AccessProgId 'Access.Application.11' `
        -RequiredAccessMajorVersion 11 -MacrosOnly -IncludeQueryMetadata
} catch {
    $exception = $_.Exception
    while ($null -ne $exception.InnerException) { $exception = $exception.InnerException }
    $failure = [ordered]@{ hresult = [int]$exception.HResult; exception_type = $exception.GetType().FullName }
    Write-Host ('Exportfehler: HRESULT ' + $failure.hresult)
}
New-Item -ItemType Directory -Path $output -Force | Out-Null
$unchanged = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -eq $original
$manifestPath = Join-Path $output 'manifest.json'
$macros = 0
$exported = 0
$objectFailures = 0
$metadataFailures = 0
$metadataQueries = 0
$metadataSkipped = 0
if (Test-Path -LiteralPath $manifestPath) {
    $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
    $macros = @($manifest.objects | Where-Object { $_.kind -eq 'macro' }).Count
    $exported = @($manifest.objects | Where-Object { $_.kind -eq 'macro' -and $_.status -eq 'source_only' }).Count
    $objectFailures = @($manifest.objects | Where-Object { $_.status -eq 'failed' -or $_.status -eq 'unavailable' }).Count
    foreach ($entry in $manifest.objects) {
        if ($entry.kind -ne 'query_metadata' -or [string]::IsNullOrEmpty($entry.file)) { continue }
        $metadata = Get-Content -LiteralPath (Join-Path $output $entry.file) -Raw | ConvertFrom-Json
        $metadataQueries += @($metadata.queries).Count
        $metadataFailures += @($metadata.queries | Where-Object { $_.status -eq 'failed' -or $_.status -eq 'partial_metadata' }).Count
        $metadataSkipped += @($metadata.queries | Where-Object { $_.status -eq 'not_inspected_nonlocal_or_action' }).Count
    }

}
$status = if (-not $unchanged) { 'original_changed' } elseif ($null -ne $failure) { 'failed' } elseif ($objectFailures -gt 0 -or $metadataFailures -gt 0) { 'completed_with_observations' } else { 'completed' }
$result = [ordered]@{
    format = 'linaccess-macro-reference-v1'
    status = $status
    original_unveraendert = $unchanged
    source_sha256 = $original.ToLowerInvariant()
    macros_found = $macros
    macro_sources_exported = $exported
    object_failures = $objectFailures
    query_metadata_entries = $metadataQueries
    query_metadata_failures = $metadataFailures
    query_metadata_skipped = $metadataSkipped
    failure = $failure
}
[IO.File]::WriteAllText((Join-Path $output 'run-result.json'), (ConvertTo-Json -InputObject $result -Depth 8),
    (New-Object Text.UTF8Encoding($false)))
$zip = Join-Path $resultRoot ('LinAccess-Makros-{0}-{1}.zip' -f $stamp, $nonce)
Compress-Archive -Path (Join-Path $output '*') -DestinationPath $zip
Write-Host ''
Write-Host ('Ergebnis-ZIP: ' + $zip)
Write-Host ('Status: ' + $status + '; Original unveraendert: ' + $unchanged)
Write-Host ('Makros gefunden/exportiert: ' + $macros + '/' + $exported)
Write-Host 'Bitte nur dieses Ergebnis-ZIP im Chat hochladen.'
Write-Host 'Es enthaelt Makrodefinitionen und Abfragemetadaten, keine MDB und keine VBA-Quellen.'
