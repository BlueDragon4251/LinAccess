#Requires -Version 5.1
<#
  Optional Windows/Access metadata exporter for LinAccess.
  Runs only on Windows with Microsoft Access installed. Original database is
  copied to a fresh temporary folder before COM opens it; source is never edited.
  Access.SaveAsText is an undocumented/unsupported compatibility mechanism on
  some versions; every failure is written to the manifest.

  Usage:
  powershell.exe -NoProfile -File .\Export-LinAccessBridge.ps1 `
    -DatabasePath "C:\Data\Kunden.accdb" -OutputDirectory "C:\Data\LinAccessExport"

  Explicit Access 2003 selection (only if registered on Windows):
    -AccessProgId Access.Application.11 -RequiredAccessMajorVersion 11
#>
param(
    [Parameter(Mandatory = $true)][string]$DatabasePath,
    [Parameter(Mandatory = $true)][string]$OutputDirectory,
    [ValidatePattern('^Access\.Application(\.\d+)?$')][string]$AccessProgId = 'Access.Application',
    [ValidateRange(0, 99)][int]$RequiredAccessMajorVersion = 0,
    [switch]$MacrosOnly,
    [switch]$IncludeQueryMetadata
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$access = $null
$accessVersion = ''
$database = $null
$scratch = $null
$source = (Resolve-Path -LiteralPath $DatabasePath -ErrorAction Stop).ProviderPath
if (-not (Test-Path -LiteralPath $source -PathType Leaf)) {
    throw 'Access source is not a file.'
}
$suffix = [IO.Path]::GetExtension($source).ToLowerInvariant()
if ($suffix -notin @('.accdb', '.mdb')) {
    throw 'Only .mdb and .accdb sources are accepted.'
}
$destination = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $destination) {
    if ((Get-ChildItem -LiteralPath $destination -Force | Select-Object -First 1)) {
        throw 'Output directory must be empty. Existing files will not be overwritten.'
    }
} else {
    New-Item -ItemType Directory -Path $destination -Force | Out-Null
}
$artifactDir = Join-Path $destination 'objects'
New-Item -ItemType Directory -Path $artifactDir -Force | Out-Null

$script:entries = New-Object 'System.Collections.Generic.List[object]'
$script:counter = 0
$originalHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash

function New-ArtifactPath([string]$kind, [string]$extension) {
    $script:counter++
    $basename = '{0:d6}_{1}{2}' -f $script:counter, $kind, $extension
    return [pscustomobject]@{
        relative = ('objects/' + $basename)
        absolute = (Join-Path $artifactDir $basename)
    }
}

function Add-Entry([string]$kind, [string]$name, [string]$status,
                   [string]$format, [string]$file, [string]$detail) {
    $digest = ''
    if (-not [string]::IsNullOrEmpty($file)) {
        $fullPath = Join-Path $destination ($file.Replace('/', [IO.Path]::DirectorySeparatorChar))
        if (Test-Path -LiteralPath $fullPath -PathType Leaf) {
            $digest = (Get-FileHash -LiteralPath $fullPath -Algorithm SHA256).Hash.ToLowerInvariant()
        }
    }
    $script:entries.Add([pscustomobject]@{
        kind   = $kind
        name   = $name
        status = $status
        format = $format
        file   = $file
        sha256 = $digest
        detail = $detail
    })
}

function Export-SaveAsText([string]$kind, [int]$accessType, [string]$name) {
    $target = New-ArtifactPath $kind '.txt'
    try {
        # Access object type constants: acQuery=1, acForm=2, acReport=3, acMacro=4, acModule=5.
        $script:access.SaveAsText($accessType, $name, $target.absolute)
        if (-not (Test-Path -LiteralPath $target.absolute -PathType Leaf)) {
            throw 'SaveAsText returned without producing a file.'
        }
        Add-Entry $kind $name 'source_only' 'access-saveastext' $target.relative ''
    } catch {
        Add-Entry $kind $name 'failed' 'access-saveastext' '' ('HRESULT ' + [string]$_.Exception.HResult)
    }
}

function Export-Query([string]$name, [string]$sql) {
    $target = New-ArtifactPath 'query' '.sql'
    try {
        [IO.File]::WriteAllText($target.absolute, $sql, (New-Object System.Text.UTF8Encoding($false)))
        Add-Entry 'query' $name 'source_only' 'access-query-sql' $target.relative ''
    } catch {
        Add-Entry 'query' $name 'failed' 'access-query-sql' '' $_.Exception.Message
    }
}


# Read existing metadata only. No recordset, action query or VBA is invoked.
function Get-MetadataFailure($record, [string]$stage) {
    $exception = $record.Exception
    while ($null -ne $exception.InnerException) { $exception = $exception.InnerException }
    return [pscustomobject]@{
        stage = $stage
        hresult = [int]$exception.HResult
        exception_type = $exception.GetType().FullName
    }
}

function Export-QueryMetadata {
    $records = New-Object 'System.Collections.Generic.List[object]'
    $externalTables = $false
    foreach ($table in $script:database.TableDefs) {
        # dbAttachedTable/dbAttachedODBC; never read/export connection strings.
        if (([long]$table.Attributes -band 1610612736) -ne 0) { $externalTables = $true }
    }
    foreach ($query in $script:database.QueryDefs) {
        $failures = New-Object 'System.Collections.Generic.List[object]'
        $fields = New-Object 'System.Collections.Generic.List[object]'
        $parameters = New-Object 'System.Collections.Generic.List[object]'
        $entry = [ordered]@{
            name = [string]$query.Name
            status = 'metadata_only'
            query_type = $null
            sql_sha256 = ''
            fields = @()
            parameters = @()
            failures = @()
        }
        try {
            $entry.query_type = [int]$query.Type
            $sqlBytes = [Text.Encoding]::UTF8.GetBytes([string]$query.SQL)
            $hash = [Security.Cryptography.SHA256]::Create()
            try { $entry.sql_sha256 = ([BitConverter]::ToString($hash.ComputeHash($sqlBytes))).Replace('-', '').ToLowerInvariant() }
            finally { $hash.Dispose() }
            # Action queries and external sources are classified without opening
            # them. Field/parameter inspection remains metadata only.
            if ($externalTables -or $entry.query_type -notin @(0, 16, 128)) {
                $entry.status = 'not_inspected_nonlocal_or_action'
            } else {
                for ($parameterIndex = 0; $parameterIndex -lt [int]$query.Parameters.Count; $parameterIndex++) {
                    $parameter = $query.Parameters.Item($parameterIndex)
                    try {
                        $parameters.Add([pscustomobject]@{
                            name = [string]$parameter.Name
                            dao_type = [int]$parameter.Type
                        })
                    } catch { $failures.Add((Get-MetadataFailure $_ 'parameter')) }
                }
                for ($fieldIndex = 0; $fieldIndex -lt [int]$query.Fields.Count; $fieldIndex++) {
                    $field = $query.Fields.Item($fieldIndex)
                    try {
                        $item = [ordered]@{
                            name = [string]$field.Name
                            dao_type = [int]$field.Type
                            size = [int]$field.Size
                            source_table = $null
                            source_field = $null
                        }
                        try { $item.source_table = [string]$field.SourceTable }
                        catch { $failures.Add((Get-MetadataFailure $_ 'field_source_table')) }
                        try { $item.source_field = [string]$field.SourceField }
                        catch { $failures.Add((Get-MetadataFailure $_ 'field_source_field')) }
                        $fields.Add([pscustomobject]$item)
                    } catch { $failures.Add((Get-MetadataFailure $_ 'field')) }
                }
                if ($failures.Count -gt 0) { $entry.status = 'partial_metadata' }
            }
        } catch {
            $entry.status = 'failed'
            $failures.Add((Get-MetadataFailure $_ 'query'))
        }
        $entry.fields = @($fields.ToArray())
        $entry.parameters = @($parameters.ToArray())
        $entry.failures = @($failures.ToArray())
        $records.Add([pscustomobject]$entry)
    }
    $target = New-ArtifactPath 'query_metadata' '.json'
    $payload = [ordered]@{
        format = 'linaccess-access-query-fields-v1'
        records_executed = $false
        original_vba_requested = $false
        queries = @($records.ToArray())
    }
    [IO.File]::WriteAllText($target.absolute, (ConvertTo-Json -InputObject $payload -Depth 10),
        (New-Object Text.UTF8Encoding($false)))
    Add-Entry 'query_metadata' '(DAO query fields)' 'metadata_only' 'access-dao-field-metadata' $target.relative ''
}

try {
    $scratch = Join-Path ([IO.Path]::GetTempPath()) ('linaccess-bridge-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $scratch | Out-Null
    $workingCopy = Join-Path $scratch ('access-copy' + $suffix)
    Copy-Item -LiteralPath $source -Destination $workingCopy -ErrorAction Stop

    try {
        $access = New-Object -ComObject $AccessProgId -ErrorAction Stop
    } catch {
        throw ('Microsoft Access COM automation is unavailable for ' + $AccessProgId +
            '. Check whether that version is registered on this Windows PC. ' + $_.Exception.Message)
    }
    $accessVersion = [string]$access.Version
    if ($RequiredAccessMajorVersion -gt 0) {
        $actualMajor = [int]($accessVersion.Split('.')[0])
        if ($actualMajor -ne $RequiredAccessMajorVersion) {
            throw ('Wrong Microsoft Access version: expected major ' +
                $RequiredAccessMajorVersion + ', COM returned ' + $accessVersion +
                ' (ProgID ' + $AccessProgId + '). No database was opened.')
        }
    }
    # Fail closed if the Access version cannot disable automation macros.
    # This is a metadata exporter, not a request to run database startup code.
    try {
        $access.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        if ([int]$access.AutomationSecurity -ne 3) {
            throw 'AutomationSecurity did not remain ForceDisable.'
        }
    } catch {
        throw ('Cannot confirm that VBA macros are disabled for Access automation: ' + $_.Exception.Message)
    }
    $access.OpenCurrentDatabase($workingCopy, $false)
    $database = $access.CurrentDb()

    if (-not $MacrosOnly) {
    foreach ($form in $access.CurrentProject.AllForms) {
        Export-SaveAsText 'form' 2 ([string]$form.Name)
    }
    foreach ($report in $access.CurrentProject.AllReports) {
        Export-SaveAsText 'report' 3 ([string]$report.Name)
    }
    foreach ($module in $access.CurrentProject.AllModules) {
        Export-SaveAsText 'module' 5 ([string]$module.Name)
    }
    foreach ($query in $database.QueryDefs) {
        Export-Query ([string]$query.Name) ([string]$query.SQL)
    }

    }
    foreach ($macro in $access.CurrentProject.AllMacros) {
        Export-SaveAsText 'macro' 4 ([string]$macro.Name)
    }
    if ($IncludeQueryMetadata) {
        try { Export-QueryMetadata }
        catch {
            Add-Entry 'query_metadata' '(DAO query fields)' 'failed' 'access-dao-field-metadata' '' ('HRESULT ' + [string]$_.Exception.HResult)
        }
    }

    if (-not $MacrosOnly) {
    # VBProject access may be disabled by Office Trust Center. Explicitly record
    # that failure; never change security settings automatically.
    try {
        $vbProject = $access.VBE.ActiveVBProject
        foreach ($component in $vbProject.VBComponents) {
            $componentName = [string]$component.Name
            $extension = if ([int]$component.Type -eq 1) { '.bas' } else { '.cls' }
            $target = New-ArtifactPath 'vba' $extension
            try {
                $component.Export($target.absolute)
                Add-Entry 'vba' $componentName 'source_only' 'vbide-export' $target.relative ''
            } catch {
                $originalError = $_.Exception.Message
                # VBIDE.Export is unreliable for some document/class modules.
                # A code-only fallback still obtains actual source lines, but
                # class designer attributes and binary properties are missing.
                try {
                    $count = [int]$component.CodeModule.CountOfLines
                    if ($count -le 0) {
                        throw 'The component exposes no VBA source lines.'
                    }
                    $code = [string]$component.CodeModule.Lines(1, $count)
                    $header = 'Attribute VB_Name = "{0}"' -f $componentName.Replace('"', '""')
                    [IO.File]::WriteAllText($target.absolute,
                        ($header + [Environment]::NewLine + $code),
                        (New-Object System.Text.UTF8Encoding($false)))
                    Add-Entry 'vba' $componentName 'source_only' 'vbide-codemodule-utf8' $target.relative (
                        'VBIDE.Export failed; imported CodeModule source only: ' + $originalError)
                } catch {
                    Add-Entry 'vba' $componentName 'failed' 'vbide-export' '' (
                        $originalError + '; CodeModule fallback: ' + $_.Exception.Message)
                }
            }
        }
    } catch {
        Add-Entry 'vba' '(VBProject)' 'unavailable' 'vbide-export' '' (
            'VBA project access unavailable (Office Trust Center, project lock, or COM error): ' + $_.Exception.Message)
    }
    }
} finally {
    if ($null -ne $database) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($database) } catch {}
        $database = $null
    }
    if ($null -ne $access) {
        try { $access.CloseCurrentDatabase() } catch {}
        try { $access.Quit() } catch {}
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($access) } catch {}
        $access = $null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($null -ne $scratch -and (Test-Path -LiteralPath $scratch)) {
        # Only delete our fresh temporary copy, never the original or export.
        Remove-Item -LiteralPath $scratch -Recurse -Force -ErrorAction SilentlyContinue
    }
}

$afterHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
if ($afterHash -ne $originalHash) {
    throw 'Source file hash changed unexpectedly. Investigate before continuing.'
}

$manifest = [ordered]@{
    format = 'linaccess-bridge-v1'
    exported_utc = [DateTime]::UtcNow.ToString('o')
    access = [ordered]@{
        prog_id = $AccessProgId
        version = $accessVersion
    }
    source = [ordered]@{
        filename = [IO.Path]::GetFileName($source)
        sha256 = $originalHash.ToLowerInvariant()
    }
    objects = @($script:entries.ToArray())
}
$json = ConvertTo-Json -InputObject $manifest -Depth 8
[IO.File]::WriteAllText((Join-Path $destination 'manifest.json'),
    $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Host ('LinAccess bridge export completed: ' + $destination)
Write-Host ('Microsoft Access version: ' + $accessVersion + ' via ' + $AccessProgId)
Write-Host ('Objects recorded: ' + $script:entries.Count)
Write-Host 'Failures and unavailable objects are explicitly listed in manifest.json.'
