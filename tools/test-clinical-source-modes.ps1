[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$CoreAssembly,

    [Parameter(Mandatory)]
    [string]$RefDbPath,

    [Parameter(Mandatory)]
    [string]$WorkbookPath,

    [string]$OutputPath
)

$ErrorActionPreference = "Stop"

foreach ($requiredPath in @($CoreAssembly, $RefDbPath, $WorkbookPath)) {
    if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) {
        throw "Required source-mode test input is missing: $requiredPath"
    }
}

$assemblyPath = (Resolve-Path -LiteralPath $CoreAssembly).ProviderPath
$resolvedRefDbPath = (Resolve-Path -LiteralPath $RefDbPath).ProviderPath
$resolvedWorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).ProviderPath
$assemblyDirectory = Split-Path -Parent $assemblyPath
Get-ChildItem -LiteralPath $assemblyDirectory -Filter "*.dll" |
    ForEach-Object {
        try {
            [void][System.Reflection.Assembly]::LoadFrom($_.FullName)
        }
        catch {
            # Some ESAPI-only dependencies cannot be loaded outside Eclipse.
        }
    }
[void][System.Reflection.Assembly]::LoadFrom($assemblyPath)

function Invoke-SourceMode {
    param(
        [string]$Name,
        [ClearPlan.Core.Constraints.ConstraintSourceMode]$Mode,
        [string]$RefDb,
        [string]$Workbook
    )

    $options = [ClearPlan.Core.Settings.ConstraintSourceOptions]::new()
    $options.Mode = $Mode
    $options.RefDbJsonPath = $RefDb
    $options.ExcelWorkbookPath = $Workbook
    $options.IncludeInactiveTables = $false

    $service = [ClearPlan.Core.Constraints.ConstraintCatalogService]::new()
    $result = $service.Load($options, $assemblyDirectory)
    $statistics = if ($null -eq $result.Catalog) {
        $null
    }
    else {
        $result.Catalog.Statistics
    }
    [pscustomobject]@{
        name = $Name
        activeSource = $result.ActiveSource
        isUsable = $result.IsUsable
        warnings = @($result.Warnings)
        errors = @($result.Errors)
        tables = if ($null -eq $statistics) { 0 } else { $statistics.TotalTables }
        constraints = if ($null -eq $statistics) { 0 } else { $statistics.TotalConstraints }
        structures = if ($null -eq $statistics) { 0 } else { $statistics.TotalStructures }
    }
}

$missingRefDb = Join-Path ([IO.Path]::GetTempPath()) (
    "clearplan-missing-refdb-" + [Guid]::NewGuid().ToString("N") + ".json"
)
$results = @(
    Invoke-SourceMode `
        -Name "automatic-refdb" `
        -Mode ([ClearPlan.Core.Constraints.ConstraintSourceMode]::Automatic) `
        -RefDb $resolvedRefDbPath `
        -Workbook $resolvedWorkbookPath
    Invoke-SourceMode `
        -Name "automatic-excel-fallback" `
        -Mode ([ClearPlan.Core.Constraints.ConstraintSourceMode]::Automatic) `
        -RefDb $missingRefDb `
        -Workbook $resolvedWorkbookPath
    Invoke-SourceMode `
        -Name "explicit-refdb-bad-path" `
        -Mode ([ClearPlan.Core.Constraints.ConstraintSourceMode]::RefDb) `
        -RefDb $missingRefDb `
        -Workbook $resolvedWorkbookPath
    Invoke-SourceMode `
        -Name "explicit-excel" `
        -Mode ([ClearPlan.Core.Constraints.ConstraintSourceMode]::Excel) `
        -RefDb $missingRefDb `
        -Workbook $resolvedWorkbookPath
)
$diagnostics = $results | ConvertTo-Json -Depth 6

if ($results[0].activeSource -ne "RefDB" -or -not $results[0].isUsable) {
    throw "Automatic mode did not select a usable RefDB catalog: $diagnostics"
}
if ($results[1].activeSource -ne "Excel" -or -not $results[1].isUsable) {
    throw "Automatic mode did not fall back to a usable Excel catalog: $diagnostics"
}
if (-not ($results[1].warnings -match "Excel aktiv")) {
    throw "Automatic Excel fallback did not report its source switch."
}
if ($results[2].activeSource -ne "" -or $results[2].isUsable -or
    $results[2].errors.Count -eq 0) {
    throw "Explicit RefDB mode did not fail closed for a missing path."
}
if ($results[3].activeSource -ne "Excel" -or -not $results[3].isUsable) {
    throw "Explicit Excel mode did not select a usable Excel catalog."
}

$json = $diagnostics
if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
    $outputDirectory = Split-Path -Parent $OutputPath
    if (-not [string]::IsNullOrWhiteSpace($outputDirectory)) {
        [IO.Directory]::CreateDirectory($outputDirectory) | Out-Null
    }
    [IO.File]::WriteAllText(
        [IO.Path]::GetFullPath($OutputPath),
        $json + [Environment]::NewLine,
        [Text.UTF8Encoding]::new($false)
    )
}

$json
Write-Host "PASS clinical constraint source modes"
