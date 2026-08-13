[CmdletBinding()]
param(
    [string]$RepositoryRoot = (
        [System.IO.Path]::GetFullPath(
            (Join-Path $PSScriptRoot "..")))
)

$ErrorActionPreference = "Stop"

$mainWindowPath = Join-Path `
    $RepositoryRoot "ClearPlan.Simulator\MainWindow.xaml.cs"
$mainWindowSource = Get-Content -LiteralPath $mainWindowPath -Raw
$failures = New-Object System.Collections.Generic.List[string]

$visibleActions = [ordered]@{
    "ReportRequested" = "OnReportRequested"
    "OpenPlanRequested" = "OnOpenPlanRequested"
    "DvhResetRequested" = "OnDvhResetRequested"
    "DvhExportRequested" = "OnDvhExportRequested"
    "StructureMappingRequested" = "OnStructureMappingRequested"
}

foreach ($entry in $visibleActions.GetEnumerator()) {
    $subscribePattern =
        "nextViewModel\." +
        [regex]::Escape($entry.Key) +
        "\s*\+=\s*" +
        [regex]::Escape($entry.Value)
    if ($mainWindowSource -notmatch $subscribePattern) {
        $failures.Add(
            "Missing simulator action subscription: " +
            "$($entry.Key) -> $($entry.Value)")
    }

    $unsubscribePattern =
        "workspaceViewModel\." +
        [regex]::Escape($entry.Key) +
        "\s*-=\s*" +
        [regex]::Escape($entry.Value)
    if ($mainWindowSource -notmatch $unsubscribePattern) {
        $failures.Add(
            "Missing simulator action unsubscription: " +
            "$($entry.Key) -> $($entry.Value)")
    }
}

function Require-HandlerText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$HandlerName,

        [Parameter(Mandatory = $true)]
        [string[]]$RequiredPatterns
    )

    $handler = [regex]::Match(
        $mainWindowSource,
        "(?s)private void " +
        [regex]::Escape($HandlerName) +
        "\s*\(.*?(?=\r?\n        private (?:void|int|string|async))")
    if (-not $handler.Success) {
        $failures.Add("Missing simulator action handler: $HandlerName")
        return
    }

    foreach ($pattern in $RequiredPatterns) {
        if ($handler.Value -notmatch $pattern) {
            $failures.Add(
                "Simulator handler '$HandlerName' omits behavior: " +
                $pattern)
        }
    }
}

Require-HandlerText `
    -HandlerName "OnOpenPlanRequested" `
    -RequiredPatterns @(
        "SimulatorSnapshotActions\.GetOpenPlanStatus",
        "SimulatorStatusText\.Text")
Require-HandlerText `
    -HandlerName "OnDvhResetRequested" `
    -RequiredPatterns @(
        "workspaceViewModel\.ResetDvhSelections",
        "SynchronizeSnapshotDvhSelections",
        "SimulatorStatusText\.Text")
Require-HandlerText `
    -HandlerName "OnDvhExportRequested" `
    -RequiredPatterns @(
        "GetDeterministicDvhExportPath",
        "ExportCurrentDvh",
        "InitialDirectory\s*=\s*defaultReportDirectory")
Require-HandlerText `
    -HandlerName "OnStructureMappingRequested" `
    -RequiredPatterns @(
        "SimulatorSnapshotActions\.ApplyStructureMapping",
        "workspaceViewModel\.MarkStructureMappingApplied",
        "SimulatorStatusText\.Text")
Require-HandlerText `
    -HandlerName "OnReportRequested" `
    -RequiredPatterns @(
        "SynchronizeSnapshotDvhSelections",
        "reportService\.Export\s*\(\s*snapshot")

if ($failures.Count -gt 0) {
    $failures | ForEach-Object {
        Write-Error $_ -ErrorAction Continue
    }
    throw "Simulator action validation failed."
}

Write-Output "Simulator action validation passed."
