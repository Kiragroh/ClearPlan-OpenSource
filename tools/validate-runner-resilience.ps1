[CmdletBinding()]
param(
    [string]$RepositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)

$ErrorActionPreference = "Stop"

function Assert-Contains {
    param(
        [string]$Path,
        [string]$Pattern,
        [string]$Message
    )

    $content = Get-Content -LiteralPath $Path -Raw
    if ($content -notmatch $Pattern) {
        throw $Message
    }
}

$settingsPath = Join-Path $RepositoryRoot `
    "ClearPlan.Script\Helpers\ClearPlanSettings.cs"
$planSelectPath = Join-Path $RepositoryRoot `
    "ClearPlan.Script\PlanSelectView.xaml.cs"
$runnerAppPath = Join-Path $RepositoryRoot `
    "ClearPlan.Runner\App.xaml.cs"
$reporterPath = Join-Path $RepositoryRoot `
    "ClearPlan.Runner\RunnerFailureReporter.cs"
$runnerProjectPath = Join-Path $RepositoryRoot `
    "ClearPlan.Runner\ClearPlan.Runner.csproj"
$settingsViewPath = Join-Path $RepositoryRoot `
    "ClearPlan.Script\Views\SettingsView.xaml"

Assert-Contains `
    -Path $settingsPath `
    -Pattern "WritablePathFallback\.Ensure(FileParent|Directory)" `
    -Message "ClearPlan settings do not use writable-path fallback handling."
Assert-Contains `
    -Path $settingsPath `
    -Pattern "PathSettingsIni\.Apply" `
    -Message "ClearPlan does not load all paths from settings.ini."
Assert-Contains `
    -Path $settingsPath `
    -Pattern "PathSettingsIni\.Serialize" `
    -Message "ClearPlan does not save edited paths back to settings.ini."
Assert-Contains `
    -Path $planSelectPath `
    -Pattern "TryLogPlanSelection" `
    -Message "Plan selection logging is still allowed to escape into the WPF dispatcher."
Assert-Contains `
    -Path $runnerAppPath `
    -Pattern "DispatcherUnhandledException\s*\+=" `
    -Message "Runner does not register a dispatcher-level failure boundary."
if (-not (Test-Path -LiteralPath $reporterPath -PathType Leaf)) {
    throw "Runner failure reporter source is missing."
}
Assert-Contains `
    -Path $runnerProjectPath `
    -Pattern "Compile Include=`"RunnerFailureReporter\.cs`"" `
    -Message "Runner failure reporter is not compiled into the executable."

$requiredGuiPathBindings = @(
    "RefDbJsonPath",
    "ExcelWorkbookPath",
    "ConstraintTemplatesDirectory",
    "DefaultConventionalTemplate",
    "DefaultHypofractionatedTemplate",
    "DefaultPlanSumTemplate",
    "LogsDirectory",
    "ReportsDirectory",
    "ExportsDirectory",
    "StateDirectory",
    "UsageLogFile",
    "ActivityLogFile",
    "VersionSeenUsersFile",
    "ChangeLogFile",
    "FeedbackFile",
    "PrescriptionSettingsPath",
    "PlanCheckResourcesPath"
)
foreach ($binding in $requiredGuiPathBindings) {
    Assert-Contains `
        -Path $settingsViewPath `
        -Pattern ("Binding\s+" + [regex]::Escape($binding) + "[,}]") `
        -Message "ClearPlan settings GUI is missing the $binding path."
}

Write-Host "Runner path resilience validation passed."
