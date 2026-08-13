$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).ProviderPath
$settingsPath = Join-Path $repoRoot "ClearPlan.Script\Helpers\ClearPlanSettings.cs"
$mainViewPath = Join-Path $repoRoot "ClearPlan.Script\MainView.xaml"
$mainViewCodePath = Join-Path $repoRoot "ClearPlan.Script\MainView.xaml.cs"

$settings = Get-Content -LiteralPath $settingsPath -Raw
$mainView = Get-Content -LiteralPath $mainViewPath -Raw
$mainViewCode = Get-Content -LiteralPath $mainViewCodePath -Raw

if ($settings -notmatch "NetworkSourceTimeout" -or
    $settings -notmatch "TaskScheduler\.Default" -or
    $settings -notmatch "\.Wait\(NetworkSourceTimeout\)") {
    throw "Constraint-source UNC access is not bounded off the UI thread."
}

if ($settings -notmatch
    "resolvedConfiguredPath\s*=\s*ResolvePath\(configuredPath\)" -or
    $settings -notmatch
    "IsUncPath\(resolvedConfiguredPath\)") {
    throw "Relative sources resolving below a UNC base are not bounded."
}

if ($mainView -match "<views:SettingsView") {
    throw "The hidden legacy surface still constructs SettingsView eagerly."
}

if ($mainViewCode -notmatch "LegacySettingsTab_Selected" -or
    $mainViewCode -notmatch "new\s+ClearPlan\.Views\.SettingsView") {
    throw "The legacy settings page is not created lazily on selection."
}

$refreshCallCount = (
    [regex]::Matches(
        $mainViewCode,
        "_clinicalReviewHost\.TryRefresh\(\)")).Count
if ($refreshCallCount -ne 1) {
    throw "The clinical snapshot must be refreshed from exactly one guarded call site."
}

Write-Host "Network source and clinical startup bounds validation passed."
