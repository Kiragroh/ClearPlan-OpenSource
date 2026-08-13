[CmdletBinding()]
param(
    [string]$RepositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)

$ErrorActionPreference = "Stop"
$mainViewCodePath = Join-Path `
    $RepositoryRoot "ClearPlan.Script\MainView.xaml.cs"
$mainViewXamlPath = Join-Path `
    $RepositoryRoot "ClearPlan.Script\MainView.xaml"
$pqmViewModelPath = Join-Path `
    $RepositoryRoot "ClearPlan.Script\ViewModels\PQMSummaryViewModel.cs"

$source = Get-Content -LiteralPath $mainViewCodePath -Raw
$xaml = Get-Content -LiteralPath $mainViewXamlPath -Raw
$pqmViewModelSource = Get-Content -LiteralPath $pqmViewModelPath -Raw

if ($xaml -match '\bMouseMove\s*=' -or
    $source -match '\bUserControl_MouseMove\b' -or
    $source -match '\bpublic\s+int\s+w\s*=') {
    throw "Clinical initialization must not depend on a mouse movement."
}

foreach ($requiredPattern in @(
    'Loaded\s*\+=\s*MainView_Loaded',
    'Unloaded\s*\+=\s*MainView_Unloaded',
    'InitializeClinicalDefaults\s*\(',
    'PqmDefaultReviewPolicy\.ShouldIgnore',
    'pqm\.Ignore\s*=\s*true',
    'RefreshClinicalReviewWorkspace\s*\('
)) {
    if ($source -notmatch $requiredPattern) {
        throw "Deterministic clinical initialization is missing: $requiredPattern"
    }
}

$defaultsMatch = [regex]::Match(
    $source,
    '(?s)private void InitializeClinicalDefaults\(\).*?^\s*}\r?$',
    [Text.RegularExpressions.RegexOptions]::Multiline)
if (-not $defaultsMatch.Success) {
    throw "Could not locate InitializeClinicalDefaults."
}
if ($defaultsMatch.Value -match '\bLogLiveMining\s*\(') {
    throw "Optional network/file mining must not run during initial UI setup."
}
if ($defaultsMatch.Value -match '\.ToString\(\)\.(?:Contains|StartsWith)' -or
    $defaultsMatch.Value -match '\.StructureName\.ToString\(\)') {
    throw "Clinical default initialization is not null-safe."
}
if ($defaultsMatch.Value -match 'StartsWith\(\s*"0\.0"' -or
    $defaultsMatch.Value -match 'structureName\.Contains\(\s*"tr"') {
    throw "Clinical defaults still hide valid zero metrics or broad letter matches."
}

if ($source -match
    'FindChildGroup<CheckBox>\(pqmDataGrid|checkBoxlist\d*\[[a-zA-Z]+\]') {
    throw "PQM review logic still indexes visual checkbox controls."
}
if ($pqmViewModelSource -notmatch
    '(?s)public bool Ignore\s*\{.*?NotifyPropertyChanged\("Ignore"\)') {
    throw "The bound Ignore property does not notify the WPF view."
}
if ($pqmViewModelSource -notmatch
    '(?s)public bool Accept\s*\{.*?NotifyPropertyChanged\("Accept"\)') {
    throw "The bound Accept property does not notify the WPF view."
}

Write-Host "PQM deterministic initialization safety validation passed."
