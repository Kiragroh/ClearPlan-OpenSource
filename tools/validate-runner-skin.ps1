[CmdletBinding()]
param(
    [string]$RepositoryRoot
)

$ErrorActionPreference = "Stop"
if ([string]::IsNullOrWhiteSpace($RepositoryRoot)) {
    $RepositoryRoot = (Resolve-Path -LiteralPath (
        [System.IO.Path]::Combine($PSScriptRoot, ".."))).ProviderPath
}

$runnerRoot = Join-Path $RepositoryRoot "ClearPlan.Runner"
$appXaml = Join-Path $runnerRoot "App.xaml"
$appCode = Join-Path $runnerRoot "App.xaml.cs"
$styler = Join-Path $runnerRoot "RunnerWindowStyler.cs"
$styleDictionary = Join-Path $runnerRoot "Styles\ClinicalBlueprintRunner.xaml"

foreach ($required in @($appXaml, $appCode, $styler, $styleDictionary)) {
    if (-not (Test-Path -LiteralPath $required -PathType Leaf)) {
        throw "Runner skin file is missing: $required"
    }
}

$appText = Get-Content -LiteralPath $appCode -Raw
$stylerText = Get-Content -LiteralPath $styler -Raw
$xamlText = Get-Content -LiteralPath $appXaml -Raw
$styleText = Get-Content -LiteralPath $styleDictionary -Raw

if ([regex]::Matches($appText, "ScriptRunner\.Run\(new Script\(\)\);").Count -ne 1) {
    throw "The existing ScriptRunner.Run(new Script()) entry point must occur exactly once."
}

$registerIndex = $appText.IndexOf("RunnerWindowStyler.Register();", [StringComparison]::Ordinal)
$runIndex = $appText.IndexOf("ScriptRunner.Run(new Script());", [StringComparison]::Ordinal)
if ($registerIndex -lt 0 -or $registerIndex -gt $runIndex) {
    throw "RunnerWindowStyler must register before ScriptRunner.Run."
}

foreach ($requiredToken in @(
    'Source="Styles/ClinicalBlueprintRunner.xaml"',
    'EsapiEssentials.PluginRunner',
    'ClearPlan Runner',
    'try',
    'catch'
)) {
    if (($xamlText + $appText + $stylerText) -notmatch [regex]::Escape($requiredToken)) {
        throw "Runner skin contract is missing: $requiredToken"
    }
}

foreach ($requiredStyle in @(
    'x:Key="RunnerNavigationBrush"',
    'TargetType="{x:Type Button}"',
    'TargetType="{x:Type TextBox}"',
    'TargetType="{x:Type ComboBox}"',
    'TargetType="{x:Type DataGrid}"',
    'IsKeyboardFocused'
)) {
    if ($styleText -notmatch [regex]::Escape($requiredStyle)) {
        throw "Runner style contract is missing: $requiredStyle"
    }
}

$forbidden = @(
    "BindingFlags\.NonPublic",
    "\.GetField\(",
    "\.GetProperty\(",
    "CreateApplication\(",
    "BeginModifications\(",
    "SaveModifications\(",
    "AddSetupBeam\("
)
foreach ($pattern in $forbidden) {
    $match = Select-String -Path (Join-Path $runnerRoot "*.cs") -Pattern $pattern
    if ($match) {
        throw "Forbidden runner implementation detected: $pattern"
    }
}

Write-Host "Essentials runner skin validation passed."
