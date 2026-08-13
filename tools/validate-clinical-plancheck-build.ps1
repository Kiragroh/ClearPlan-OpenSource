[CmdletBinding()]
param(
    [string]$RepositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)

$ErrorActionPreference = "Stop"
$projectPath = Join-Path $RepositoryRoot "ClearPlan.Script\ClearPlan.csproj"
$deployPath = Join-Path $RepositoryRoot "tools\deploy-internal-clearplan.ps1"
$backupPath = Join-Path $RepositoryRoot "tools\backup-internal-clearplan.ps1"
$readmePath = Join-Path $RepositoryRoot "README.md"
$projectSource = Get-Content -LiteralPath $projectPath -Raw
$deploySource = Get-Content -LiteralPath $deployPath -Raw
$backupSource = Get-Content -LiteralPath $backupPath -Raw
$readmeSource = Get-Content -LiteralPath $readmePath -Raw

function Get-ArrayBody {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Source,

        [Parameter(Mandatory = $true)]
        [string]$VariableName
    )

    $match = [regex]::Match(
        $Source,
        ('\$' + [regex]::Escape($VariableName) +
            '\s*=\s*@\((?<body>[\s\S]*?)\)'))
    if (-not $match.Success) {
        throw "Deployment array is missing: $VariableName"
    }

    return $match.Groups["body"].Value
}

if ($projectSource -notmatch
    '<ErrorCalculatorSource\s+Condition=.*?>Calculators\\ErrorCalculator\.cs</ErrorCalculatorSource>') {
    throw "ClearPlan.csproj does not expose a safe default ErrorCalculatorSource override."
}
if ($projectSource -notmatch
    '<Compile\s+Include="\$\(ErrorCalculatorSource\)"') {
    throw "ClearPlan.csproj does not compile the configured ErrorCalculatorSource."
}
if ($deploySource -notmatch
    '/p:ErrorCalculatorSource=\$clinicalErrorCalculator') {
    throw "Clinical deployment does not compile against the preserved PlanCheck source."
}
if ($deploySource -notmatch
    'clinical-build-manifest\.json') {
    throw "Clinical deployment does not emit and verify a binary provenance manifest."
}
if ($deploySource -notmatch
    'planCheckSourceSha256') {
    throw "Clinical binary provenance does not record the PlanCheck source hash."
}
if ($deploySource -notmatch
    'Assert-CompiledTargetsReplaceable') {
    throw "Clinical deployment does not stop before copying when runtime files are in use."
}

$clinicalArtifactNames =
    Get-ArrayBody -Source $deploySource -VariableName "clinicalArtifactNames"
$builtFiles =
    Get-ArrayBody -Source $deploySource -VariableName "builtFiles"
foreach ($artifactName in @(
    "ClearPlan.Presentation.dll",
    "ClearPlan.Presentation.pdb"
)) {
    if ($clinicalArtifactNames -notmatch
        ('"' + [regex]::Escape($artifactName) + '"')) {
        throw "Clinical build manifest omits required presentation artifact: $artifactName"
    }
    if ($builtFiles -notmatch
        ('"' + [regex]::Escape($artifactName) + '"')) {
        throw "Clinical deployment copy list omits required presentation artifact: $artifactName"
    }
}

if ($deploySource -notmatch
    'ExpectedArtifactNames\s+\$clinicalArtifactNames') {
    throw "Clinical build-manifest verification does not require every configured artifact."
}
if ($deploySource -notmatch
    '\$presentationRoot\s*=\s*Join-Path\s+\$resolvedSource\s+"ClearPlan\.Presentation"') {
    throw "Clinical deployment does not include the shared presentation sources."
}
if ($backupSource -notmatch '"ClearPlan\.Presentation"') {
    throw "Clinical rollback backup omits the shared presentation project."
}
if ($readmeSource -notmatch
    '(?m)^-\s+`ClearPlan\.Presentation`:') {
    throw "Repository layout does not document the shared presentation project."
}

Write-Host "Clinical PlanCheck build provenance validation passed."
