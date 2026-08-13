[CmdletBinding()]
param(
    [string]$RepositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)

$ErrorActionPreference = "Stop"

$builderPath = Join-Path $RepositoryRoot "ClearPlan.Script\Review\EsapiReviewSnapshotBuilder.cs"
$mapperPath = Join-Path $RepositoryRoot "ClearPlan.Core\Review\ClinicalReviewValueMapper.cs"
$scriptProjectPath = Join-Path $RepositoryRoot "ClearPlan.Script\ClearPlan.csproj"
$coreProjectPath = Join-Path $RepositoryRoot "ClearPlan.Core\ClearPlan.Core.csproj"
$deployPath = Join-Path $RepositoryRoot "tools\deploy-internal-clearplan.ps1"

foreach ($requiredPath in @(
    $builderPath,
    $mapperPath,
    $scriptProjectPath,
    $coreProjectPath,
    $deployPath
)) {
    if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) {
        throw "Clinical review snapshot source is missing: $requiredPath"
    }
}

$builderSource = Get-Content -LiteralPath $builderPath -Raw
$mapperSource = Get-Content -LiteralPath $mapperPath -Raw
$scriptProjectSource = Get-Content -LiteralPath $scriptProjectPath -Raw
$coreProjectSource = Get-Content -LiteralPath $coreProjectPath -Raw
$deploySource = Get-Content -LiteralPath $deployPath -Raw
$builderCode = [regex]::Replace(
    [regex]::Replace(
        $builderSource,
        '/\*[\s\S]*?\*/',
        ''),
    '(?m)//.*$',
    '')

if ($scriptProjectSource -notmatch
    '<Compile\s+Include="Review\\EsapiReviewSnapshotBuilder\.cs"') {
    throw "ClearPlan.Script does not compile EsapiReviewSnapshotBuilder.cs."
}
if ($coreProjectSource -notmatch
    '<Compile\s+Include="Review\\ClinicalReviewValueMapper\.cs"') {
    throw "ClearPlan.Core does not compile ClinicalReviewValueMapper.cs."
}
foreach ($relativePath in @(
    "ClearPlan.Script\Review\EsapiReviewSnapshotBuilder.cs",
    "ClearPlan.Core\Review\ClinicalReviewValueMapper.cs"
)) {
    if ($deploySource -notmatch [regex]::Escape('"' + $relativePath + '"')) {
        throw "Clinical deployment omits required source: $relativePath"
    }
}

foreach ($requiredPattern in @(
    'ReviewSnapshot\s+Build\s*\(\s*MainViewModel\s+source\s*,\s*ClearPlanSettings\s+settings\s*\)',
    'DoseValuePresentation\.Absolute',
    'VolumePresentation\.Relative',
    'Dispatcher\.VerifyAccess\s*\(',
    'ReviewSnapshotValidator\.Validate\s*\(',
    'DoseUnit\.cGy',
    'DoseUnit\.Gy'
)) {
    if ($builderSource -notmatch $requiredPattern) {
        throw "Snapshot builder contract is missing: $requiredPattern"
    }
}

$forbiddenPatterns = [ordered]@{
    "asynchronous or parallel execution" =
        '\bTask\s*\.\s*(?:Run|Factory)|\bnew\s+(?:Task|Thread)\b|\bThreadPool\s*\.|\bParallel\s*\.|\.AsParallel\s*\(|\busing\s+System\.Threading'
    "filesystem access" =
        '\b(?:File|Directory|FileInfo|DirectoryInfo|FileStream|Path)\s*\.|\bFileStream\s*\('
    "network access" =
        '\b(?:HttpClient|WebClient|WebRequest|Socket|TcpClient|NamedPipeClientStream)\b'
    "PlanCheck recalculation" =
        '\b(?:new\s+ErrorCalculator|GetErrors|GetRefs|Calculate2?|ErrorCalculator)\s*\('
    "patient modification" =
        '\b(?:BeginModifications|SaveModifications|Save|AddCourse|AddExternalPlanSetup|AddPlanSum|AddBeam|AddArcBeam|RemoveBeam|SetPrescription)\s*\('
    "live ESAPI object in DTO state" =
        '\b(?:Structure|PlanningItem|Patient|Course|PlanSetup|PlanSum|DVHData|DoseValue)\s+(?:Structure|PlanningItem|Patient|Course|PlanSetup|PlanSum|DvhData|DoseValue)\s*\{'
}
foreach ($entry in $forbiddenPatterns.GetEnumerator()) {
    if ($builderCode -match $entry.Value) {
        throw "Snapshot builder contains forbidden $($entry.Key): $($Matches[0])"
    }
}

if ($builderSource -notmatch '\bsource\.ErrorGrid\b' -or
    $builderSource -notmatch '\bsource\.RefGrid\b') {
    throw "Snapshot builder must consume the already-populated ErrorGrid and RefGrid."
}
if ($builderSource -match '\bActiveConstraintPath\s*\.\s*ActivePath\b') {
    throw "Snapshot builder must not copy a private constraint source path."
}
if ($builderSource -match '\b(?:UID|PlanningItemUID)\b') {
    throw "Snapshot builder must not copy DICOM or planning-item UIDs."
}
if ($mapperSource -notmatch '\bSanitizeClinicalLabel\s*\(') {
    throw "Pure review mapper must expose the path/UID sanitization helper."
}

Write-Host "Clinical review snapshot builder validation passed."
