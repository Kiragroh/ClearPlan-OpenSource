[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ErrorCalculatorPath,

    [string]$ExpectedHash
)

$ErrorActionPreference = "Stop"
$resolvedPath = (Resolve-Path -LiteralPath $ErrorCalculatorPath).ProviderPath
$content = Get-Content -LiteralPath $resolvedPath -Raw
$hash = (Get-FileHash -LiteralPath $resolvedPath -Algorithm SHA256).Hash

$requiredMethods = @("Calculate", "Calculate2")
$methodInventory = foreach ($method in $requiredMethods) {
    $matches = [regex]::Matches(
        $content,
        "(?m)^\s*public\s+[^\r\n{;]+\s+$method\s*\(")
    [pscustomobject]@{
        name = $method
        count = $matches.Count
    }
}

$missingMethods = @($methodInventory |
    Where-Object { $_.count -ne 1 } |
    ForEach-Object { $_.name })
if ($missingMethods.Count -gt 0) {
    throw "PlanCheck surface is incomplete. Expected exactly one public method: $($missingMethods -join ', ')"
}

$descriptions = @([regex]::Matches(
    $content,
    'AddNewRow\s*\(\s*"((?:[^"\\]|\\.)*)"') |
    ForEach-Object { $_.Groups[1].Value } |
    Select-Object -Unique)
$checkMethods = @([regex]::Matches(
    $content,
    '(?m)^\s*public\s+List<ErrorViewModel>\s+([A-Za-z_][A-Za-z0-9_]*)\s*\(') |
    ForEach-Object { $_.Groups[1].Value })
$addNewRowReferences = [Math]::Max(
    0,
    ([regex]::Matches($content, '\bAddNewRow\s*\(')).Count - 1)
$calculateCalls = ([regex]::Matches($content, '\bCalculate\s*\(')).Count
$calculate2Calls = ([regex]::Matches($content, '\bCalculate2\s*\(')).Count

if (-not [string]::IsNullOrWhiteSpace($ExpectedHash) -and
    $hash -ne $ExpectedHash.ToUpperInvariant()) {
    throw "Clinical ErrorCalculator.cs hash changed. Expected $ExpectedHash but found $hash."
}

[pscustomobject]@{
    path = $resolvedPath
    sha256 = $hash
    bytes = (Get-Item -LiteralPath $resolvedPath).Length
    methods = $methodInventory
    calculateReferences = $calculateCalls
    calculate2References = $calculate2Calls
    checkMethods = $checkMethods
    checkCallCount = $addNewRowReferences
    checkDescriptions = $descriptions
    checkDescriptionCount = $descriptions.Count
}
