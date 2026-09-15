[CmdletBinding()]
param([ValidateSet('Debug','Release')][string]$Configuration = 'Debug')
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).ProviderPath
$msbuild = & (Join-Path $PSScriptRoot 'resolve-msbuild.ps1')
Push-Location -LiteralPath $root
try {
    & $msbuild .\ClearPlan.sln '/t:ClearPlan_Core_Tests;ClearPlan_Simulator' "/p:Configuration=$Configuration" /p:Platform=x64 /nologo
    if ($LASTEXITCODE) { throw 'Portable review build failed.' }
    & (Join-Path $root "artifacts\bin\$Configuration\ClearPlan.Core.Tests.exe")
    if ($LASTEXITCODE) { throw 'Portable review tests failed.' }
    & (Join-Path $PSScriptRoot 'validate-public-release.ps1') -Root $root
    if ($LASTEXITCODE) { throw 'Public privacy/configuration validation failed.' }
    Write-Host 'Portable software verification completed. No native TPS or clinical commissioning is implied.'
}
finally { Pop-Location }
