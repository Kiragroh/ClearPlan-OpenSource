[CmdletBinding()]
param(
    [string]$RepositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path,

    [string]$WorkerRoot
)

$ErrorActionPreference = "Stop"

if (-not [string]::IsNullOrWhiteSpace($WorkerRoot)) {
    [System.Reflection.Assembly]::LoadFrom(
        (Join-Path $WorkerRoot "ClearPlan.Core.dll")) | Out-Null
    [System.Reflection.Assembly]::LoadFrom(
        (Join-Path $WorkerRoot "Newtonsoft.Json.dll")) | Out-Null
    $scriptAssembly = [System.Reflection.Assembly]::LoadFrom(
        (Join-Path $WorkerRoot "ClearPlan.esapi.dll"))
    $settingsType = $scriptAssembly.GetType(
        "ClearPlan.Helpers.ClearPlanSettings",
        $true)

    $settings = $settingsType.GetMethod("Reload").Invoke($null, @())
    $settings.Paths.ReportsDirectory = "Reports\EditedInGui"
    $settings.Save()

    $iniPath = Join-Path $WorkerRoot "settings.ini"
    $jsonPath = Join-Path $WorkerRoot "settings.json"
    $iniText = Get-Content -LiteralPath $iniPath -Raw
    $jsonText = Get-Content -LiteralPath $jsonPath -Raw
    if ($iniText -notmatch "(?m)^ReportsDirectory\s*=\s*Reports\\EditedInGui\s*$") {
        throw "Edited GUI path was not persisted in settings.ini."
    }
    if ($jsonText -match '"(paths|clinicalPaths|refDbJsonPath|excelWorkbookPath)"\s*:') {
        throw "settings.json still contains a file-system path property."
    }

    $reloaded = $settingsType.GetMethod("Reload").Invoke($null, @())
    if ($reloaded.Paths.ReportsDirectory -ne "Reports\EditedInGui") {
        throw "Edited settings.ini path did not survive a reload."
    }

    Write-Host "Settings INI save/reload integration passed."
    exit 0
}

$builtRoot = Join-Path $RepositoryRoot "built"
$temporaryRoot = Join-Path `
    ([System.IO.Path]::GetTempPath()) `
    ("ClearPlanSettingsIntegration\" + [guid]::NewGuid().ToString("N"))

try {
    New-Item -ItemType Directory -Path $temporaryRoot | Out-Null
    Copy-Item -Path (Join-Path $builtRoot "*") `
        -Destination $temporaryRoot `
        -Recurse

    & powershell.exe `
        -NoProfile `
        -ExecutionPolicy Bypass `
        -File $PSCommandPath `
        -RepositoryRoot $RepositoryRoot `
        -WorkerRoot $temporaryRoot
    if ($LASTEXITCODE -ne 0) {
        throw "Settings INI worker failed with exit code $LASTEXITCODE."
    }
}
finally {
    if (Test-Path -LiteralPath $temporaryRoot -PathType Container) {
        Remove-Item -LiteralPath $temporaryRoot -Recurse -Force
    }
}
