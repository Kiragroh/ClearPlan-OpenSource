[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
if (Test-Path -LiteralPath $vswhere -PathType Leaf) {
    $resolved = & $vswhere `
        -latest `
        -products * `
        -requires Microsoft.Component.MSBuild `
        -find "MSBuild\**\Bin\MSBuild.exe" |
        Select-Object -First 1

    if (-not [string]::IsNullOrWhiteSpace($resolved) -and
        (Test-Path -LiteralPath $resolved -PathType Leaf)) {
        Write-Output ([System.IO.Path]::GetFullPath($resolved))
        exit 0
    }
}

$frameworkFallback = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe"
if (Test-Path -LiteralPath $frameworkFallback -PathType Leaf) {
    Write-Output $frameworkFallback
    exit 0
}

throw "No usable MSBuild installation was found through vswhere or the .NET Framework fallback."
