[CmdletBinding()]
param(
    [string]$Root
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($Root)) {
    $Root = Split-Path -Parent $PSScriptRoot
}

$repositoryRoot = [System.IO.Path]::GetFullPath($Root)
$projectPaths = @(
    (Join-Path $repositoryRoot 'ClearPlan.Core\ClearPlan.Core.csproj'),
    (Join-Path $repositoryRoot 'ClearPlan.Presentation\ClearPlan.Presentation.csproj'),
    (Join-Path $repositoryRoot 'ClearPlan.Reporting\ClearPlan.Reporting.csproj'),
    (Join-Path $repositoryRoot 'ClearPlan.Reporting.MigraDoc\ClearPlan.Reporting.MigraDoc.csproj'),
    (Join-Path $repositoryRoot 'ClearPlan.Simulator\ClearPlan.Simulator.csproj')
)
$forbiddenNamePattern = '^(VMS\.TPS\.|EsapiEssentials(?:\.|$))'
$failures = New-Object System.Collections.Generic.List[string]

function Get-AssemblyReferenceNames {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AssemblyPath
    )

    $encodedPath = [Convert]::ToBase64String(
        [Text.Encoding]::UTF8.GetBytes($AssemblyPath))
    $inspectionCommand = @"
`$ErrorActionPreference = 'Stop'
`$path = [Text.Encoding]::UTF8.GetString(
    [Convert]::FromBase64String('$encodedPath'))
`$assembly = [System.Reflection.Assembly]::ReflectionOnlyLoadFrom(`$path)
`$assembly.GetReferencedAssemblies() |
    ForEach-Object { `$_.Name }
"@
    $encodedCommand = [Convert]::ToBase64String(
        [Text.Encoding]::Unicode.GetBytes($inspectionCommand))
    $systemDirectory = [Environment]::GetFolderPath(
        [Environment+SpecialFolder]::System)
    $powershellExecutable = Join-Path `
        $systemDirectory 'WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $powershellExecutable -PathType Leaf)) {
        throw "Windows PowerShell is required for .NET Framework assembly inspection."
    }
    $output = & $powershellExecutable `
        -NoProfile `
        -NonInteractive `
        -EncodedCommand $encodedCommand 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw (($output | ForEach-Object { [string]$_ }) -join [Environment]::NewLine)
    }

    return @($output | ForEach-Object { ([string]$_).Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}

foreach ($projectPath in $projectPaths) {
    if (-not (Test-Path -LiteralPath $projectPath -PathType Leaf)) {
        $failures.Add("Missing reporting project: $projectPath")
        continue
    }

    [xml]$project = Get-Content -LiteralPath $projectPath -Raw
    $references = @(
        $project.Project.ItemGroup.Reference |
            Where-Object { $_ -and $_.Include } |
            ForEach-Object { ([string]$_.Include).Split(',')[0].Trim() }
    )

    foreach ($reference in $references) {
        if ($reference -match $forbiddenNamePattern) {
            $failures.Add(
                "Forbidden project reference '$reference' in $projectPath")
        }
    }
}

$assemblyNames = @(
    'ClearPlan.Core.dll',
    'ClearPlan.Presentation.dll',
    'ClearPlan.Reporting.dll',
    'ClearPlan.Reporting.MigraDoc.dll',
    'ClearPlan.Simulator.exe'
)
$candidateDirectories = @(
    (Join-Path $repositoryRoot 'built'),
    (Join-Path $repositoryRoot 'debug'),
    (Join-Path $repositoryRoot 'artifacts\bin\Debug'),
    (Join-Path $repositoryRoot 'artifacts\bin\Release'),
    (Join-Path $repositoryRoot 'artifacts\simulator\Debug'),
    (Join-Path $repositoryRoot 'artifacts\simulator\Release')
)

foreach ($directory in $candidateDirectories) {
    if (-not (Test-Path -LiteralPath $directory -PathType Container)) {
        continue
    }

    foreach ($assemblyName in $assemblyNames) {
        $assemblyPath = Join-Path $directory $assemblyName
        if (-not (Test-Path -LiteralPath $assemblyPath -PathType Leaf)) {
            continue
        }

        try {
            foreach ($reference in
                     (Get-AssemblyReferenceNames -AssemblyPath $assemblyPath)) {
                if ($reference -match $forbiddenNamePattern) {
                    $failures.Add(
                        "Forbidden assembly reference '$reference' in $assemblyPath")
                }
            }
        }
        catch {
            $failures.Add(
                "Could not inspect assembly references in ${assemblyPath}: $($_.Exception.Message)")
        }
    }
}

if ($failures.Count -gt 0) {
    foreach ($failure in $failures) {
        Write-Host "FAIL $failure" -ForegroundColor Red
    }
    exit 1
}

Write-Host 'PASS public simulator projects and assemblies are vendor-independent.'
exit 0
