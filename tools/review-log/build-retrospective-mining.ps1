param(
    # Folder with the built native ClearPlan assemblies (ClearPlan.esapi.dll, ClearPlan.Core.dll, ESAPI references ...).
    [string]$NativeDirectory = (Join-Path $PSScriptRoot '..\..\artifacts\clinical-report-build'),
    # Output folder; deploy the exe next to ClearPlan.esapi.dll so it reads the same settings.ini/settings.json.
    [string]$OutputDirectory = '',
    # Same default compiler as tools/clinical-report/build-interoperability-tools.ps1.
    [string]$Compiler = 'C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\Roslyn\csc.exe'
)
# Compiles the read-only retrospective mining entry (ESAPI standalone executable, x64).
# Compile only: this script never starts the executable, opens Eclipse or touches a network share.
$ErrorActionPreference = 'Stop'
$nativeRoot = (Resolve-Path -LiteralPath $NativeDirectory).ProviderPath
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) { $OutputDirectory = $nativeRoot }
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$frameworkRef = 'C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8'
$compiler = $Compiler
if (-not (Test-Path -LiteralPath $compiler -PathType Leaf)) {
    $compiler = & (Join-Path $PSScriptRoot '..\resolve-msbuild.ps1') | Split-Path | Join-Path -ChildPath 'Roslyn\csc.exe'
}
if (-not (Test-Path -LiteralPath $compiler -PathType Leaf)) { throw "Roslyn csc.exe not found: $compiler" }
$refs = @('System.dll', 'System.Core.dll', 'System.Xaml.dll', 'Microsoft.CSharp.dll') +
    @('ClearPlan.Core.dll', 'ClearPlan.Presentation.dll', 'ClearPlan.Rendering.dll', 'ClearPlan.esapi.dll', 'Newtonsoft.Json.dll',
      'OxyPlot.dll', 'OxyPlot.Wpf.dll', 'VMS.TPS.Common.Model.API.dll', 'VMS.TPS.Common.Model.Types.dll' | ForEach-Object { Join-Path $nativeRoot $_ }) +
    @('WindowsBase.dll', 'PresentationCore.dll', 'PresentationFramework.dll' | ForEach-Object { Join-Path $frameworkRef $_ })
$out = Join-Path $OutputDirectory 'ClearPlan.RetrospectiveReviewMining.exe'
& $compiler /nologo /target:exe /platform:x64 ('/out:' + $out) @($refs | ForEach-Object { '/r:' + $_ }) (Join-Path $PSScriptRoot 'mining\RetrospectiveReviewMining.cs')
if ($LASTEXITCODE -ne 0) { throw 'Retrospective mining compile failed.' }
Write-Host "Compiled (not executed): $out"
