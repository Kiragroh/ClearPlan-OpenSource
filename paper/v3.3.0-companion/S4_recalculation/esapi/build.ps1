# Compiles the K1 reference export against the ESAPI assemblies of the isolated
# phantom candidate. Writes only to -OutDir; deploys nothing and changes no catalog.
param(
    [Parameter(Mandatory = $true)][string]$EsapiDir,
    [Parameter(Mandatory = $true)][string]$OutDir
)
$ErrorActionPreference = 'Stop'
$compiler = 'C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\Roslyn\csc.exe'
$framework = 'C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8'
$refs = @('System.dll', 'System.Core.dll', 'System.Web.Extensions.dll', 'WindowsBase.dll' | ForEach-Object { Join-Path $framework $_ }) +
        @('VMS.TPS.Common.Model.API.dll', 'VMS.TPS.Common.Model.Types.dll' | ForEach-Object { Join-Path $EsapiDir $_ })
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$out = Join-Path $OutDir 'ClearPlanK1ReferenceExport.esapi.dll'
& $compiler /nologo /target:library /platform:x64 /optimize+ "/out:$out" ($refs | ForEach-Object { "/r:$_" }) (Join-Path $PSScriptRoot 'ClearPlanK1ReferenceExport.cs')
if ($LASTEXITCODE -ne 0) { throw 'Compilation failed' }
Get-FileHash -Algorithm SHA256 $out | Format-List Path, Hash
