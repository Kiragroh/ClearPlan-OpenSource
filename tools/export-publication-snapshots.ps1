[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$SimulatorDirectory,
    [Parameter(Mandatory=$true)][string]$OutputDirectory
)
$ErrorActionPreference='Stop'
$runtime=(Resolve-Path -LiteralPath $SimulatorDirectory).ProviderPath
$output=[IO.Path]::GetFullPath($OutputDirectory)
if($output.StartsWith('\\')) { throw 'Publication evidence requires an explicit local output directory.' }
[Reflection.Assembly]::LoadFrom((Join-Path $runtime 'Newtonsoft.Json.dll')) | Out-Null
$core=[Reflection.Assembly]::LoadFrom((Join-Path $runtime 'ClearPlan.Core.dll'))
$factory=$core.GetType('ClearPlan.Core.Simulation.SyntheticPublicationScenarioFactory',$true)
$create=$factory.GetMethod('Create',[Type[]]@([string]))
$serializer=$core.GetType('ClearPlan.Core.Review.ReviewSnapshotJson',$true)
[IO.Directory]::CreateDirectory($output) | Out-Null
foreach($id in @('publication-single-layer','publication-dual-layer')) {
    # No external snapshot or patient input is accepted by this exporter.
    $snapshot=$create.Invoke($null,@($id))
    if(-not $snapshot.Synthetic -or $snapshot.ScenarioId -ne $id) { throw 'Synthetic factory contract failed.' }
    $text=$serializer.GetMethod('Serialize').Invoke($null,@($snapshot))
    $path=Join-Path $output ($id+'.json')
    [IO.File]::WriteAllText($path,$text,[Text.UTF8Encoding]::new($false))
    Write-Output ('Exported compiled synthetic snapshot: '+$id)
}
