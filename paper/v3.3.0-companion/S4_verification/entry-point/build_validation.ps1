$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$stage=Get-Content -LiteralPath (Join-Path $root 'staging.json') -Raw | ConvertFrom-Json
$native=Join-Path $stage.candidate 'A'
$src=Join-Path $root 'native_source'
$compiler='C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\Roslyn\csc.exe'
$framework='C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8'
$refs=@('System.dll','System.Core.dll','System.Xaml.dll','System.Net.Http.dll','Microsoft.CSharp.dll') +
    @('ClearPlan.Core.dll','ClearPlan.Presentation.dll','ClearPlan.Rendering.dll','ClearPlan.Reporting.dll','ClearPlan.Reporting.MigraDoc.dll',
      'ClearPlan.esapi.dll','EsapiEssentials.dll','Newtonsoft.Json.dll','OxyPlot.dll','OxyPlot.Wpf.dll',
      'VMS.TPS.Common.Model.API.dll','VMS.TPS.Common.Model.Types.dll' | ForEach-Object {Join-Path $native $_}) +
    @('WindowsBase.dll','PresentationCore.dll','PresentationFramework.dll' | ForEach-Object {Join-Path $framework $_})
$refArgs=@($refs|ForEach-Object {'/r:'+$_})
$shared=Join-Path $PSScriptRoot 'PhantomValidation.cs'
$report=@('HeadlessPlanReport.cs','NativeMlcDiagnostics.cs','ReportDoseSafety.cs','ReportDvhSelection.cs','ProtectedReportSidecar.cs'|ForEach-Object{Join-Path $src $_})
& $compiler /nologo /target:library /platform:x64 ('/out:'+(Join-Path $native 'ClearPlan.HeadlessReport.dll')) $refArgs $report $shared
if($LASTEXITCODE -ne 0){throw 'Headless compile failed'}
& $compiler /nologo /target:exe /platform:x64 /main:ClearPlan.ReviewLogMining.Program ('/out:'+(Join-Path $native 'ClearPlan.RetrospectiveReviewMining.exe')) $refArgs (Join-Path $src 'RetrospectiveReviewMining.cs') $shared
if($LASTEXITCODE -ne 0){throw 'Mining compile failed'}
& $compiler /nologo /target:exe /platform:x64 /main:ClearPlan.PhantomValidation.ConfigurationSetup ('/out:'+(Join-Path $native 'ConfigurationSetup.exe')) $refArgs $shared
if($LASTEXITCODE -ne 0){throw 'Configuration compile failed'}
Copy-Item -LiteralPath (Join-Path $native 'ClearPlan.HeadlessReport.dll') -Destination (Join-Path $stage.candidate 'B\ClearPlan.HeadlessReport.dll')
& (Join-Path $native 'ConfigurationSetup.exe') (Join-Path $stage.candidate 'Configuration') (Join-Path $stage.candidate 'revision1.json') (Join-Path $stage.candidate 'revision2.json')
if($LASTEXITCODE -ne 0){throw 'Configuration import failed'}
& (Join-Path $native 'ClearPlan.RetrospectiveReviewMining.exe') --input (Join-Path $native 'phantom-plans.csv') --validate-input-only
if($LASTEXITCODE -ne 0){throw 'Mining input check failed'}
Write-Output 'Candidate compilation and explicit phantom list validation completed.'
