[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$Source,[Parameter(Mandatory=$true)][string]$Pdf)
$ErrorActionPreference='Stop'
$inputPath=(Resolve-Path -LiteralPath $Source).Path
$outputPath=[IO.Path]::GetFullPath($Pdf)
if([IO.Path]::GetExtension($inputPath) -ne '.docx' -or [IO.Path]::GetExtension($outputPath) -ne '.pdf'){throw 'DOCX input and PDF output required.'}
if(Test-Path -LiteralPath $outputPath){throw 'Use a new PDF output path; do not overwrite an earlier inspected render.'}
[IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($outputPath)) | Out-Null
$wordRender=$null;$documentRender=$null
try {
 $wordRender=New-Object -ComObject Word.Application
 $wordRender.Visible=$false
 $wordRender.DisplayAlerts=0
 $wordRender.AutomationSecurity=3
 $documentRender=$wordRender.Documents.Open($inputPath,$false,$true)
 $documentRender.Repaginate()
 $pageCount=$documentRender.ComputeStatistics(2)
 $documentRender.ExportAsFixedFormat($outputPath,17,$false,0,0,1,1,0,$true,$true,0,$true,$true,$false)
 [pscustomobject]@{Pages=$pageCount;Bytes=(Get-Item -LiteralPath $outputPath).Length;Pdf=$outputPath;SourceReadOnly=$true} | ConvertTo-Json
}
finally {
 if($null -ne $documentRender){$documentRender.Close(0);[Runtime.InteropServices.Marshal]::ReleaseComObject($documentRender) | Out-Null}
 if($null -ne $wordRender){$wordRender.Quit(0);[Runtime.InteropServices.Marshal]::ReleaseComObject($wordRender) | Out-Null}
}
