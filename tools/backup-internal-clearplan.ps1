[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$TargetRoot,

    [string]$BackupRoot
)

$ErrorActionPreference = "Stop"

$resolvedTarget = (Resolve-Path -LiteralPath $TargetRoot).ProviderPath.TrimEnd("\")
if ([System.IO.Path]::GetFileName($resolvedTarget) -ne "ClearPlan_OpenSource") {
    throw "Refusing backup: target must resolve to the exact ClearPlan_OpenSource directory."
}

$targetParent = [System.IO.Path]::GetDirectoryName($resolvedTarget)
if ([string]::IsNullOrWhiteSpace($BackupRoot)) {
    $resolvedBackupRoot = $targetParent
}
else {
    $resolvedBackupRoot = [System.IO.Path]::GetFullPath($BackupRoot).TrimEnd("\")
    if (-not (Test-Path -LiteralPath $resolvedBackupRoot -PathType Container)) {
        New-Item -ItemType Directory -Path $resolvedBackupRoot | Out-Null
    }
    $resolvedBackupRoot = (Resolve-Path -LiteralPath $resolvedBackupRoot).ProviderPath.TrimEnd("\")
}

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backupPath = Join-Path $resolvedBackupRoot "ClearPlan_OpenSource_backup_$stamp"
if (Test-Path -LiteralPath $backupPath) {
    throw "Refusing to overwrite an existing backup: $backupPath"
}

New-Item -ItemType Directory -Path $backupPath | Out-Null

$relativeTargets = @(
    "ClearPlan.Script",
    "ClearPlan.Core",
    "ClearPlan.Presentation",
    "built",
    "ClearPlan.sln",
    "settings.json",
    "settings.ini"
)

$copied = New-Object System.Collections.Generic.List[string]
foreach ($relativeTarget in $relativeTargets) {
    $source = Join-Path $resolvedTarget $relativeTarget
    if (-not (Test-Path -LiteralPath $source)) {
        continue
    }

    Copy-Item -LiteralPath $source -Destination $backupPath -Recurse
    $copied.Add($relativeTarget)
}

if ($copied.Count -eq 0) {
    throw "Backup target contained none of the expected ClearPlan files."
}

$manifest = [ordered]@{
    createdAt = (Get-Date).ToString("o")
    source = $resolvedTarget
    backup = $backupPath
    copied = $copied
}
$manifest | ConvertTo-Json -Depth 4 |
    Set-Content -LiteralPath (Join-Path $backupPath "backup-manifest.json") -Encoding UTF8

Write-Output $backupPath
