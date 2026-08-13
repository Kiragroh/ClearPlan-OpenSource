param(
    [string]$Root = ""
)

$ErrorActionPreference = "Stop"
if ([string]::IsNullOrWhiteSpace($Root)) {
    $Root = Split-Path -Parent $PSScriptRoot
}
$Root = [IO.Path]::GetFullPath($Root)

$patterns = @(
    @{ Name = "Old site-specific binary name"; Pattern = "PlanCheck-UKE" },
    @{ Name = "Old institutional assembly metadata"; Pattern = "Michigan Medicine Radiation Oncology" },
    @{ Name = "Personal assembly company metadata"; Pattern = 'AssemblyCompany\("Maximilian Grohmann"\)' },
    @{ Name = "Site-specific Windows domain"; Pattern = 'oncology\\' },
    @{ Name = "Site-specific meeting-board wording"; Pattern = "InfoTafel|FrÃ¼hbesprechung|Frühbesprechung" },
    @{ Name = "Private clinical UNC host"; Pattern = '\\\\medizin\.uni-leipzig\.de[\\/]' },
    @{ Name = "Private RefDB filename"; Pattern = "RSAlign_refdb_constraint_tables\.local\.json" },
    @{ Name = "Private clinical settings path"; Pattern = "Ersteinstellungen_Termine" },
    @{ Name = "Private deployment overlay"; Pattern = "deployment[\\/]private" }
)

$relativeExcludes = @(
    "tools\validate-public-release.ps1"
)

function Get-RelativePath {
    param(
        [Parameter(Mandatory = $true)][string]$BasePath,
        [Parameter(Mandatory = $true)][string]$FullPath
    )

    $baseUri = New-Object System.Uri(($BasePath.TrimEnd("\") + "\"))
    $fullUri = New-Object System.Uri($FullPath)
    return [System.Uri]::UnescapeDataString($baseUri.MakeRelativeUri($fullUri).ToString()).Replace("/", "\")
}

$git = Get-Command git -ErrorAction SilentlyContinue
if ($null -ne $git) {
    $relativeFiles = & git -C $Root ls-files --cached --others --exclude-standard
    if ($LASTEXITCODE -ne 0) {
        throw "Could not enumerate the public Git worktree."
    }

    $files = $relativeFiles |
        ForEach-Object {
            $candidate = Join-Path $Root $_
            if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                Get-Item -LiteralPath $candidate
            }
        } |
        Where-Object { -not $_.PSIsContainer }
}
else {
    $files = Get-ChildItem -Path $Root -Recurse -File |
        Where-Object {
            $_.FullName -notmatch "\\.git\\" -and
            $_.FullName -notmatch "\\built\\" -and
            $_.FullName -notmatch "\\packages\\" -and
            $_.FullName -notmatch "\\bin\\" -and
            $_.FullName -notmatch "\\obj\\"
        }
}

$files = $files | Where-Object {
    $relative = Get-RelativePath -BasePath $Root -FullPath $_.FullName
    $relativeExcludes -notcontains $relative
}

$findings = New-Object System.Collections.Generic.List[string]
$textExtensions = @(".cs", ".xaml", ".config", ".json", ".ini", ".md", ".cff", ".sln", ".ps1", ".py", ".xml")
$textFiles = $files | Where-Object { $_.Extension -in $textExtensions }

foreach ($entry in $patterns) {
    $matches = Select-String -Path $textFiles.FullName -Pattern $entry.Pattern -AllMatches -ErrorAction SilentlyContinue
    foreach ($match in $matches) {
        $relativePath = Get-RelativePath -BasePath $Root -FullPath $match.Path
        $findings.Add(("{0}:{1}: {2}: {3}" -f $relativePath, $match.LineNumber, $entry.Name, $match.Line.Trim()))
    }
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$officeFiles = $files | Where-Object { $_.Extension -in @(".docx", ".xlsx") }
foreach ($officeFile in $officeFiles) {
    $archive = [System.IO.Compression.ZipFile]::OpenRead($officeFile.FullName)
    try {
        foreach ($zipEntry in $archive.Entries |
            Where-Object { $_.FullName -match "\.(xml|rels)$" }) {
            $stream = $zipEntry.Open()
            $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
            try {
                $xmlText = $reader.ReadToEnd()
            }
            finally {
                $reader.Dispose()
                $stream.Dispose()
            }

            foreach ($patternEntry in $patterns) {
                if ($xmlText -match $patternEntry.Pattern) {
                    $relativePath = Get-RelativePath -BasePath $Root -FullPath $officeFile.FullName
                    $findings.Add(("{0}!{1}: {2}" -f
                        $relativePath,
                        $zipEntry.FullName,
                        $patternEntry.Name))
                }
            }
        }
    }
    finally {
        $archive.Dispose()
    }
}

if ($findings.Count -gt 0) {
    Write-Host "Public release validation failed:" -ForegroundColor Red
    $findings | Sort-Object | ForEach-Object { Write-Host "  $_" }
    exit 1
}

$settingsPath = Join-Path $Root "ClearPlan.Script\Distribution\settings.json"
$settings = Get-Content -Raw $settingsPath | ConvertFrom-Json

if ($null -eq $settings.privacy) {
    Write-Host "Public release validation failed: settings.json has no privacy section." -ForegroundColor Red
    exit 1
}

if ($settings.privacy.logPatientIdentifiers -ne $false -or
    $settings.privacy.logUserIdentifiers -ne $false -or
    $settings.privacy.usePatientIdentifiersInFilenames -ne $false) {
    Write-Host "Public release validation failed: privacy defaults must not log direct identifiers." -ForegroundColor Red
    exit 1
}

$pathProperties = @(
    "paths",
    "clinicalPaths"
)
foreach ($pathProperty in $pathProperties) {
    if ($null -ne $settings.PSObject.Properties[$pathProperty]) {
        Write-Host "Public release validation failed: file-system paths must not be stored in settings.json." -ForegroundColor Red
        exit 1
    }
}
foreach ($sourcePathProperty in @("refDbJsonPath", "excelWorkbookPath")) {
    if ($null -ne $settings.constraintSource.PSObject.Properties[$sourcePathProperty]) {
        Write-Host "Public release validation failed: constraint source paths must be stored in settings.ini." -ForegroundColor Red
        exit 1
    }
}

$pathSettingsPath = Join-Path $Root "ClearPlan.Script\Distribution\settings.ini"
$pathSettingsText = Get-Content -LiteralPath $pathSettingsPath -Raw
$requiredPathKeys = @(
    "RefDbJsonPath",
    "ExcelWorkbookPath",
    "ConstraintTemplatesDirectory",
    "DefaultConventionalTemplate",
    "DefaultHypofractionatedTemplate",
    "DefaultPlanSumTemplate",
    "LogsDirectory",
    "ReportsDirectory",
    "CsvExportDirectory",
    "StateDirectory",
    "UsageLogFile",
    "ActivityLogFile",
    "VersionSeenUsersFile",
    "ChangeLogFile",
    "FeedbackFile",
    "PrescriptionSettingsPath",
    "PlanCheckResourcesPath"
)
foreach ($pathKey in $requiredPathKeys) {
    if ($pathSettingsText -notmatch ("(?m)^\s*" + [regex]::Escape($pathKey) + "\s*=")) {
        Write-Host "Public release validation failed: settings.ini is missing $pathKey." -ForegroundColor Red
        exit 1
    }
}

Write-Host "Public release validation passed."
