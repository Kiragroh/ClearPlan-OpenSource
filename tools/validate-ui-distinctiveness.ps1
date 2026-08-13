param(
    [string]$RepositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)

$xamlFiles = @(
    "ClearPlan.Script/MainView.xaml",
    "ClearPlan.Script/PlanSelectView.xaml",
    "ClearPlan.Script/Styles/ClinicalBlueprint.xaml",
    "ClearPlan.Presentation/Styles/ClinicalBlueprint.xaml",
    "ClearPlan.Script/Views/SettingsView.xaml"
)

$legacyMarkers = @(
    "PapayaWhip",
    "Background=""#00519A""",
    "Background=""WhiteSmoke""",
    "ClearPlanSectionHeader",
    "<RowDefinition Height=""3.5*""/>"
)

$requiredMarkers = @(
    "ClinicalBlueprintSurface",
    "ClinicalBlueprintCobalt",
    "ClinicalBlueprintAmber",
    "ClinicalBlueprintNavigation",
    "ClinicalBlueprintNavigationSelected",
    "Gesamtansicht",
    "PQM",
    "PlanCheck",
    "Feldnamen",
    "DVH",
    "Einstellungen",
    "OverviewPlanGrid",
    "OverviewPqmGrid",
    "OverviewPlanCheckGrid",
    "OverviewFieldGrid",
    "OverviewDvhPlot",
    "MappedStructureDisplay",
    "OverviewPlotModel",
    "DvhStructures",
    "PqmDetailPage",
    "PlanCheckDetailPage",
    "FieldNamesDetailPage",
    "DvhDetailPage",
    "SettingsDetailPage"
)

$failed = $false

foreach ($relativePath in $xamlFiles) {
    $path = Join-Path $RepositoryRoot $relativePath
    if (-not (Test-Path $path)) {
        Write-Error "Missing XAML file: $relativePath"
        $failed = $true
        continue
    }

    $content = Get-Content -LiteralPath $path -Raw

    foreach ($marker in $legacyMarkers) {
        if ($content.Contains($marker)) {
            Write-Error "$relativePath still contains legacy visual marker: $marker"
            $failed = $true
        }
    }
}

$combinedXaml = ($xamlFiles | ForEach-Object {
    Get-Content -LiteralPath (Join-Path $RepositoryRoot $_) -Raw
}) -join "`n"

foreach ($marker in $requiredMarkers) {
    if (-not $combinedXaml.Contains($marker)) {
        Write-Error "Missing open-source GUI style marker: $marker"
        $failed = $true
    }
}

$mainViewModelPath = Join-Path $RepositoryRoot "ClearPlan.Script/ViewModels/MainViewModel.cs"
$mainViewModel = Get-Content -LiteralPath $mainViewModelPath -Raw
$requiredViewModelMarkers = @(
    'NotifyPropertyChanged("PqmSummaries")',
    "Tag = structureId,",
    "OverviewPlotModel.InvalidatePlot(true)"
)

foreach ($marker in $requiredViewModelMarkers) {
    if (-not $mainViewModel.Contains($marker)) {
        Write-Error "Missing integrated-overview view-model contract: $marker"
        $failed = $true
    }
}

$mainViewPath = Join-Path $RepositoryRoot "ClearPlan.Script/MainView.xaml"
$mainView = Get-Content -LiteralPath $mainViewPath -Raw
$overviewFieldGrid = [regex]::Match(
    $mainView,
    '<DataGrid x:Name="OverviewFieldGrid"[\s\S]*?</DataGrid>')

if (-not $overviewFieldGrid.Success) {
    Write-Error "Missing OverviewFieldGrid block."
    $failed = $true
}
else {
    foreach ($binding in @(
        'Header="ID-Status" Binding="{Binding IdStatus}"',
        'Header="Name-Status" Binding="{Binding NameStatus}"'
    )) {
        if (-not $overviewFieldGrid.Value.Contains($binding)) {
            Write-Error "OverviewFieldGrid does not expose separate status: $binding"
            $failed = $true
        }
    }
}

if (-not $mainView.Contains(
    'Content="PDF-Report" Click="PrintButtonClicked"')) {
    Write-Error "Gesamtansicht is missing the direct PDF report action."
    $failed = $true
}

$fieldPreviewViewModelPath = Join-Path $RepositoryRoot `
    "ClearPlan.Script/ViewModels/FieldNamePreviewViewModel.cs"
$fieldPreviewViewModel = Get-Content -LiteralPath `
    $fieldPreviewViewModelPath -Raw
if (-not $fieldPreviewViewModel.Contains('"ID okay"')) {
    Write-Error "Conforming field IDs are not labeled 'ID okay'."
    $failed = $true
}

if ($failed) {
    exit 1
}

Write-Host "UI distinctiveness validation passed."
