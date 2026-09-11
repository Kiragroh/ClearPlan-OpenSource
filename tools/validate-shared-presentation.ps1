param(
    [string]$Root = (Split-Path -Parent $PSScriptRoot)
)

$ErrorActionPreference = 'Stop'

$resolvedRoot = (Resolve-Path -LiteralPath $Root).Path
$projectDirectory = Join-Path $resolvedRoot 'ClearPlan.Presentation'
$projectPath = Join-Path $projectDirectory 'ClearPlan.Presentation.csproj'
$viewPath = Join-Path $projectDirectory 'Views\ReviewWorkspaceView.xaml'
$viewModelPath = Join-Path $projectDirectory 'ViewModels\ReviewWorkspaceViewModel.cs'
$plotFactoryPath = Join-Path $projectDirectory 'Plot\ReviewPlotFactory.cs'
$resourceAnchorPath = Join-Path $projectDirectory 'ResourceAssemblyAnchor.cs'
$sharedStylePath = Join-Path $projectDirectory 'Styles\ClinicalBlueprint.xaml'
$clinicalStylePath = Join-Path $resolvedRoot 'ClearPlan.Script\Styles\ClinicalBlueprint.xaml'
$clinicalProjectPath = Join-Path $resolvedRoot 'ClearPlan.Script\ClearPlan.csproj'
$clinicalViewCodePath = Join-Path $resolvedRoot 'ClearPlan.Script\MainView.xaml.cs'
$clinicalAssemblyPath = Join-Path $resolvedRoot 'built\ClearPlan.esapi.dll'

$failures = New-Object System.Collections.Generic.List[string]

function Require-File {
    param(
        [string]$Path,
        [string]$Description
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        $script:failures.Add("Missing $Description`: $Path")
        return $false
    }

    return $true
}

function Require-Text {
    param(
        [string]$Text,
        [string]$Pattern,
        [string]$Description
    )

    if ($Text -notmatch $Pattern) {
        $script:failures.Add("Missing contract: $Description")
    }
}

function Reject-Text {
    param(
        [string]$Text,
        [string]$Pattern,
        [string]$Description
    )

    if ($Text -match $Pattern) {
        $script:failures.Add("Forbidden presentation dependency/access: $Description")
    }
}

$requiredFiles = @(
    @{ Path = $projectPath; Description = 'shared presentation project' },
    @{ Path = $viewPath; Description = 'shared review workspace view' },
    @{ Path = $viewModelPath; Description = 'snapshot-backed review workspace view model' },
    @{ Path = $plotFactoryPath; Description = 'review plot factory' },
    @{ Path = $sharedStylePath; Description = 'shared Clinical Blueprint resources' },
    @{ Path = $clinicalStylePath; Description = 'clinical style dictionary bridge' },
    @{ Path = $clinicalProjectPath; Description = 'clinical project' },
    @{ Path = $clinicalViewCodePath; Description = 'clinical view code-behind' },
    @{ Path = $clinicalAssemblyPath; Description = 'built clinical assembly' }
)

$allFilesPresent = $true
foreach ($requiredFile in $requiredFiles) {
    if (-not (Require-File -Path $requiredFile.Path -Description $requiredFile.Description)) {
        $allFilesPresent = $false
    }
}

if ($allFilesPresent) {
    [xml]$projectXml = Get-Content -LiteralPath $projectPath -Raw
    $namespaceManager = New-Object System.Xml.XmlNamespaceManager($projectXml.NameTable)
    $namespaceManager.AddNamespace('msb', 'http://schemas.microsoft.com/developer/msbuild/2003')

    $targetFramework = $projectXml.SelectSingleNode(
        '//msb:TargetFrameworkVersion',
        $namespaceManager)
    if ($null -eq $targetFramework -or $targetFramework.InnerText -ne 'v4.8') {
        $failures.Add('ClearPlan.Presentation must target .NET Framework 4.8.')
    }

    $projectReferences = @(
        $projectXml.SelectNodes('//msb:ProjectReference', $namespaceManager) |
            ForEach-Object { $_.Include }
    )
    if ($projectReferences.Count -ne 1 -or
        $projectReferences[0] -notmatch 'ClearPlan\.Core\\ClearPlan\.Core\.csproj$') {
        $failures.Add(
            'ClearPlan.Presentation must reference exactly ClearPlan.Core as a project.')
    }

    $allowedAssemblyReferences = @(
        'OxyPlot',
        'OxyPlot.Wpf',
        'PresentationCore',
        'PresentationFramework',
        'System',
        'System.Core',
        'System.Xaml',
        'WindowsBase'
    )
    $assemblyReferences = @(
        $projectXml.SelectNodes('//msb:Reference', $namespaceManager) |
            ForEach-Object { ($_.Include -split ',')[0] }
    )
    foreach ($reference in $assemblyReferences) {
        if ($allowedAssemblyReferences -notcontains $reference) {
            $failures.Add("Unexpected assembly reference in Presentation: $reference")
        }
    }
    foreach ($reference in @('OxyPlot', 'OxyPlot.Wpf', 'PresentationCore',
            'PresentationFramework')) {
        if ($assemblyReferences -notcontains $reference) {
            $failures.Add("Missing required Presentation assembly reference: $reference")
        }
    }

    $presentationFiles = Get-ChildItem -LiteralPath $projectDirectory -Recurse -File |
        Where-Object { $_.Extension -in @('.cs', '.xaml', '.csproj') }
    $presentationText = ($presentationFiles |
        ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw }) -join "`n"
    $presentationCode = ($presentationFiles |
        Where-Object { $_.Extension -eq '.cs' } |
        ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw }) -join "`n"

    Reject-Text -Text $presentationText `
        -Pattern '(?i)VMS\.TPS|EsapiEssentials|ClearPlan\.Reporting' `
        -Description 'TPS, EsapiEssentials, or Reporting reference'
    Reject-Text -Text $presentationCode `
        -Pattern '(?m)^\s*using\s+System\.(IO|Net)\s*;' `
        -Description 'System.IO or System.Net namespace'
    Reject-Text -Text $presentationCode `
        -Pattern '(?i)\b(File|Directory|FileInfo|DirectoryInfo|WebClient|HttpClient)\s*\.' `
        -Description 'file or network API call'
    Reject-Text -Text $presentationCode `
        -Pattern '(?i)CreateApplication|OpenPatient|BeginModifications|SaveModifications' `
        -Description 'TPS runtime or mutation API'

    $viewText = Get-Content -LiteralPath $viewPath -Raw
    foreach ($controlName in @(
            'OverallReviewView',
            'OverviewSourceGrid',
            'OverviewPlanGrid',
            'OverviewPqmGrid',
            'OverviewPlanCheckGrid',
            'OverviewFieldGrid',
            'OverviewStructureMappingGrid',
            'OverviewDvhPlot',
            'PqmDetailGrid',
            'PlanCheckDetailGrid',
            'FieldDetailGrid',
            'StructureMappingDetailGrid',
            'DvhDetailPlot',
            'DvhFilterTextBox',
            'DvhSelectedOnlyToggle',
            'DvhSelectAllButton',
            'DvhClearSelectionButton',
            'DvhSeriesList',
            'ReportButton')) {
        Require-Text -Text $viewText `
            -Pattern ("x:Name=`"{0}`"" -f [regex]::Escape($controlName)) `
            -Description "stable control name $controlName"
    }
    foreach ($tabName in @(
            'OverviewTab',
            'PqmTab',
            'PlanCheckTab',
            'FieldsTab',
            'DvhTab')) {
        Require-Text -Text $viewText `
            -Pattern ("x:Name=`"{0}`"" -f [regex]::Escape($tabName)) `
            -Description "stable tab name $tabName"
    }

    $overviewScrollBlock = [regex]::Match(
        $viewText,
        '(?s)<ScrollViewer\s+x:Name="OverallReviewView".*?>')
    if (-not $overviewScrollBlock.Success) {
        $failures.Add('Missing OverallReviewView ScrollViewer block.')
    }
    else {
        Require-Text -Text $overviewScrollBlock.Value `
            -Pattern 'HorizontalScrollBarVisibility="Disabled"' `
            -Description 'overview tables measured to the finite viewport width'
    }
    foreach ($binding in @(
            'ModeBadgeText',
            'ModeBadgeDescription',
            'OverviewPlotModel',
            'DetailPlotModel',
            'FilteredDvhSeries',
            'DvhFilterText',
            'ShowSelectedDvhOnly',
            'SelectAllDvhCommand',
            'ClearDvhSelectionCommand',
            'ReportCommand',
            'OpenPlanCommand',
            'ResetDvhCommand',
            'ExportDvhCommand')) {
        Require-Text -Text $viewText `
            -Pattern ("Binding(?:\s+DataContext\.)?\s*{0}" -f [regex]::Escape($binding)) `
            -Description "binding $binding"
    }

    $viewModelText = Get-Content -LiteralPath $viewModelPath -Raw
    foreach ($propertyName in @(
            'OverviewPlotModel',
            'DetailPlotModel',
            'FilteredDvhSeries',
            'DvhFilterText',
            'ShowSelectedDvhOnly',
            'SelectAllDvhCommand',
            'ClearDvhSelectionCommand',
            'ModeBadgeText',
            'ModeBadgeDescription',
            'ReportCommand',
            'OpenPlanCommand',
            'ResetDvhCommand',
            'ExportDvhCommand',
            'NavigateCommand')) {
        Require-Text -Text $viewModelText `
            -Pattern ("\b{0}\b" -f [regex]::Escape($propertyName)) `
            -Description "view-model host surface $propertyName"
    }
    $plotCreationCount = (
        [regex]::Matches(
            $viewModelText,
            'ReviewPlotFactory\.Create\s*\(')).Count
    if ($plotCreationCount -lt 2) {
        $failures.Add(
            'OverviewPlotModel and DetailPlotModel must be created independently.')
    }

    $dvhPlotBlock = [regex]::Match(
        $viewText,
        '<oxy:PlotView\s+x:Name="DvhDetailPlot"[\s\S]*?/>')
    if (-not $dvhPlotBlock.Success) {
        $failures.Add('Missing DvhDetailPlot XAML block.')
    }
    else {
        Require-Text -Text $dvhPlotBlock.Value `
            -Pattern 'MinWidth="0"' `
            -Description 'DvhDetailPlot uses finite host width'
        Require-Text -Text $dvhPlotBlock.Value `
            -Pattern 'MinHeight="0"' `
            -Description 'DvhDetailPlot uses finite host height'
    }
    Require-Text -Text $viewText `
        -Pattern '<ColumnDefinition\s+x:Name="DvhStructureColumn"[\s\S]*?Width="210"[\s\S]*?MinWidth="190"[\s\S]*?MaxWidth="300"\s*/>' `
        -Description 'bounded resizable DVH structure panel'
    Require-Text -Text $viewText `
        -Pattern '<ColumnDefinition\s+Width="5"\s*/>' `
        -Description '5 px DVH splitter column'
    Require-Text -Text $viewText `
        -Pattern '<GridSplitter[\s\S]*?Width="5"[\s\S]*?/>' `
        -Description '5 px DVH GridSplitter'

    $clinicalStyleText = Get-Content -LiteralPath $clinicalStylePath -Raw
    Require-Text -Text $clinicalStyleText `
        -Pattern 'ClearPlan\.Presentation;component/Styles/ClinicalBlueprint\.xaml' `
        -Description 'clinical dictionary bridge to shared resources'

    [xml]$clinicalProjectXml =
        Get-Content -LiteralPath $clinicalProjectPath -Raw
    $clinicalNamespaceManager =
        New-Object System.Xml.XmlNamespaceManager(
            $clinicalProjectXml.NameTable)
    $clinicalNamespaceManager.AddNamespace(
        'msb',
        'http://schemas.microsoft.com/developer/msbuild/2003')
    $clinicalProjectReferences = @(
        $clinicalProjectXml.SelectNodes(
            '//msb:ProjectReference',
            $clinicalNamespaceManager) |
            ForEach-Object { $_.Include }
    )
    if (-not ($clinicalProjectReferences |
            Where-Object {
                $_ -match
                    'ClearPlan\.Presentation\\ClearPlan\.Presentation\.csproj$'
            })) {
        $failures.Add(
            'ClearPlan.Script must reference ClearPlan.Presentation so the shared resource assembly is deployed.')
    }

    [void](Require-File `
        -Path $resourceAnchorPath `
        -Description 'shared resource assembly anchor')

    $clinicalViewCode = Get-Content -LiteralPath $clinicalViewCodePath -Raw
    Require-Text -Text $clinicalViewCode `
        -Pattern 'ResourceAssemblyAnchor\.EnsureLoaded\s*\(\s*\)\s*;' `
        -Description 'side-effect-free presentation assembly anchor'
    $anchorPosition = $clinicalViewCode.IndexOf(
        'ResourceAssemblyAnchor.EnsureLoaded',
        [System.StringComparison]::Ordinal)
    $initializePosition = $clinicalViewCode.IndexOf(
        'InitializeComponent();',
        [System.StringComparison]::Ordinal)
    if ($anchorPosition -lt 0 -or
        $initializePosition -lt 0 -or
        $anchorPosition -gt $initializePosition) {
        $failures.Add(
            'The presentation assembly anchor must execute before InitializeComponent().')
    }

    try {
        $clinicalAssembly =
            [System.Reflection.Assembly]::LoadFile($clinicalAssemblyPath)
        $clinicalReferenceNames = @(
            $clinicalAssembly.GetReferencedAssemblies() |
                ForEach-Object { $_.Name }
        )
        if ($clinicalReferenceNames -notcontains 'ClearPlan.Presentation') {
            $failures.Add(
                'Built ClearPlan.esapi.dll does not reference ClearPlan.Presentation.')
        }
    }
    catch {
        $failures.Add(
            "Could not inspect clinical assembly references: $($_.Exception.Message)")
    }

    $sharedStyleText = Get-Content -LiteralPath $sharedStylePath -Raw
    foreach ($styleKey in @(
            'ClinicalBlueprintNavigation',
            'ClinicalBlueprintNavigationSelected',
            'ClinicalBlueprintAmber',
            'ClinicalBlueprintTeal',
            'ClinicalBlueprintCard',
            'ClinicalBlueprintNavigationTab',
            'ClinicalBlueprintPrimaryButton')) {
        Require-Text -Text $sharedStyleText `
            -Pattern ("x:Key=`"{0}`"" -f [regex]::Escape($styleKey)) `
            -Description "shared style/resource $styleKey"
    }
}

if ($failures.Count -gt 0) {
    Write-Host 'Shared presentation validation FAILED:' -ForegroundColor Red
    foreach ($failure in $failures) {
        Write-Host " - $failure" -ForegroundColor Red
    }
    exit 1
}

Write-Host 'Shared presentation validation passed.' -ForegroundColor Green
