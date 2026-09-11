[CmdletBinding()]
param(
    [string]$RepositoryRoot
)

$ErrorActionPreference = "Stop"
if ([string]::IsNullOrWhiteSpace($RepositoryRoot)) {
    $scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
    $RepositoryRoot = [System.IO.Path]::GetFullPath(
        (Join-Path $scriptDirectory ".."))
}
$failures = New-Object System.Collections.Generic.List[string]

function Read-RequiredSource {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RelativePath
    )

    $path = Join-Path $RepositoryRoot $RelativePath
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        $failures.Add("Missing responsive review source: $RelativePath")
        return ""
    }

    return Get-Content -LiteralPath $path -Raw
}

function Require-Pattern {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Text,

        [Parameter(Mandatory = $true)]
        [string]$Pattern,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    if ($Text -notmatch $Pattern) {
        $failures.Add("Missing responsive contract: $Description")
    }
}

$policySource = Read-RequiredSource `
    "ClearPlan.Core\Review\ReviewWindowSizePolicy.cs"
$coreProject = Read-RequiredSource `
    "ClearPlan.Core\ClearPlan.Core.csproj"
$scriptSource = Read-RequiredSource `
    "ClearPlan.Script\Script.cs"
$planSelectSource = Read-RequiredSource `
    "ClearPlan.Script\PlanSelectView.xaml.cs"
$simulatorSource = Read-RequiredSource `
    "ClearPlan.Simulator\MainWindow.xaml.cs"
$simulatorXaml = Read-RequiredSource `
    "ClearPlan.Simulator\MainWindow.xaml"
$simulatorArguments = Read-RequiredSource `
    "ClearPlan.Simulator\SimulatorArguments.cs"
$simulatorProject = Read-RequiredSource `
    "ClearPlan.Simulator\ClearPlan.Simulator.csproj"
$layoutProbeSource = Read-RequiredSource `
    "ClearPlan.Simulator\ResponsiveLayoutProbeService.cs"
$layoutProbeTest = Read-RequiredSource `
    "tools\test-responsive-layout.ps1"
$simulatorSmoke = Read-RequiredSource `
    "tools\test-simulator.ps1"
$sharedXaml = Read-RequiredSource `
    "ClearPlan.Presentation\Views\ReviewWorkspaceView.xaml"
$legacyXaml = Read-RequiredSource `
    "ClearPlan.Script\MainView.xaml"

foreach ($contract in @(
    @("MinimumWidth\s*=\s*1180", "minimum width 1180"),
    @("MinimumHeight\s*=\s*720", "minimum height 720"),
    @("MaximumInitialWidth\s*=\s*1680", "maximum initial width 1680"),
    @("WidthFraction\s*=\s*0\.92", "width fraction 0.92"),
    @("HeightFraction\s*=\s*0\.88", "height fraction 0.88"),
    @("double\.IsNaN", "NaN rejection"),
    @("double\.IsInfinity", "infinity rejection")
)) {
    Require-Pattern `
        -Text $policySource `
        -Pattern $contract[0] `
        -Description $contract[1]
}

foreach ($contract in @(
    @('IsLayoutProbeMode', "layout-probe mode"),
    @('"--layout-probe"', "layout-probe size option"),
    @('"--output"', "layout-probe JSON option")
)) {
    Require-Pattern `
        -Text $simulatorArguments `
        -Pattern $contract[0] `
        -Description $contract[1]
}
Require-Pattern `
    -Text $simulatorProject `
    -Pattern 'Compile Include="ResponsiveLayoutProbeService\.cs"' `
    -Description "layout-probe service compiled by simulator"
foreach ($contract in @(
    @('workspaceFitsViewport', "rendered workspace viewport result"),
    @('footerFullyVisible', "rendered footer result"),
    @('sourceStatusFullyVisible', "rendered source-status result"),
    @('dvhContentWidth', "rendered DVH content width"),
    @('dvhPlotWidth', "rendered DVH plot width"),
    @('dvhPlotHeight', "rendered DVH plot height"),
    @('UpdateLayout', "real WPF layout execution")
)) {
    Require-Pattern `
        -Text $layoutProbeSource `
        -Pattern $contract[0] `
        -Description $contract[1]
}
Require-Pattern `
    -Text $layoutProbeTest `
    -Pattern '1180\.0' `
    -Description "1180x720 rendered layout probe"
Require-Pattern `
    -Text $layoutProbeTest `
    -Pattern '1256\.72' `
    -Description "1256.72x720 rendered layout probe"
Require-Pattern `
    -Text $simulatorSmoke `
    -Pattern 'test-responsive-layout\.ps1' `
    -Description "layout probe integrated in simulator smoke"
Require-Pattern `
    -Text $coreProject `
    -Pattern 'Compile Include="Review\\ReviewWindowSizePolicy\.cs"' `
    -Description "policy compiled by ClearPlan.Core"

if (([regex]::Matches(
    $scriptSource,
    "ReviewWindowSizePolicy\.Calculate")).Count -ne 1) {
    $failures.Add(
        "Clinical host must calculate responsive size exactly once.")
}
foreach ($contract in @(
    @("SystemParameters\.WorkArea", "clinical work area"),
    @("mainWindow\.Width\s*=\s*windowSize\.Width", "clinical width"),
    @("mainWindow\.Height\s*=\s*windowSize\.Height", "clinical height"),
    @("mainWindow\.MinWidth\s*=\s*ReviewWindowSizePolicy\.MinimumWidth", "clinical minimum width"),
    @("mainWindow\.MinHeight\s*=\s*ReviewWindowSizePolicy\.MinimumHeight", "clinical minimum height"),
    @("HorizontalContentAlignment\s*=\s*HorizontalAlignment\.Stretch", "clinical horizontal stretch"),
    @("VerticalContentAlignment\s*=\s*VerticalAlignment\.Stretch", "clinical vertical stretch"),
    @("SizeToContent\s*=\s*SizeToContent\.Manual", "clinical manual sizing"),
    @("ResizeMode\s*=\s*ResizeMode\.CanResizeWithGrip", "clinical resize grip")
)) {
    Require-Pattern `
        -Text $scriptSource `
        -Pattern $contract[0] `
        -Description $contract[1]
}
$workspaceRoot = [regex]::Match(
    $sharedXaml,
    '(?s)<UserControl\s+.*?>').Value
Require-Pattern `
    -Text $workspaceRoot `
    -Pattern 'MinHeight="600"' `
    -Description "workspace minimum fits a 720-high outer window"
Require-Pattern `
    -Text $sharedXaml `
    -Pattern 'x:Name="NavigationFooter"' `
    -Description "measurable navigation footer"
Require-Pattern `
    -Text $sharedXaml `
    -Pattern 'x:Name="NavigationSourceStatusText"' `
    -Description "measurable navigation source status"
if ($scriptSource -match "PrimaryScreenHeight\s*\*\s*0\.85") {
    $failures.Add("Clinical host still uses the legacy height-only sizing.")
}
if ($planSelectSource -match
    "(?:MainWindow\.Height|PrimaryScreenHeight)") {
    $failures.Add("Plan selection still resets the host height.")
}

foreach ($contract in @(
    @("ReviewWindowSizePolicy\.Calculate", "simulator policy application"),
    @("SystemParameters\.WorkArea", "simulator work area"),
    @("ResizeMode\s*=\s*ResizeMode\.CanResizeWithGrip", "interactive simulator resize grip"),
    @("Width\s*=\s*VisualCaptureService\.CaptureWidth", "fixed capture width"),
    @("Height\s*=\s*VisualCaptureService\.CaptureHeight", "fixed capture height"),
    @("ResizeMode\s*=\s*ResizeMode\.NoResize", "fixed capture resize mode")
)) {
    Require-Pattern `
        -Text $simulatorSource `
        -Pattern $contract[0] `
        -Description $contract[1]
}

$simulatorWindowTag = [regex]::Match(
    $simulatorXaml,
    '(?s)<Window\s+.*?>').Value
if ($simulatorWindowTag -match
    '(?m)\s(?:Width|Height|MinWidth|MinHeight)="\d') {
    $failures.Add(
        "Simulator XAML still fixes the interactive window size.")
}
Require-Pattern `
    -Text $simulatorWindowTag `
    -Pattern 'HorizontalContentAlignment="Stretch"' `
    -Description "simulator horizontal content stretch"
Require-Pattern `
    -Text $simulatorWindowTag `
    -Pattern 'VerticalContentAlignment="Stretch"' `
    -Description "simulator vertical content stretch"

foreach ($contract in @(
    @('x:Name="DvhContentGrid"', "named stretching DVH content grid"),
    @('x:Name="DvhStructurePanel"', "named DVH structure panel"),
    @('Width="210"', "compact DVH structure pane width"),
    @('MinWidth="190"', "resizable DVH structure pane minimum"),
    @('MaxWidth="300"', "resizable DVH structure pane maximum"),
    @('<ColumnDefinition Width="5"/>', "DVH splitter width"),
    @('<ColumnDefinition Width="\*"/>', "DVH plot star width"),
    @('x:Name="DvhDetailPlot"', "DVH detail plot"),
    @('MinWidth="0"', "DVH plot can consume finite host width"),
    @('MinHeight="0"', "DVH plot can consume finite host height")
)) {
    Require-Pattern `
        -Text $sharedXaml `
        -Pattern $contract[0] `
        -Description $contract[1]
}
$dvhTabBlock = [regex]::Match(
    $sharedXaml,
    '(?s)<TabItem\s+x:Name="DvhTab".*?</TabItem>').Value
Require-Pattern `
    -Text $dvhTabBlock `
    -Pattern '<Grid Margin="16"' `
    -Description "DVH page margin 16"
Require-Pattern `
    -Text $dvhTabBlock `
    -Pattern 'Grid\.Column="2"[^>]*\sPadding="10"' `
    -Description "compact DVH plot-card padding"
if ($dvhTabBlock -match '<ScrollViewer') {
    $failures.Add(
        "DVH detail layout must not be measured through a ScrollViewer.")
}

$legacyRoot = [regex]::Match(
    $legacyXaml,
    '(?s)<UserControl\s+.*?>').Value
Require-Pattern `
    -Text $legacyRoot `
    -Pattern 'HorizontalContentAlignment="Stretch"' `
    -Description "legacy root horizontal stretch"
Require-Pattern `
    -Text $legacyRoot `
    -Pattern 'VerticalContentAlignment="Stretch"' `
    -Description "legacy root vertical stretch"
$legacyNavigation = [regex]::Match(
    $legacyXaml,
    '(?s)<TabControl\s+x:Name="MainNavigation".*?>').Value
Require-Pattern `
    -Text $legacyNavigation `
    -Pattern 'HorizontalContentAlignment="Stretch"' `
    -Description "legacy TabControl horizontal stretch"
Require-Pattern `
    -Text $legacyNavigation `
    -Pattern 'VerticalContentAlignment="Stretch"' `
    -Description "legacy TabControl vertical stretch"
$legacyContent = [regex]::Match(
    $legacyXaml,
    '(?s)<ContentPresenter\s+x:Name="PART_SelectedContentHost".*?/>').Value
Require-Pattern `
    -Text $legacyContent `
    -Pattern 'HorizontalAlignment="Stretch"' `
    -Description "legacy content presenter horizontal stretch"
Require-Pattern `
    -Text $legacyContent `
    -Pattern 'VerticalAlignment="Stretch"' `
    -Description "legacy content presenter vertical stretch"

if ($failures.Count -gt 0) {
    $failures | ForEach-Object {
        Write-Error $_ -ErrorAction Continue
    }
    throw "Responsive review validation failed."
}

Write-Output "Responsive review validation passed."
