[CmdletBinding()]
param(
    [string]$SimulatorDirectory,
    [string]$ArtifactsDirectory
)

$ErrorActionPreference = "Stop"
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$repositoryRoot = [System.IO.Path]::GetFullPath(
    (Join-Path $scriptDirectory ".."))

if ([string]::IsNullOrWhiteSpace($SimulatorDirectory)) {
    $SimulatorDirectory = Join-Path `
        $repositoryRoot "artifacts\simulator\Release"
}
if ([string]::IsNullOrWhiteSpace($ArtifactsDirectory)) {
    $runId = [Guid]::NewGuid().ToString("N")
    $ArtifactsDirectory = Join-Path `
        $repositoryRoot ("artifacts\layout-probe\" + $runId)
}

$SimulatorDirectory = [System.IO.Path]::GetFullPath(
    $SimulatorDirectory)
$ArtifactsDirectory = [System.IO.Path]::GetFullPath(
    $ArtifactsDirectory)
$simulatorPath = Join-Path `
    $SimulatorDirectory "ClearPlan.Simulator.exe"
if (-not (Test-Path -LiteralPath $simulatorPath -PathType Leaf)) {
    throw "Simulator executable is missing: $simulatorPath"
}

New-Item `
    -ItemType Directory `
    -Path $ArtifactsDirectory `
    -Force | Out-Null

function Assert-Near {
    param(
        [double]$Expected,
        [double]$Actual,
        [string]$Description
    )

    if ([Math]::Abs($Expected - $Actual) -gt 1.0) {
        throw "$Description expected $Expected but found $Actual."
    }
}

function Invoke-LayoutProbe {
    param(
        [double]$Width,
        [double]$Height,
        [bool]$RequireHorizontalOverflow
    )

    $culture = [Globalization.CultureInfo]::InvariantCulture
    $widthText = $Width.ToString("0.##", $culture)
    $heightText = $Height.ToString("0.##", $culture)
    $sizeText = $widthText + "x" + $heightText
    $outputPath = Join-Path `
        $ArtifactsDirectory ("layout-" + $sizeText + ".json")

    $probeArguments = @(
        "--scenario",
        "field-and-mapping",
        "--layout-probe",
        $sizeText,
        "--output",
        $outputPath
    )
    $quotedArguments = foreach ($argument in $probeArguments) {
        if ($argument -match '\s|"') {
            '"' + $argument.Replace('"', '\"') + '"'
        }
        else {
            $argument
        }
    }

    $process = Start-Process `
        -FilePath $simulatorPath `
        -ArgumentList $quotedArguments `
        -PassThru
    if (-not $process.WaitForExit(30000)) {
        Stop-Process -Id $process.Id -Force
        throw "Layout probe timed out: $sizeText"
    }
    if ($process.ExitCode -ne 0) {
        throw "Layout probe failed with exit code $($process.ExitCode): $sizeText"
    }
    if (-not (Test-Path -LiteralPath $outputPath -PathType Leaf)) {
        throw "Layout probe did not create JSON: $outputPath"
    }

    $result = Get-Content -LiteralPath $outputPath -Raw |
        ConvertFrom-Json
    Assert-Near `
        -Expected $Width `
        -Actual ([double]$result.windowActualWidth) `
        -Description "Rendered window width"
    Assert-Near `
        -Expected $Height `
        -Actual ([double]$result.windowActualHeight) `
        -Description "Rendered window height"

    if (-not [bool]$result.workspaceFitsViewport) {
        throw "Workspace exceeds the rendered viewport: $sizeText"
    }
    if (-not [bool]$result.footerFullyVisible) {
        throw "Navigation footer is clipped: $sizeText"
    }
    if (-not [bool]$result.sourceStatusFullyVisible) {
        throw "Navigation source status is clipped: $sizeText"
    }
    if ([double]$result.sourceStatusHeight -le 0.0) {
        throw "Navigation source status was not rendered: $sizeText"
    }
    if (-not [bool]$result.dvhVerticalOverflowAvailable -or
        [double]$result.dvhScrollableHeight -le 0.0) {
        throw "DVH vertical overflow is unavailable: $sizeText"
    }
    if ($RequireHorizontalOverflow -and
        (-not [bool]$result.dvhHorizontalOverflowAvailable -or
         [double]$result.dvhScrollableWidth -le 0.0)) {
        throw "DVH horizontal overflow is unavailable: $sizeText"
    }

    return $result
}

$minimum = Invoke-LayoutProbe `
    -Width 1180.0 `
    -Height 720.0 `
    -RequireHorizontalOverflow $true
$commonLaptop = Invoke-LayoutProbe `
    -Width 1256.72 `
    -Height 720.0 `
    -RequireHorizontalOverflow $false

Write-Host "PASS responsive WPF layout probe"
Write-Host "Artifacts: $ArtifactsDirectory"
Write-Host (
    ("1180x720 workspace={0:0.##} footerBottom={1:0.##} " +
     "DVH scroll={2:0.##}x{3:0.##}") -f
        [double]$minimum.workspaceViewportHeight,
        [double]$minimum.footerBottom,
        [double]$minimum.dvhScrollableWidth,
        [double]$minimum.dvhScrollableHeight)
Write-Host (
    ("1256.72x720 workspace={0:0.##} footerBottom={1:0.##} " +
     "DVH scroll={2:0.##}x{3:0.##}") -f
        [double]$commonLaptop.workspaceViewportHeight,
        [double]$commonLaptop.footerBottom,
        [double]$commonLaptop.dvhScrollableWidth,
        [double]$commonLaptop.dvhScrollableHeight)
