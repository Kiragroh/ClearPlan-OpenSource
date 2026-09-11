[CmdletBinding()]
param(
    [string]$AssemblyDirectory = ""
)

# Detached synthetic regression probe; no ESAPI access and no visible windows.
$ErrorActionPreference = "Stop"
if (-not $AssemblyDirectory) { $AssemblyDirectory = Join-Path $PSScriptRoot "..\artifacts\bin\Debug" }
if ([Threading.Thread]::CurrentThread.ApartmentState -ne "STA") {
    throw "Run this probe with powershell.exe -STA -File."
}
Push-Location (Join-Path $PSScriptRoot "..")
try {
    $assemblyRoot = [IO.Path]::GetFullPath($AssemblyDirectory)
    Add-Type -AssemblyName WindowsBase, PresentationFramework, PresentationCore
    $references = @("System.dll", "System.Core.dll",
        [Windows.Threading.Dispatcher].Assembly.Location,
        [Windows.Application].Assembly.Location,
        [Windows.Media.Visual].Assembly.Location)
    foreach ($name in @("Newtonsoft.Json.dll", "OxyPlot.dll", "ClearPlan.Core.dll", "ClearPlan.Rendering.dll", "ClearPlan.Presentation.dll")) {
        $path = Join-Path $assemblyRoot $name
        [void][Reflection.Assembly]::LoadFrom($path)
        $references += $path
    }
    $testAssembly = Add-Type -Path "ClearPlan.Core.Tests\TestAssert.cs", "ClearPlan.Core.Tests\ReviewDvhInteractionTests.cs" -ReferencedAssemblies $references -PassThru
    $testType = $testAssembly | Where-Object FullName -eq "ClearPlan.Core.Tests.ReviewDvhInteractionTests"
    $failures = 0
    foreach ($method in $testType.GetMethods([Reflection.BindingFlags]"Public,Static")) {
        try {
            $method.Invoke($null, @())
            Write-Host ("PASS DVH interaction: " + $method.Name)
        }
        catch {
            $failures++
            $failure = $_.Exception
            while ($failure.InnerException) { $failure = $failure.InnerException }
            Write-Host ("FAIL DVH interaction: " + $method.Name + " - " + $failure.Message)
        }
    }
    if ($failures -gt 0) { throw "$failures DVH interaction regression(s) failed." }
}
finally {
    Pop-Location
}
