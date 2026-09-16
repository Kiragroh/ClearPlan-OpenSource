# RTPLAN adapter dependencies

[Documentation](index.md) · [Configuration guide](configuration-guide.md)

`ClearPlan.Dicom` targets .NET Framework 4.8 and is independent of ESAPI. It uses
the official [fo-dicom 5.2.6 NuGet package](https://www.nuget.org/packages/fo-dicom/5.2.6),
not a custom DICOM binary parser. Native image codecs are not needed or bundled.
The package repository commit is `fb8beee64a38290160c4a966a630a6724a5d54d7`.

## Reproducible restore

Run `tools/build-and-test-dicom.ps1 -Configuration Release -Platform x64` using
Visual Studio MSBuild, a current .NET SDK, and the .NET Framework 4.8 targeting
pack. The full `tools/build-and-test.ps1` pipeline runs this step before building
the legacy solution. The two checked-in `packages.lock.json` files pin exact
package versions and content hashes. Locked mode rejects dependency drift.
Restore uses the official NuGet v3 source; dependencies are not checked into Git.

## Inventory

Versions and license expressions below were checked against the restored package
metadata. The `System.ValueTuple` package includes its MIT `LICENSE.TXT`.

| Package | Version | License |
| --- | --- | --- |
| fo-dicom | 5.2.6 | MS-PL |
| CommunityToolkit.HighPerformance | 8.4.0 | MIT |
| Microsoft.Bcl.AsyncInterfaces | 8.0.0 | MIT |
| Microsoft.Bcl.HashCode | 6.0.0 | MIT |
| Microsoft.Extensions.Configuration.Abstractions | 8.0.0 | MIT |
| Microsoft.Extensions.Configuration.Binder | 8.0.2 | MIT |
| Microsoft.Extensions.DependencyInjection | 8.0.1 | MIT |
| Microsoft.Extensions.DependencyInjection.Abstractions | 8.0.2 | MIT |
| Microsoft.Extensions.Logging | 8.0.1 | MIT |
| Microsoft.Extensions.Logging.Abstractions | 8.0.2 | MIT |
| Microsoft.Extensions.Options | 8.0.2 | MIT |
| Microsoft.Extensions.Options.ConfigurationExtensions | 8.0.0 | MIT |
| Microsoft.Extensions.Primitives | 8.0.0 | MIT |
| System.Buffers | 4.6.1 | MIT |
| System.Diagnostics.DiagnosticSource | 8.0.1 | MIT |
| System.Memory | 4.6.0 | MIT |
| System.Numerics.Vectors | 4.6.0 | MIT |
| System.Runtime.CompilerServices.Unsafe | 6.1.0 | MIT |
| System.Text.Encoding.CodePages | 8.0.0 | MIT |
| System.Text.Encodings.Web | 8.0.0 | MIT |
| System.Text.Json | 8.0.6 | MIT |
| System.Threading.Channels | 8.0.0 | MIT |
| System.Threading.Tasks.Extensions | 4.6.0 | MIT |
| System.ValueTuple | 4.5.0 | MIT |

The existing `Newtonsoft.Json` reference and `ClearPlan.Core` project are reused.
See [third-party notices](../THIRD-PARTY-NOTICES.md) and the unmodified upstream
[fo-dicom license/notices](../licenses/fo-dicom-5.2.6-LICENSE.txt), retained in full
even though optional image codec packages are not installed.

## Runtime packaging and host compatibility

The library build copies the complete managed runtime dependency set into
`artifacts/dicom/<Configuration>`. Bundle the complete DLL set, not only
`ClearPlan.Dicom.dll` and `fo-dicom.core.dll`. Preserve the existing licensed
vendor assembly handling; the adapter itself has no vendor assembly reference.
Include this inventory, `THIRD-PARTY-NOTICES.md`, and the `licenses` directory in
binary distributions that contain the adapter.

Standalone .NET Framework executables require the binding redirects generated
in `artifacts/dicom-tests/<Configuration>/ClearPlan.Dicom.Tests.exe.config`.
An Eclipse plug-in DLL's adjacent `.config` is not automatically the host's
configuration. Assembly resolution and coexistence with already-loaded package
versions therefore require a separate live-host smoke test. Synthetic byte-file
tests prove neither that host compatibility nor clinical correctness. Do not
change the clinical host's configuration or permissions automatically.

The adapter initializes fo-dicom without logging providers and does not create
DICOM network services. No external request is made when reading a local plan.
Package retrieval occurs only during development/build restore.

The NuGet vulnerability query `dotnet list ClearPlan.Dicom/ClearPlan.Dicom.csproj
package --vulnerable --include-transitive --source https://api.nuget.org/v3/index.json`
reported no known vulnerable packages on 2026-09-07. This is a point-in-time feed
check, not a security certification; rerun it for each release.
