# Developer handoff

[Documentation](index.md) · [Architecture](architecture.md) · [Configuration](configuration-guide.md) · [Contributing](../CONTRIBUTING.md)

Code baseline:
[`98aed7bc2fb3ecbced7cc2a520f691da36c53bfa`](https://github.com/Kiragroh/ClearPlan-OpenSource/tree/98aed7bc2fb3ecbced7cc2a520f691da36c53bfa),
version `3.2.0-dev.20260916.1`. Record the commit and configuration revisions actually
evaluated, not just a version string or an older release's results.

## Patient-free build and test

Requirements: Windows x64, MSBuild, .NET Framework 4.8 developer tools. From the
repository root, under an approved execution policy:

```powershell
.\tools\test-portable-review.ps1
.\artifacts\simulator\Debug\ClearPlan.Simulator.exe --scenario mixed-review
```

This builds portable test/simulator targets, runs the C# suite, and validates
public privacy/configuration. It does not execute licensed ESAPI integration.
Baseline verification ran 495 top-level tests with zero failures; rerun on your
checkout. Nested PASS messages do not increase the top-level count.

After copying files while preserving timestamps, force a rebuild to avoid stale binaries:

```powershell
$clearPlanMsbuild = & .\tools\resolve-msbuild.ps1
& $clearPlanMsbuild .\ClearPlan.sln `
  '/t:ClearPlan_Core_Tests:Rebuild;ClearPlan_Simulator:Rebuild' `
  /p:Configuration=Debug /p:Platform=x64
.\artifacts\bin\Debug\ClearPlan.Core.Tests.exe
.\tools\validate-public-release.ps1
git diff --check
```

Use solution targets: the solution maps reporting-project configurations correctly.
Do not rebuild over a running test process holding output DLLs open. Focused tests:

```powershell
.\artifacts\bin\Debug\ClearPlan.Core.Tests.exe ReviewLanguage
.\artifacts\bin\Debug\ClearPlan.Core.Tests.exe EsapiBev
.\artifacts\bin\Debug\ClearPlan.Core.Tests.exe ConfigurationHistory
python -m unittest discover -s examples\raystation\tests -v
```

Python tests use fake contexts. Workbook, synthetic DICOM, capture, and package
checks have additional dependencies; see [reproducibility](reproducibility.md).
Record skipped tests as skipped, not passed.

## Licensed Eclipse build

Supply approved compatible `VMS.TPS.Common.Model.API.dll` and
`VMS.TPS.Common.Model.Types.dll` locally in `ClearPlan-DLLs`. Do not redistribute them.

```powershell
.\tools\build-and-test.ps1 -Configuration Debug -Platform x64 -SkipPaperArtifacts
```

Run the plugin/runner only in an approved ESAPI-capable environment. Configure the
copied output settings and perform protected read-only acceptance. Preserve a
locally governed complete `ErrorCalculator.cs`; the public starter is not its
replacement. Freeze the exact local source/dependency closure.

## Where to extend

| Change | Location | Verify |
|---|---|---|
| Local goals/aliases | Catalogs and configuration validators | Units, scope, laterality, ambiguity, missing dose |
| Generic calculation | `ClearPlan.Core` | Independent synthetic cases and unavailable states |
| Native observation | `ClearPlan.Script/Review` | Owner thread, detached workers, cancellation, context freshness |
| Display label | Localization catalog and explicit display binding | Both languages; unchanged IDs, values, canonical status |
| Report layout | Report document and HTML/PDF renderers | Same-plan identity, visibility, pagination, optional-image failure |
| Another TPS | Snapshot contract / RayStation example | Explicit semantics, read-only API inventory, native acceptance |

Keep clinic-specific rules, identifiers, paths, profiles, and secrets out of generic
source. Prefer structured native observations to fuzzy message parsing. Add a failing
regression before a production fix.

## Open limits and handoff checks

- Public checks are examples, not the complete private PlanCheck engine.
- CT approval configuration currently adapts one documented legacy fallback tuple,
  not all CT checks; empty approvals do not pass.
- The RayStation example is unvalidated and not an equivalent full-feature port.
- PAM depends on physical single-/dual-layer geometry, target choice, numerical
  projection, and weighting; read the [method](plan-analysis-method.md).
- DRR is a display proxy, collision is sampled model review, and estimated dose rate
  is not delivery evidence.
- A cooperative timeout cannot interrupt a blocked vendor call. Only the BEV builder
  owns the optional report's 45-second budget; keep post-await context/cancellation guards.
- Portable source contracts do not replace native API tests or clinical commissioning.

Before handoff, record exact source/configuration, run focused and full relevant
tests, inspect patient-free GUI and PDF/HTML, audit public diffs and embedded
metadata, and state untested limitations and rollback. Capture success is not visual QA.

The Technical Note aims to describe configurable workflow integration and transparent
boundaries. Clinical detection performance, cross-vendor equivalence, and accepted
publication require separate evidence; do not infer them from software tests.
