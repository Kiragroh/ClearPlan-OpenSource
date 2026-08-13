# ClearPlan v3.1.0 reproducibility

This workflow rebuilds and checks only the vendor-free public components. It
does not require Eclipse, ESAPI, a patient database, or the untracked
`VMS.TPS.*` assemblies used by the separate clinical build.

## Tested environment

- Windows x64 build 22631.7376 (23H2)
- .NET Framework 4.8 target pack; installed runtime release key `533320`
- MSBuild `18.5.4.18101`
- PowerShell `7.4.14`
- Python `3.13.13`
- Git for Windows `2.54.0.windows.1`
- Newtonsoft.Json assembly `13.0.0.0`, OxyPlot assemblies `2.0.0.0`,
  and PDFsharp/MigraDoc assemblies `1.50.5147.0`

The Release projects use deterministic C# compilation and omit debug symbols.
Scenario JSON is serialized with invariant numerical formatting. The package
writer stores sorted entries with a fixed ZIP timestamp, avoiding
runtime-specific Deflate output. The final archive was byte-identical when
packaged from Windows PowerShell 5.1 and PowerShell 7.4.14.

## Rebuild and verify

From the repository root of a checkout of tag `v3.1.0`, run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
  -File tools\build-public-release.ps1
```

The script builds the public test harness and simulator, runs 125 C# tests and
20 fake-context Python tests, exports one Python-produced adapter snapshot and
validates it with the C# contract validator, validates the workbook and public tree, performs
the simulator layout/capture/report smoke test, checks that the packaged seven
JSON files are byte-identical to the reviewed source scenarios, and creates
the public ZIP twice. The two archive hashes must match before the release ZIP
and its `.sha256` file are written to `artifacts\release`.

## Oracle and evidence boundaries

PQM values are calculated from sampled synthetic DVHs and independently
recomputed in the C# harness. PlanCheck deviations, mapping ambiguity, and
optional-source fallback are deliberate fixed scenario inputs. Field-name
suggestions are produced by the tested naming implementation. The
requirements-to-evidence links are listed in
`paper/TRACEABILITY_MATRIX.md`.

The workflow establishes deterministic behavior for the declared synthetic
cases in the documented environment. It does not establish clinical
sensitivity, specificity, runtime compatibility with RayStation, or
site-specific Eclipse commissioning.
