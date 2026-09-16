# Reproducibility and evidence scopes

[Documentation](index.md) · [Configuration guide](configuration-guide.md)

The archived v3.2.0 release and its paper evidence describe an earlier snapshot.
For current development-source verification, start with the
[developer handoff](developer-handoff.md). The packaging workflow below is retained
for release engineering; a successful run is not a new release or clinical
acceptance. Do not relabel archived v3.2.0 artifacts with development bytes.

## Requirements and environment provenance

The vendor-free build uses Windows x64, .NET Framework 4.8 developer tools and
MSBuild. Python with `openpyxl` supports workbook and adapter tests. The optional
DICOM project additionally requires SDK-style project support and its locked
NuGet dependencies. Eclipse, a patient database and licensed `VMS.*` assemblies
are not needed for the public Core, Simulator or synthetic DICOM tests.

Use an institution-approved script execution policy and tool installation. These
instructions do not require changing that policy or disabling endpoint controls.
The default Python command is `python`; supply its executable path explicitly
when it is not on PATH. No bundled private development runtime is required.

Record the actual OS, .NET runtime/targeting pack, MSBuild, PowerShell, Python,
`openpyxl`, Git and dependency versions with each evaluated build. Do not reuse
an earlier workstation inventory as evidence for another run. The collector's
UTC values are execution timestamps, not independent validation of workstation
or server clock accuracy; a deterministic fixture timestamp is not a test date.

## Historical v3.2.0 packaging workflow

For reproducing the reviewed v3.2.0 source checkout, the workflow is:

```powershell
$clearPlanPython = 'python'
.\tools\build-public-release.ps1 -Version 3.2.0 `
  -PythonExecutable $clearPlanPython -DevelopmentPreview
```

`-DevelopmentPreview` permits candidate work while manuscript gates remain
pending. For the final release gate, rerun without it. Do not use
`-SkipSimulatorSmoke` for final verification; that switch explicitly leaves
UI/report checks unexecuted.

The entry point builds only the public Core/Simulator targets and the separate
locked DICOM project. It runs the suites below, validates a Python-produced
snapshot with the C# contract validator, checks public-tree/package inputs and
vendor-free assemblies, checks version identity, exercises simulator smoke
capture/report paths and packages the same build twice.

| Evidence scope | Actual execution | Boundary |
| --- | --- | --- |
| Core/contract | `ClearPlan.Core.Tests.exe` | Each registered group must have a matching successful execution row; nested subcase output does not inflate the group total. |
| Synthetic DICOM | `ClearPlan.Dicom.Tests.exe` after locked restore/build | Generated RTPLAN byte-file tests, not patient exports or a commissioned native adapter. |
| Illustrative RayStation | Python `unittest` plus Python-to-C# snapshot validation | Fake contexts; no proof of runtime compatibility or full metric equivalence in RayStation. |
| Stock catalog | Python parser/transcription tests and the C# distributed-workbook loader test | The private source-transcription class is skipped on a clean public checkout. Skipped classes/tests are recorded separately, never as passes. |

The current execution receipt is
`artifacts/public-release-test-evidence.json`. It records actual passed counts,
skipped tests/groups, execution times, tested-input hashes and output digests.
The active stock workbook and the smaller example workbook are inventoried
separately. Neither registered test counts nor workbook row counts are evidence
of clinical validity.

To refresh software evidence after building, without rewriting a manuscript:

```powershell
& $clearPlanPython .\tools\collect-paper-evidence.py --configuration Release `
  --release-version 3.2.0 --include-dicom `
  --output artifacts\technical-note-evidence.json
```

The collector reruns the suites; it does not infer their results from source
registration. Its output explicitly leaves manuscript, figure and remote-release
verification pending. Copying a receipt into the paper evidence directory is a
separate reviewed publication step, not a test execution.

## Deterministic scenarios and figure inputs

The seven checked-in baseline/fault scenarios are `baseline-pass`,
`target-underdose`, `oar-overdose`, `metadata-plancheck`, `field-and-mapping`,
`optional-path-fallback` and `mixed-review`. Their PlanCheck deviations, mapping
ambiguities and optional-source failures are declared synthetic inputs. Numeric
PQM values are calculated from their synthetic DVHs; field suggestions use the
tested naming implementation.

`publication-single-layer` and `publication-dual-layer` are two additional
generated fixtures exposed by the same simulator selection/CLI. One mathematical
ellipsoid phantom supplies their CT planes, contours, analytic dose, voxel-sampled
DVHs and target-quality calculations. Their DRRs use the same synthetic anatomy.
The distinct MLC apertures exercise single-layer or jawless dual-layer geometry
and PAM; the estimated dose rate uses explicit synthetic profiles.

**The analytic phantom dose is not calculated from the MLC apertures.** This
consistency demonstration is not validation of dose calculation or machine
delivery. The original seven scenarios retain their separately labeled image
fixtures; do not infer anatomical correspondence from their adjacent panels.

For example, generate a publication-fixture capture/report set with:

```powershell
.\artifacts\simulator\Release\ClearPlan.Simulator.exe `
  --scenario publication-single-layer --capture-all artifacts\publication-single-layer
.\artifacts\simulator\Release\ClearPlan.Simulator.exe `
  --scenario publication-dual-layer --capture-all artifacts\publication-dual-layer
```

Capture success is not visual QA. Inspect every intended GUI image and PDF/HTML
page for labels, dose/structure overlays, clipping, units, legends and privacy.
The compiled synthetic factory can also be exported for numerical inspection:

```powershell
.\tools\export-publication-snapshots.ps1 -SimulatorDirectory artifacts\simulator\Release `
  -OutputDirectory artifacts\publication-snapshots
python tools\build-paper-figures.py --capture-dir artifacts\publication-dual-layer `
  --output-dir artifacts\paper-figures
```

Figure generation requires Node.js, Playwright and its Chromium browser. It uses
only local embedded captures and an original SVG; no clinical image or remote
asset is loaded. Pass `--node-executable` for an explicit runtime path. Exported
JSON omits rendering-only geometry; it is not a replacement for image inspection.
The seven JSON fixture files remain separately enumerated package inputs; the
two publication fixtures are generated by compiled public source, not additional
patient-derived JSON files.

## Checked source to released artifact

Use the final reviewed source/artifact manifests and execution evidence for
commit IDs, file inventories, environment details and SHA-256 values; do not
substitute an older tag or a private development commit. Before publishing:

1. Freeze the reviewed public source and configuration inventory. Record each
   allowed relative path and exact-byte SHA-256, including original license and
   third-party notices. Keep native captures, site-specific rules and credentials out.
2. Build from that checkout and retain its commands, logs and execution receipt.
   A dirty checkout requires an explicit manifest of the evaluated changes;
   `repositoryCommit` alone does not identify those working-tree bytes.
3. Verify that packaged files match the allowlisted build outputs and reviewed
   synthetic inputs. The simulator package must not acquire licensed ESAPI or
   optional DICOM dependencies merely because they exist in a development folder.
4. Inspect publication artifacts separately, including metadata and embedded
   resources. Link requirements to the final paper traceability/evidence records.
5. Read back the newly published tag/commit and download the release assets.
   Compare their inventory and hashes against the approved manifests. A release
   URL in README or CFF does not establish that this step occurred.

The ZIP writer uses sorted entries and a fixed archive timestamp. Matching two
archives created from one build demonstrates deterministic **packaging**, not
independent clean-build reproducibility. Claim the latter only after a separately
recorded clean build and an explicitly defined output comparison.

## Native, snapshot and publication boundaries

Eclipse inspection is a separate licensed, read-only workflow. Its detached
snapshots, dosimetry, CT/BEV images and reports remain confidential even if a
display name is hidden. A neutral schema is not an anonymization guarantee.
Public examples may contain synthetic fixtures, original diagrams, or patient-free
configuration screens—not native cases with identifying labels removed. Clinical
publication material requires separate institutional review and is not part of
this public development-source update.

Optional [ARIA document upload](ARIA_REPORT_UPLOAD.md) is an explicitly confirmed
write to the patient record, separate from read-only treatment-plan access.
Sparse FHIR readback can be verified against stored PDF bytes only through
explicitly configured, restricted UNC roots; an empty root list disables this
filesystem route. Metadata-only readback does not prove that the uploaded PDF
was stored intact. This local interoperability path and its protected receipts
are not part of the public synthetic evidence set.

Software tests do not establish clinical sensitivity/specificity, dose-model
agreement, actual delivery behavior or site-specific commissioning. See the
[local commissioning scaffold](local-commissioning-scaffold.md) and the
[candidate release record](releases/v3.2.0.md).
