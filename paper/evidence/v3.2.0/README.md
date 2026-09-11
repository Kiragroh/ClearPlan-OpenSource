# v3.2.0 publication provenance

The evaluated software revision is
`ca3e381813489542bae4551db9db23c8c3da23d8`. Subsequent publication-only commits
add the manuscript, figures, synthetic evidence, and document-generation tools;
they do not change application calculations. The v3.2.0 tag identifies that
combined source/publication tree. The release verification asset records the
actual remote tag and downloaded asset hashes after publication.

`publication-manifest.json` binds the publication inputs and outputs below by
SHA-256. It does not claim independent clinical validation. The older
`../technical_note_evidence.json` is the actual clean-software execution receipt
at the evaluated revision. Its collection-time unpublished status is preserved.
The final release additionally carries a fresh execution receipt from the tagged
tree. Both receipts identify tested inputs, timestamps, commands, and suite
output hashes. Native test records and private source data are not included.

## Scope and environment

The two raw JSON snapshots are generated directly from the compiled synthetic
factory; there is no patient or DICOM input. JSON deliberately omits detached
rendering geometry: regenerate image panels from the compiled factory, not by
treating the JSON files as clinical-image exports. Both scenarios use the same
analytical dose; different illustrative apertures do not calculate that dose.
Default report selection includes PTV_60, CTV_60, SpinalCord, Parotid_L, and
Parotid_R. External supplies whole-plan isodose-volume metrics but is unselected.
All optional CP0 BEV pages were included. No comparison plan is part of S1.

Generation environment: Windows 11 Enterprise 64-bit, 10.0.22631; .NET Framework
4.8 project targets; MSBuild 18.5.4.18101; Python 3.12.14 with openpyxl 3.1.5,
python-docx 1.2.0, and pypdf 6.10.0; Node.js 24.19.0. Word's local PDF renderer
was used for the manuscript. The application generated its own PDF/HTML report.
GUI captures are 1600 x 1000 pixels; the standalone DVH is 1400 x 800 pixels.

The Core suite passed 385 tests without skips. Separate DICOM and illustrative
RayStation suites each passed 20 tests. Nine stock-parser tests passed. The
collector records the unavailable private transcription class as one skipped
group; direct pytest collection reports its ten parametrized cases as skipped.
Those are different skip accounting levels, not ten additional passed tests.
The RayStation tests use fake contexts, not a commissioned RayStation runtime.

## Reproduction commands

From the tagged checkout on Windows, using an approved execution policy:

```powershell
.\tools\build-public-release.ps1 -Version 3.2.0 -PythonExecutable python
.\tools\export-publication-snapshots.ps1 `
  -SimulatorDirectory artifacts\simulator\Release `
  -OutputDirectory artifacts\publication\raw
.\artifacts\simulator\Release\ClearPlan.Simulator.exe `
  --scenario publication-single-layer --capture-all artifacts\publication\single
.\artifacts\simulator\Release\ClearPlan.Simulator.exe `
  --scenario publication-dual-layer --capture-all artifacts\publication\dual
node tools\paper_figures\generate_architecture.cjs artifacts\publication\figures
node tools\paper_figures\compose_workspace.cjs `
  --capture-dir artifacts\publication\dual --output-dir artifacts\publication\figures
python tools\build-technical-note-docx.py `
  --source paper\ClearPlan_ZMP_short_communication.md `
  --output artifacts\publication\ClearPlan_ZMP_technical_note.docx
.\tools\render-manuscript-word.ps1 `
  -Source artifacts\publication\ClearPlan_ZMP_technical_note.docx `
  -Pdf artifacts\publication\ClearPlan_ZMP_technical_note.pdf
```

Use fresh output directories. Node figure tools require the documented browser
dependencies. `--capture-all` also creates the synthetic PDF, HTML, and isolated
DVH export. Raw factory snapshots were generated twice and matched byte-for-byte.
Report/layout reproduction may depend on installed fonts and rendering versions;
the distributed hashes identify the exact inspected artifacts, not a promise
that PDF metadata will be identical on another workstation.

## Inspection record

Both 11-page synthetic reports, all 22 GUI/DVH PNGs, both embedded-resource HTML
quicklooks, and both figures received separate visual/source checks. Mapping
states, target-metric availability, the isolated DVH legend/notice, and HTML
contour/isodose legends were specifically rechecked. No unresolved material
artifact defect was found. The final 14-page manuscript PDF was rasterized and
all pages inspected after the baseline/fault wording correction; tables, math,
figures, captions, and references are readable. The editable DOCX builder passed
19 regression tests. These are software/editorial checks, not clinical approval.

The numeric claim audit found one baseline-count wording overstatement, now
corrected in abstract, methods, and results. Source release availability remains
a separate gate: the release's remote-readback record, not an earlier draft
review or execution receipt, establishes actual publication. Coauthor approval
and current journal portal requirements remain author submission actions.
