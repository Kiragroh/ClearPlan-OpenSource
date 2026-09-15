# ClearPlan

ClearPlan is an open-source .NET Framework 4.8/ESAPI framework for read-only radiotherapy plan review. It combines plan quality metric (PQM) evaluation, the existing PlanCheck surface, dose-volume histogram visualization, structured reporting, failure-mode logging, and a PlanFieldNamer-compatible field-name preview.

The public package is an institution-neutral baseline. Clinical constraint sets, paths, identifiers, and acceptance criteria remain local and require independent commissioning.

## Current development source

The current tree contains the `3.2.0-dev.20260915.7` review-workspace update after
the archived v3.2.0 release. It adds configurable isodose levels with one shared
GUI legend, compact field/report layouts, optional cached geometry playback and
separate gantry/couch orientation views. It is **development source, not a new
clinically commissioned release**. Cite the exact Git commit when reproducing
this state; the v3.2.0 release remains an earlier frozen artifact.

No clinical manuscript images, patient reports, local device measurements,
private clinical rules or deployment settings are supplied. Public collision
catalogs are intentionally empty. See the [development notes](docs/releases/2026-09-16-development-source.md)
and [receiving-clinic guide](docs/CLINIC_ONBOARDING.md).

For a vendor-free source verification on Windows with MSBuild/.NET Framework 4.8:

```powershell
.\tools\test-portable-review.ps1
```

## What is included

This tree contains the **v3.2.0 software and Technical Note materials**. It adds
target-specific PAM with physical single-/dual-layer MLC geometry, Paddick CI and
its reciprocal, GI/HI, total MU, explicitly estimated dose-rate trajectories,
plan comparison, shared structure visibility and a TG-263-oriented alias editor.
The active-plan PDF and HTML quicklook include compact findings, three CT planes
with dose/structure overlays and optional field-start BEV/DRR panels.
See the [analysis method](docs/plan-analysis-method.md),
[dose-rate limitations](docs/DOSE_RATE_ESTIMATE.md),
[review controls](docs/review-controls/README.md) and
[alias configuration](docs/STRUCTURE_ALIASES.md).
Local clinical commissioning remains necessary; neither a version label nor a
passing software test constitutes treatment approval. The Technical Note is an
author-review manuscript, not an accepted publication. See the
[paper and synthetic supplement](paper/README.md) for current files and evidence.

- `ClearPlan.Core`: ESAPI-independent constraint, settings, alias, table-selection, and field-naming logic
- `ClearPlan.Presentation`: vendor-neutral shared WPF review workspace and DVH presentation
- RefDB JSON and Excel XLSX adapters that populate one normalized internal catalog
- automatic RefDB-first loading with a visible Excel fallback
- deterministic alias and left/right laterality handling
- plan-context table selection with explicit confirmation for ambiguous results
- a read-only, plan-aware field preview with separate ID/name status and explicit angle-table-UZ/GUZ ordering
- a Clinical Blueprint interface with one integrated plan/PQM/PlanCheck/field/DVH review surface, a direct PDF report action, and complete detail/settings pages
- a standalone vendor-free simulator with seven checked-in deterministic scenarios plus two generated publication phantoms, wide interactive DVHs, watermarked reports and reproducible captures
- a versioned neutral review-snapshot contract shared by the simulator and the read-only Eclipse host
- an illustrative, unvalidated RayStation snapshot adapter that withholds clinical pass/fail classifications unless semantic equivalence has been commissioned
- editable settings for paths, aliases, constraints, field-name defaults, target rules and enabled checks, with author/time/hash history and restore-as-new-revision
- an editable Stock 2024 workbook with eight fractionation tables, 495 OAR rules, 70 canonical structures, a review-only audit queue and traceable sources; the small original example workbook is retained
- executable C# core/contract tests, separate synthetic DICOM and illustrative RayStation tests, stock parser/transcription checks, simulator smoke tests and release validators; executed counts and skips are recorded from each actual run
- optional, explicitly configured ARIA report-document upload with patient/provider binding and audit receipts; disabled by default and separate from read-only treatment-plan inspection

The open-source starter PlanCheck contains general examples. A clinical deployment can preserve its own complete `ErrorCalculator.cs`; the supplied deployment tooling explicitly excludes that file from replacement.

## Quick start: patient-free simulator

For a downloaded simulator ZIP, extract it into a local folder and start
`ClearPlan.Simulator.exe` from that folder. Keep its accompanying DLLs and
`Scenarios` directory together. Windows and .NET Framework 4.8 are required.
The executable is a synthetic demonstration, not a clinical TPS connection.

The simulator does not require Eclipse, ESAPI, a patient, or proprietary TPS
assemblies. Build and start the integrated publication scenario:

```powershell
$msbuild = & .\tools\resolve-msbuild.ps1
& $msbuild .\ClearPlan.sln /t:ClearPlan_Simulator /p:Configuration=Release /p:Platform=x64
.\artifacts\simulator\Release\ClearPlan.Simulator.exe --scenario publication-single-layer
```

Use `publication-dual-layer` for the jawless dual-layer example. Both publication
fixtures use the same mathematical phantom definitions for CT overlays, DVHs,
goals and target-quality metrics. Their dose is analytic, **not calculated from
the MLC apertures**; this is not physical dosimetric validation.

The original scenario IDs are `mixed-review`, `baseline-pass`, `target-underdose`,
`oar-overdose`, `metadata-plancheck`, `field-and-mapping`, and
`optional-path-fallback`. Every simulator window, capture, and report is marked
`SYNTHETIC DEMONSTRATION — NOT FOR CLINICAL USE`.

The public build requires Windows, .NET Framework 4.8 developer tools and MSBuild.
The release checks additionally use Python with `openpyxl`; the separate optional
DICOM project uses an SDK-style locked NuGet restore. No proprietary TPS assembly
is required for the simulator or these synthetic tests.

Run the complete deterministic UI/report smoke test with:

```powershell
powershell -NoProfile -File tools\test-simulator.ps1
```

## Quick start: Eclipse/ESAPI mode

This is a separate local build mode. The public release does not redistribute
licensed Varian assemblies or a preconfigured clinical ruleset.

1. Add institution-approved copies of `VMS.TPS.Common.Model.API.dll` and
   `VMS.TPS.Common.Model.Types.dll` to `ClearPlan-DLLs`.
2. Build the Eclipse workstation output:

   ```powershell
   powershell -NoProfile -File tools\build-and-test.ps1 `
     -Configuration Debug -Platform x64 -SkipPaperArtifacts
   ```

3. Open `built\settings.ini` or use the `Einstellungen` page to edit paths.
4. Select a source mode, enter source and output paths, and choose
   `Quelle testen`. Non-path options remain in `built\settings.json`.
5. Save only after the catalog has validated successfully.
6. Start `built\ClearPlan.Runner.exe` or load the built plugin in an
   ESAPI-capable environment.

For patient-free screenshots inside the Eclipse/ESAPI window, select the
session-only `Demodaten` switch in the ClearPlan header. The shared workspace
then hides the clinical plan and constraint context and displays the
watermarked deterministic `mixed-review` snapshot. Switching it off rebuilds
the current read-only clinical snapshot; the setting is never persisted.

Every file-system path is stored in `settings.ini`. Relative paths are resolved from the directory containing the active settings files, normally `built`. Absolute local or UNC paths are supported for clinical use. If an optional log, report, export, or state path is temporarily unavailable, ClearPlan uses a local directory under `built` instead of terminating the clinical UI.

## Constraint source modes

`constraintSource.mode` supports:

- `Automatic`: load RefDB when it is configured, readable, and valid; otherwise load Excel and show the fallback
- `RefDb`: load only RefDB JSON; a bad or missing path is a blocking error
- `Excel`: load only the Excel workbook; a bad or missing path is a blocking error

Relevant non-path option:

```json
{
  "constraintSource": {
    "mode": "Automatic",
    "includeInactiveTables": false
  }
}
```

The `[Paths]` section in `settings.ini` contains RefDB, Excel, templates, logs, reports, exports, state, auxiliary files, and optional clinical support paths. The settings page exposes all of them together with the non-path options and privacy switches. `Quelle testen` validates paths and catalog content without replacing the active catalog. `Speichern` validates first and writes both settings files atomically. `Neu laden` reloads settings and source status.

## Public Excel schema

The default fallback is `ClearPlan.Script\Distribution\ConstraintTemplates\ClearPlan_StockConstraints2024.xlsx`. It contains eight current 2024 compilation tables (1/3/5/8/10/15/20 fractions and conventional), and safe TG-263-oriented glossary aliases. The three machine-readable sheets are:

- `Structures`: canonical structure identity, code, DICOM type, laterality, aliases, and active status
- `Tables`: table identity, display name, plan/PlanSum compatibility, fraction and dose ranges, site/regime hints, and active status
- `Constraints`: relationship to table and structure, PQM objective, comparator, numeric value, unit, goal, variation, and active status

Keep identifiers unique and relationships explicit. The validator rejects missing headers, duplicate identifiers, broken table/structure references, and invalid objective values. Do not merge cells or add presentation-only rows inside the data ranges.

Stock tables have `requires_confirmation=true`: compatibility is a suggestion, not automatic clinical approval. The 318 non-executable source rules and 42 excluded alias cases stay visible in `ReviewRows`; `Sources` preserves reference codes and provenance. No institution-specific prescription links, target placeholders, cropped-organ equivalences or new acceptance thresholds are invented. Configure a local copy or use RefDB. See [stock maintenance and verification](tools/stock-catalog/README.md).

The smaller original `ClearPlan_DefaultConstraints.xlsx` remains an example fixture; its existing validator is:

```powershell
python tools\validate-constraint-workbook.py
```

## Migrating legacy CSV files

CSV files are no longer active runtime defaults. To migrate an existing set:

```powershell
python tools\convert-constraint-csvs-to-xlsx.py `
  --input-directory <legacy-csv-folder> `
  --output <constraints.xlsx>
```

Then verify record-level equivalence:

```powershell
python tools\verify-constraint-migration.py `
  --input-directory <legacy-csv-folder> `
  --workbook <constraints.xlsx>
```

Review all generated identifiers, table metadata, units, aliases, and laterality before clinical use. Retain the original files as rollback evidence until commissioning is complete.

## Adding aliases safely

- Prefer one canonical structure identifier per concept.
- Enter aliases in the source data, not in application code.
- Encode laterality explicitly for paired organs.
- Do not make one unqualified alias resolve to both left and right structures.
- Treat ambiguity as a validation/review item instead of relying on row order.
- Add or update core tests when introducing a new matching rule.

## Field-name preview

The field-name page compares existing field IDs and names with read-only expected values but never modifies the plan. A valid current field prefix is preserved; otherwise the leading alphanumeric part of the plan ID is used. For example, plan `7GA_HI9-15_20` expects `7GA01`, `7GA02`, and so on. Name examples from the test suite:

- static gantry 90.2° → `90`
- clockwise arc 179° to 181° → `179-181 UZ`
- counter-clockwise arc 181° to 179° → `181-179 GUZ`
- clockwise arc with couch 30° → `179-181 T30 UZ`
- duplicate clockwise arcs → `179-181 UZa`, `179-181 UZb`

The adapter contains no plan-modification, beam-renaming, setup-field, apply, or save calls.

## Integrated review surface

`Gesamtansicht` shows the available plans, the complete PQM result list with the resolved structure visible in every row, PlanCheck findings, separate field-ID/name conformance, and the DVH in one vertically scrollable review. The focused pages remain available for PQM reassignment, reporting options, reference-point details, the full field comparison, DVH controls, and settings. DVH selections are shared through view-model state; overall and detail plots use separate synchronized OxyPlot models. The focused DVH page uses a bounded, resizable structure selector and gives the remaining finite width and height directly to the plot.

## Essentials runner presentation

The desktop executable continues to use `EsapiEssentials.PluginRunner` for patient/plan opening and the ESAPI lifecycle. Runner-only WPF resources give that existing window the Clinical Blueprint appearance. ClearPlan does not replace or reflect into the Essentials data model or commands; presentation failure is non-blocking and leaves the original runner usable.

## Illustrative RayStation adapter

`examples\raystation\clearplan_review_snapshot.py` demonstrates how a second
scripting-capable TPS can populate the neutral snapshot boundary without
writing to a treatment plan. It is intentionally conservative: unsupported
semantics remain `not-evaluated`, output must be an explicit local path, and
the example is not runtime-validated or clinically commissioned in RayStation.
See `examples\raystation\README.md` and run its tests with:

```powershell
python -m unittest discover -s examples\raystation\tests -v
```

## Privacy defaults

The public settings avoid direct patient and user identifiers in automatic logs and generated filenames. If a clinic enables direct identifiers, it must do so under local governance. Figures, bug reports, example workbooks, and pull requests must remain patient-free.

## Repository layout

- `ClearPlan.Core`: source-independent logic
- `ClearPlan.Core.Tests`: executable tests
- `ClearPlan.Presentation`: shared snapshot-backed WPF presentation
- `ClearPlan.Rendering`: shared CT/BEV image rendering
- `ClearPlan.Dicom` and `ClearPlan.Dicom.Tests`: optional read-only RTPLAN enrichment and synthetic-file tests; not required by the simulator
- `ClearPlan.Script`: ESAPI plugin, Clinical Blueprint UI, PlanCheck, and adapters
- `ClearPlan.Reporting`: report data structures
- `ClearPlan.Reporting.MigraDoc`: PDF and offline HTML quicklook rendering
- `ClearPlan.Runner`: desktop launcher
- `paper`: technical note and reproducible public-safe figures
- `tools`: build, migration, release, PlanCheck guard, and deployment utilities

Distribution files are maintained in `ClearPlan.Script\Distribution` and copied to `built` during compilation.

## Validation

For a locally configured, licensed ESAPI development environment:

```powershell
powershell -NoProfile -File tools\build-and-test.ps1 -SkipPaperArtifacts
```

It builds the requested `Debug|Release` x64 configuration, runs the C# suite
and Python adapter tests, checks a Python-produced snapshot against the
C# contract validator, validates the workbook, scans the source and Office XML
for private tokens, checks the shared workspace and network fallback
contracts, and verifies the read-only field and adapter boundaries.

For a clean vendor-free checkout, build and package only the public simulator
and its tests:

```powershell
powershell -NoProfile -File tools\build-public-release.ps1
# If Python is not on PATH, add: -PythonExecutable 'C:\path\python.exe'
```

The public workflow builds Core/Simulator and the separate locked DICOM project,
then executes the C#, DICOM, RayStation and stock suites. Its current run receipt
is `artifacts/public-release-test-evidence.json`, with actual passed counts,
skipped tests/groups, timestamps and tested-input hashes. A clean public checkout
does not contain the private stock-transcription intermediate: that test group
is explicitly skipped, while parser tests and the distributed workbook's C#
loader regression still execute. A skipped group is not counted as passed.

The workflow also creates the ZIP twice and requires byte-identical SHA-256
hashes. This proves deterministic packaging of one build, not reproducibility
of independent clean builds. Final manuscript/figure review and remote release
verification are separate gates. `-DevelopmentPreview` leaves manuscript gates
pending; `-SkipSimulatorSmoke` explicitly skips UI/report verification. Neither
option completes release verification. See `docs\reproducibility.md`.

To refresh software evidence without changing the manuscript, after building:

```powershell
python tools\collect-paper-evidence.py --configuration Release --include-dicom `
  --release-version 3.2.0 --output artifacts\technical-note-evidence.json
```

This collector executes the current suites. Test registrations are never used
as a substitute for execution, and collected evidence does not claim that a tag,
published asset or manuscript has been verified.

Individual guards:

```powershell
powershell -NoProfile -File tools\validate-public-release.ps1
powershell -NoProfile -File tools\validate-ui-distinctiveness.ps1
powershell -NoProfile -File tools\validate-readonly-field-preview.ps1
```

Use a trusted checkout under your institution's execution/signature policy; do not disable that policy to run these commands. These checks are implementation safeguards, not a substitute for clinical commissioning.

## Community and publication

- Use GitHub Issues for anonymized bugs, onboarding problems, and feature requests.
- See `CONTRIBUTING.md` before changing sources, aliases, PlanCheck, or ESAPI adapters.
- The author-review Technical Note and exclusively synthetic supplement are
  indexed in [paper/README.md](paper/README.md).
- A proposed, nonuniversal local transfer checklist is in
  `docs\local-commissioning-scaffold.md`.
- Versioned software, downloadable artifacts, and publication/readback manifests:
  [v3.2.0](https://github.com/Kiragroh/ClearPlan-OpenSource/releases/tag/v3.2.0).
  Test execution, public artifact verification, and local clinical acceptance are
  separate evidence scopes; the software collector does not verify publication.

ClearPlan is a framework; each institution remains responsible for its clinical rules, validation, deployment, and governance.
