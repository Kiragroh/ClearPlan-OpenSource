# ClearPlan

ClearPlan is an open-source .NET Framework 4.8/ESAPI framework for read-only radiotherapy plan review. It combines plan quality metric (PQM) evaluation, the existing PlanCheck surface, dose-volume histogram visualization, structured reporting, failure-mode logging, and a PlanFieldNamer-compatible field-name preview.

The public package is an institution-neutral baseline. Clinical constraint sets, paths, identifiers, and acceptance criteria remain local and require independent commissioning.

> **Independent project:** ClearPlan is maintained independently and is not an official publication of any employer, healthcare provider, university, or software vendor. It is supplied without warranty and is not a certified medical device. See [DISCLAIMER.md](DISCLAIMER.md).

## What is included

- `ClearPlan.Core`: ESAPI-independent constraint, settings, alias, table-selection, and field-naming logic
- `ClearPlan.Presentation`: vendor-neutral shared WPF review workspace and DVH presentation
- RefDB JSON and Excel XLSX adapters that populate one normalized internal catalog
- automatic RefDB-first loading with a visible Excel fallback
- deterministic alias and left/right laterality handling
- plan-context table selection with explicit confirmation for ambiguous results
- a read-only, plan-aware field preview with separate ID/name status and explicit angle-table-UZ/GUZ ordering
- a Clinical Blueprint interface with one integrated plan/PQM/PlanCheck/field/DVH review surface, a direct PDF report action, and complete detail/settings pages
- a standalone vendor-free simulator with seven checked-in deterministic synthetic scenarios, wide DVH review, watermarked reports, and reproducible captures
- a versioned neutral review-snapshot contract shared by the simulator and the read-only Eclipse host
- an illustrative, unvalidated RayStation snapshot adapter that withholds clinical pass/fail classifications unless semantic equivalence has been commissioned
- an editable settings page for source and operational paths
- a public workbook containing three example tables, 19 constraints, and seven structures
- 125 C# executable tests, 20 RayStation example tests, simulator smoke tests, and release validation tools

The open-source starter PlanCheck contains general examples. A clinical deployment can preserve its own complete `ErrorCalculator.cs`; the supplied deployment tooling explicitly excludes that file from replacement.

## Quick start: patient-free simulator

The simulator does not require Eclipse, ESAPI, a patient, or proprietary TPS
assemblies. Build and start the integrated publication scenario:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools\build-and-test.ps1 `
  -Configuration Release -Platform x64
artifacts\simulator\Release\ClearPlan.Simulator.exe --scenario mixed-review
```

The other scenario IDs are `baseline-pass`, `target-underdose`,
`oar-overdose`, `metadata-plancheck`, `field-and-mapping`, and
`optional-path-fallback`. Every simulator window, capture, and report is marked
`SYNTHETIC DEMONSTRATION — NOT FOR CLINICAL USE`.

Run the complete deterministic UI/report smoke test with:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools\test-simulator.ps1
```

## Quick start: Eclipse/ESAPI mode

This is a separate local build mode. The public release does not redistribute
licensed Varian assemblies or a preconfigured clinical ruleset.

1. Add institution-approved copies of `VMS.TPS.Common.Model.API.dll` and
   `VMS.TPS.Common.Model.Types.dll` to `ClearPlan-DLLs`.
2. Build the Eclipse workstation output:

   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File tools\build-and-test.ps1 `
     -Configuration Debug -Platform x64
   ```

3. Open `built\settings.ini` or use the `Einstellungen` page to edit paths.
4. Select a source mode, enter source and output paths, and choose
   `Quelle testen`. Non-path options remain in `built\settings.json`.
5. Save only after the catalog has validated successfully.
6. Start `built\ClearPlan.Runner.exe` or load the built plugin in an
   ESAPI-capable environment.

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

The workbook `ClearPlan.Script\Distribution\ConstraintTemplates\ClearPlan_DefaultConstraints.xlsx` has three sheets:

- `Structures`: canonical structure identity, code, DICOM type, laterality, aliases, and active status
- `Tables`: table identity, display name, plan/PlanSum compatibility, fraction and dose ranges, site/regime hints, and active status
- `Constraints`: relationship to table and structure, PQM objective, comparator, numeric value, unit, goal, variation, and active status

Keep identifiers unique and relationships explicit. The validator rejects missing headers, duplicate identifiers, broken table/structure references, and invalid objective values. Do not merge cells or add presentation-only rows inside the data ranges.

Validate the public workbook with:

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

`Gesamtansicht` shows the available plans, the complete PQM result list with the resolved structure visible in every row, PlanCheck findings, separate field-ID/name conformance, and the DVH in one vertically scrollable review. The focused pages remain available for PQM reassignment, reporting options, reference-point details, the full field comparison, DVH controls, and settings. DVH selections are shared through view-model state; overall and detail plots use separate synchronized OxyPlot models.

## Essentials runner presentation

The desktop executable continues to use `EsapiEssentials.PluginRunner` for patient/plan opening and the ESAPI lifecycle. Runner-only WPF resources give that existing window the Clinical Blueprint appearance. ClearPlan does not replace or reflect into the Essentials data model or commands; presentation failure is non-blocking and leaves the original runner usable.

## Illustrative RayStation adapter

`examples\raystation\clearplan_review_snapshot.py` demonstrates how a second
scripting-capable TPS can populate the neutral snapshot boundary without
writing to a treatment plan. It is intentionally conservative: unsupported
semantics remain `not-evaluated`, output must be an explicit local path, and
the example is not runtime-validated or clinically commissioned in RayStation.
See `examples\raystation\README.md` and run its 20 tests with:

```powershell
python -m unittest discover -s examples\raystation\tests -v
```

## Privacy defaults

The public settings avoid direct patient and user identifiers in automatic logs and generated filenames. If a clinic enables direct identifiers, it must do so under local governance. Figures, bug reports, example workbooks, and pull requests must remain patient-free.

## Repository layout

- `ClearPlan.Core`: source-independent logic
- `ClearPlan.Core.Tests`: executable tests
- `ClearPlan.Presentation`: shared snapshot-backed WPF presentation
- `ClearPlan.Script`: ESAPI plugin, Clinical Blueprint UI, PlanCheck, and adapters
- `ClearPlan.Reporting`: report data structures
- `ClearPlan.Reporting.MigraDoc`: PDF rendering
- `ClearPlan.Runner`: desktop launcher
- `paper`: technical note and reproducible public-safe figures
- `tools`: build, migration, release, PlanCheck guard, and deployment utilities

Distribution files are maintained in `ClearPlan.Script\Distribution` and copied to `built` during compilation.

## Validation

Run the full local pipeline:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools\build-and-test.ps1
```

It builds the requested `Debug|Release` x64 configuration, runs 125 C# tests
and the 20 Python adapter tests, checks a Python-produced snapshot against the
C# contract validator, validates the workbook, scans the source and Office XML
for private tokens, checks the shared workspace and network fallback
contracts, and verifies the read-only field and adapter boundaries.

For a clean vendor-free checkout, build and package only the public simulator
and its tests:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools\build-public-release.ps1
```

This public workflow also creates the ZIP twice and requires byte-identical
SHA-256 hashes. See `docs\reproducibility.md`.

Individual guards:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools\validate-public-release.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tools\validate-ui-distinctiveness.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tools\validate-readonly-field-preview.ps1
```

These checks are implementation safeguards, not a substitute for clinical commissioning.

## Community and publication

- Use GitHub Issues for anonymized bugs, onboarding problems, and feature requests.
- See `CONTRIBUTING.md` before changing sources, aliases, PlanCheck, or ESAPI adapters.
- The reproducible technical note is in `paper\ClearPlan_ZMP_short_communication.md`.
- A proposed, nonuniversal local transfer checklist is in
  `docs\local-commissioning-scaffold.md`.
- Versioned software and simulator assets are published at
  `https://github.com/Kiragroh/ClearPlan-OpenSource/releases/tag/v3.1.0`.

ClearPlan is a framework; each institution remains responsible for its clinical rules, validation, deployment, and governance.
