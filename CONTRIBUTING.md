# Contributing to ClearPlan

[Documentation](docs/index.md) · [Developer handoff](docs/developer-handoff.md)

ClearPlan is a read-only starter framework for automated radiotherapy plan review. Contributions should remain understandable outside one clinic, testable without patient data, and safe when required structures or metadata are absent.

## Before changing code

- Keep source-independent rules in `ClearPlan.Core`.
- Keep vendor API access in `ClearPlan.Script` adapters.
- Do not add proprietary vendor binaries, patient data, private network paths, credentials, or institutional constraint sets.
- Do not add plan or beam modification to the field-name preview.
- Treat the clinical PlanCheck as an independently governed component. Public modernization must not silently replace or narrow it.
- Add a failing test before changing a core rule or fixing a bug.

## Constraint sources

Both RefDB JSON and Excel must map into the same `ConstraintCatalog`. Consumer code must not branch on source-specific fields.

For Excel contributions, keep the catalog maintainable and use the existing `Structures`, `Tables`, and `Constraints` sheets. Required identifiers must be unique and all table/structure relationships valid. Do not merge cells inside the data ranges.

For RefDB contributions, validate the supported schema before joining entities. Preserve active flags, canonical structure identifiers, aliases, side-specific aliases, table metadata, and source diagnostics. A schema mismatch must fail clearly.

Automatic mode may fall back from RefDB to Excel only when the fallback is visible in status and diagnostics. Explicit modes must never fall back silently.

## Aliases and laterality

- Define aliases in source data rather than hard-coding clinic names.
- Prefer canonical identifiers and exact normalized names before broader fallbacks.
- Preserve left/right laterality for paired structures.
- Do not expand a combined-organ alias into both sides unless explicitly represented.
- Report ambiguous matches; never resolve them by row or dictionary order.
- Add tests for canonical, alias, code, DICOM-type, laterality, and ambiguity behavior as applicable.

## Table selection

New selection rules must be deterministic and explainable. Incompatible Plan/PlanSum or fractionation candidates must be excluded before scoring. A tie or low-confidence result must require confirmation. Add tests showing both the intended selection and an ambiguous or incompatible case.

## Field-name preview

The preview follows PlanFieldNamer-compatible conventions for plan-aware field-ID prefixes and numbering, rounded gantry angles, arc direction (`UZ`/`GUZ`), couch token (`T<angle>`), duplicate suffixes, and treatment order. Changes require focused tests with simple beam inputs.

The ESAPI adapter must remain read-only. The release guard scans for plan modification, rename, apply, save, and setup-field calls.

## PlanCheck changes

PlanCheck changes require separate clinical review and are not implied by a source, settings, GUI, or field-preview pull request. When a deployment is expected to preserve the clinical PlanCheck:

1. inventory its public/check surface;
2. record its SHA256 hash;
3. use the explicit deployment manifest;
4. confirm no `ErrorCalculator.cs` copy is proposed;
5. verify the post-deployment hash and surface.

## Settings and privacy

Use `settings.ini` for every file-system path and `settings.json` for non-path options. Prefer relative paths in the public starter and keep optional integrations blank or public-safe. Both files are editable through the ClearPlan settings page; save only after path and catalog validation succeeds.

Public defaults must not log direct patient/user identifiers or place them in filenames. Screenshots, figures, workbooks, test fixtures, logs, and issue reports must be patient-free: use synthetic test data or public-only configuration screens. Masked clinical captures do not meet that requirement.

## Required validation

Before opening a pull request:

```powershell
.\tools\test-portable-review.ps1
git diff --check
```

Use your approved execution policy. For licensed ESAPI changes, also run the
separate native build and protected acceptance described in the
[developer handoff](docs/developer-handoff.md); portable tests do not execute ESAPI.

If constraints changed, also state:

- which source/schema changed;
- workbook or RefDB validation results;
- migration-equivalence result when legacy CSVs were converted;
- the tests added for aliases, laterality, selection, or objectives.

If GUI behavior changed, include a synthetic screenshot and verify both the overview and affected full detail page. If field naming changed, include the exact input beam geometry and expected read-only suggestion.

## Pull request summary

A strong pull request explains:

- what changed and why;
- whether the change is a public starter behavior or an optional local extension;
- clinical, compatibility, and migration implications;
- tests and validation performed;
- any required local commissioning or rollback step.
