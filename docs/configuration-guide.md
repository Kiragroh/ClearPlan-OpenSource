# Configuration guide

[Documentation](index.md) · [Clinic onboarding](CLINIC_ONBOARDING.md) · [Developer handoff](developer-handoff.md)

Use a reviewed local copy. Public defaults show formats and workflows; they are
not prescriptions, commissioned constraints, or authorization for clinical use.

The [actual Settings gallery](settings-gallery.md) shows these controls with
public, patient-free configuration and clearly identified demo history.

## Sources and paths

Open **Settings** and its paths/source page. File-system paths belong in
`settings.ini`; other options belong in `settings.json`. Relative paths resolve
from the loaded settings directory, not the terminal directory. Distribution
templates are not automatically activated in an already running deployment.

| Mode | Behavior |
|---|---|
| `Automatic` | Try valid configured RefDB first; visibly fall back to Excel |
| `RefDb` | Require RefDB; no silent Excel substitution |
| `Excel` | Require the workbook; no silent RefDB substitution |

Test the source before saving. Inspect resolved path, catalog, errors, and fallback.
Excel uses `Structures`, `Tables`, and `Constraints`; keep IDs unique, relationships
explicit, and units valid. Do not merge cells in data ranges. Stock tables require
confirmation; `ReviewRows` is not executable policy. See [stock maintenance](../tools/stock-catalog/README.md).

## Editable settings

| Item | Path key | Effect |
|---|---|---|
| Constraints | `RefDbJsonPath`, `ExcelWorkbookPath` | Tables, goals, variations, source mappings |
| Aliases | `StructureAliasesJsonPath` | Mapping overrides; never TPS renaming |
| Default target rules | `DefaultReviewRulesJsonPath` | Target types, metric, comparator, prescription fraction |
| Field names | `FieldNamingRulesJsonPath` | Read-only expectations, token order, separators |
| Check inclusion | `PlanCheckSelectionJsonPath` | Review/report scope; not legacy execution |
| CT approvals | `CtCompatibilityJsonPath` | Exact manufacturer/model/serial/HU tuple |
| Isodose display | `IsodoseDisplayJsonPath` | Percentages, colors, visibility |
| MLC geometry | `MlcGeometryProfilesJsonPath` | Exact leaf/layer mapping and limits |
| Dose-rate estimate | `DoseRateProfilesJsonPath` | Exact identity matching and assumed speed |
| Collision | `CollisionProfilesJsonPath`, `SourceCollisionModelsJsonPath` | Local geometry/evidence; public catalogs empty |
| ARIA report transfer | `AriaUploadConfigJsonPath` | Optional integration; blank disables upload |

Keep canonical structure names consistent with the intended TG-263 convention and
preserve laterality. Do not equate cropped/derived structures without an accepted
explicit rule. See [aliases](STRUCTURE_ALIASES.md).

Shipped target examples use strict `D98% > 95%` of assigned target Rx for PTVs and
`D98% > 100%` for CTV/GTV/ITV. They are editable examples, not universal recommendations.
Missing or ambiguous Rx stays unevaluated; removed/disabled rules are not recreated.
See [target rules and PAM](review-controls/README.md#target-rules-and-pam).

Default arc naming places couch before direction: `179-181 T30 UZ`. ID and name
conformance are separate. [Field naming](DEFAULT_FIELD_NAMING.md) never changes TPS fields.

## Controlled changes and rollback

1. Select the document in **Configurations and versions**.
2. Import a source copy or edit a separate draft; external originals are preserved.
3. Enter a reason, validate, and save a revision. Save/close Excel drafts before importing changes.
4. Review revision, UTC time, actor, and hash.
5. Save Settings to activate the immutable revision path; rebuild the review and verify results.

Restore appends earlier bytes as a new revision, preserving later history. Back up
the complete managed directory. On integrity failure, preserve and investigate the
store; do not rewrite hashes or delete history. [History details](CONFIGURATION_HISTORY_STORE.md)
explain why recorded actors and hashes are not tamper-proof authentication.

## Display and report scope

Language affects presentation only. Structure visibility is shared across DVH/CT
tabs and followed by reports. Unmatched goals are hidden by default with a disclosed
count; mapping remains available. Optional field-start BEV/DRR and geometry panels
can be omitted. Hiding a row or image never makes its evidence pass.

Isodose editing adds/removes levels and changes colors or visibility. Current
display choices affect exports; versioned defaults support future use. This is
rendering of sampled dose, not recalculation of TPS dose or clinical goals.

## Privacy, integrations, and migration

Configure protected outputs and test denied/unavailable paths and their actual
fallback destinations. Identifier-light filenames do not anonymize report contents.
Leave ARIA upload disabled during read-only acceptance. Provision and commission
FHIR/VAIS through Varian support, or use the clinic's approved Webservice/direct
upload or eDocPrinter route; those are not automatic FHIR fallbacks. See [ARIA setup](ARIA_REPORT_UPLOAD.md).

Legacy CSVs are not runtime defaults. Convert retained copies with
`tools/convert-constraint-csvs-to-xlsx.py`, verify record equivalence with
`tools/verify-constraint-migration.py`, and independently review mappings, scope,
units, and laterality. Retain originals until acceptance and rollback are verified.
