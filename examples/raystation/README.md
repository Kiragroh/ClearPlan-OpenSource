# ClearPlan for RayStation: local review rules and the shared review log

> **UNVALIDATED EXAMPLE. NOT FOR CLINICAL USE.**
>
> This code has **not** been run, validated or commissioned in a clinical (or any) RayStation
> installation. All tests use a **simulated RayStation object model**. Several RayStation API calls are
> assumptions that must be verified locally (table below). The log contains confidential clinical
> dosimetry even when identifiers are pseudonymised; keep it inside an institution-approved environment.

## 1. Purpose

ClearPlan's argument does not depend on a particular planning system: a department writes its *own*
plan-review rules (goal tables, structure aliases, a few checks), applies them to every plan, and records
every result, including passes, in a structured, SQL-analysable log. That log is what makes local review
effectiveness measurable (recurring findings, goal-status shares per protocol, rule-change effects,
retrospective baselines with identical rules).

This folder shows the same concept for RayStation. A RayStation clinic can apply the same kind of local
rules and write **the same five review-log tables** as the Eclipse/ESAPI implementation:

| Table | One row per | Identical to the C# writer |
|---|---|---|
| `review` | evaluated plan (beam set) | headers, `schema_version = 1`, vocabularies, counts, hashes |
| `goal_result` | goal-table row | status `goal / variation / not_met / not_evaluated` |
| `plan_parameter` | descriptor (MU, MU/Gy, Paddick CI, GI, HI, ...) | names, units, method identifiers |
| `check_finding` | check result, **including OK** | status `ok / variation / fail / info / not_evaluated` |
| `ruleset_revision` | rule document used | SHA-256 per document, combined `ruleset_sha256` |

Eclipse and RayStation logs can share **one key file and one SQLite database**
(`tools/review-log/import_review_log.py`, views in `tools/review-log/schema.sql`). The merge is tested
against the C#-written example logs (section 7).

## 2. Layout

| Path | Role | Touches the RayStation API? |
|---|---|---|
| `raystation_adapter.py` | Thin read-only adapter: RayStation objects -> neutral `PlanData`; lazy read-only dose queries | **yes** (only here ...) |
| `raystation_patient_loader.py` | Opening listed patients for mining (`PatientDB`); **VERIFY LOCALLY** | **yes** (... and here) |
| `clearplan_review_core/` | TPS-independent core: goal tables, structure mapping, evaluator, objectives, plan descriptors, checks, review-log writer, HTML summary | no (AST-tested) |
| `run_live_review.py` | Entry script: review the currently open plan, `mode = live` | via the adapter |
| `run_retrospective_mining.py` | Entry script: explicit plan list, `mode = retrospective` | via adapter and loader |
| `rules/` | Illustrative goal table (CSV and JSON), alias file, check configuration, run configuration, plan list | no |
| `tests/` | Fake RayStation object model and unit and end-to-end tests | no |
| `clearplan_review_snapshot.py` | Earlier, independent snapshot-JSON adapter example (section 9) | yes |

Standard library only; written for the CPython 3.8+ scripting environment of current RayStation
versions (no numpy). Tested with CPython 3.8.20 and 3.11.18 (simulated API).

## 3. The shared rule and log format

* **Goal semantics** are a port of ClearPlan's `PqmNumericEvaluator`: a row is `goal` only for an
  explicit pass; with a variation tier it can be `variation`; otherwise `not_met`. Anything not strictly
  parseable, any unresolved or ambiguous structure, and any plan outside the goal table's fractionation
  scope is `not_evaluated`. A `not_evaluated` row is never counted as passed. The measured value is still
  logged when only the scope is unconfirmed.
* **Structure mapping**: TG-263 template name plus explicit aliases (goal-table column and/or a
  `ClearPlan.structure_aliases.v1` file, the same format as the Eclipse version). Case, spaces and
  underscores are ignored; TG-263 markers (`~ - + ^ ! = /`) are kept. Exactly one contoured ROI resolves.
  Several matches are `ambiguous`, none is `unresolved`. Both are not evaluated and **never substituted**.
* **CSV**: UTF-8 without BOM, RFC 4180 (comma, CRLF, double-quote escaping), invariant numbers in .NET
  `"R"` form (`25`, `0.30000000000000004`, `1E-05`), ISO-8601 UTC with milliseconds and `Z`.
  An empty field means unavailable; it is never written as 0. Rows are append-only. A file with a different
  header is never appended to. The `review` row is written last, and a shared lock file
  `.review-log.lock` serialises writers (compatible with the C# writer's exclusive lock).
* **Pseudonyms**: HMAC-SHA256 with the same 32-byte base64 key file as ClearPlan for Eclipse, truncated
  to 128 bit, prefixes `pat_`, `crs_`, `pln_`, `usr_`. The same key gives the same pseudonym in both systems.
  For example, the same Windows login gives the same `user_key`. Mapping to RayStation:
  `patient_key` = PatientID; `course_key` = PatientID + case name; `plan_key` = composite
  PatientID + case name + `"<plan name>/<beam set DicomPlanLabel>"` (no DICOM plan UID attribute is
  assumed, so live and retrospective reviews of the same beam set link); `user_key` = Windows login
  name. Free text is scrubbed of the patient ID, name parts and user ID. The key must not lie inside the
  log directory (enforced).
* **Implementation provenance** (schema version 1 has no planning-system column):
  `clearplan_version = clearplan-raystation-example/0.1.0`, and `build_id = py-sha256:<16 hex>` over
  the core package, adapter, loader and entry scripts. Proposed additive column: see section 8.
* **Rule documents** in `ruleset_revision`: `constraints-goal-table` (category `constraint_catalog`),
  `structure-aliases` (`structure_aliases`, same key as Eclipse) and `review-checks`
  (`check_configuration`).

## 4. Configuring rules

**Goal table** (`rules/example_goal_table.json` or `.csv`; the values are illustrative software-test
numbers, **not a clinical protocol**):

| Column | Required | Meaning |
|---|---|---|
| `template_structure` | yes | TG-263 primary name, e.g. `Rectum` |
| `aliases` | no | `|`-separated exact alternative names, e.g. `Rektum` |
| `objective` | yes | `D98%[Gy]`, `D2%[Gy]`, `D0.035cc[Gy]`, `Dmax[Gy]`, `Dmean[Gy]`, `V20Gy[%]`, `V95%[%]` (dose in % of Rx), `CV20Gy[cc]` |
| `comparator` | yes | `<`, `<=`, `>`, `>=`, `=` |
| `goal` | yes | number (dot or comma decimal; no grouping) |
| `variation` | no | acceptable-variation limit; must not be stricter than `goal` |
| `template_code`, `source_label` | no | stable row code (default `<structure>-<objective>`), rule source label |
| `fractions`, `dose_per_fraction_gy`, `total_dose_gy` | no | table scope (CSV: identical in all rows; JSON: `"scope"`) |

JSON additionally allows `"targets"` (template names used for target descriptors).
The loader rejects a malformed table as a whole and never applies part of it.

**Quantities** (all from read-only queries on `TreatmentCourse.TotalDose`; cGy is converted to Gy):

| Quantity | Definition used |
|---|---|
| `D<x>%` | `GetDoseAtRelativeVolumes([x/100])` |
| `D<x>cc` | relative volume = x / ROI volume (`GetRoiVolume`, cm3); unavailable if x > ROI volume |
| `Dmax / Dmin / Dmean` | `GetDoseStatistic(Max / Min / Average)` |
| `V<d>[%]`, `V<d>[cc]` | `GetRelativeVolumeAtDoseValues([d])` x 100 or x ROI volume |
| `CV<d>` | ROI volume (or 100 %) minus V(d) |
| `paddick_ci` | TV_PIV^2 / (TV x PIV); TV_PIV = target rel. volume at Rx x TV, PIV = External rel. volume at Rx x External volume |
| `gradient_index` | V_External(50 % Rx) / V_External(100 % Rx) |
| `homogeneity_index` | (D2% - D98%) / Rx |
| `total_mu`, `mu_per_gy` | sum of `BeamMU`; total MU / dose per fraction |
| `pam`, `mean_aperture_cm2`, `small_aperture_fraction`, `estimated_duration_s` | **written as `unavailable` with a note**: they need MLC leaf geometry or a dose-rate profile that this read-only example does not read. |

ROIs with zero or unknown volume are not evaluated. Rx is the beam-set prescription. Plans with more than
one beam set are not evaluated against the table and the run is `partial`, because `TotalDose` cannot be
attributed to one fractionation.

**Checks** (`rules/example_review_checks.json`, schema `ClearPlan.review_checks.v1`). Every enabled
check writes exactly one row, OK included:

| `check_code` | Config | Result |
|---|---|---|
| `plan-has-dose` | – | ok / fail / not_evaluated (dose presence unknown) |
| `fractionation-scope` | goal-table scope | ok / fail (mismatch: goals not evaluated) / info (no scope declared) / not_evaluated |
| `ct-age-days` | `ct_max_age_days` | ok / fail; not_evaluated if the examination date is missing or in the future; info if not configured |
| `ct-device-approved` | `approved_ct_devices` (manufacturer / model / station / imaging_system) | ok / fail; not_evaluated if needed device fields are missing; info if not configured |
| `field-naming` | `field_name_regex` (full match of every beam name) | ok / fail (count of non-matching beams) / info |

`disabled_checks` skips checks; they are counted in `review.n_checks_excluded`.

## 5. Running

Copy `rules/example_run_config.json` (schema `ClearPlan.raystation_run.v1`), and set `log_dir`, `key_file`
(the **same** key file as ClearPlan for Eclipse if logs are to be merged) and the rule paths. Relative
paths resolve from the configuration file. The scripts find it via `CLEARPLAN_RS_RUN_CONFIG` or as
`clearplan_rs_run.json` next to the scripts.

**Live review** (plan open in RayStation, run `run_live_review.py` as a reviewed script):
appends one `mode = live` review. If `html_summary_dir` is set, it also writes an identifier-free
HTML summary of exactly the logged rows (not a clinical report; `report_written` stays 0).

**Retrospective mining**: `run_retrospective_mining.py` with an explicit, authorised list
(`rules/example_mining_list.csv`: `patient_id,case_name,plan_name[,beam_set]`; `#` lines are ignored).
Each row is opened through `raystation_patient_loader.open_plan_readonly` and reviewed with the same
rules (`mode = retrospective`). Errors give a `failed` review row with a short code (`patient_not_found`,
`case_not_found`, `beam_set_ambiguous`, `read_error:<type>`, ...), and the batch continues. Nothing is
saved. Start mining with no patient open. `--validate-input-only` only checks the list.

**SQL**:

```powershell
python ..\..\tools\review-log\import_review_log.py --log-dir <eclipse_log> --log-dir <raystation_log> --db review_log.sqlite --export-views views
```

## 6. Safety boundary

* No RayStation write calls of any kind: no Save, Create, Delete, Set*, Add*, Remove*, Import, Export,
  Update, Compute, ROI algebra or dose recomputation, and no attribute assignment on RayStation objects.
  `tests/test_readonly_boundary.py` enforces this statically (AST): every PascalCase method call in the
  adapter, loader and entry scripts must be on an explicit allowlist, write-like identifiers and
  attribute stores are rejected, `PatientDB`/`LoadPatient` may appear only in the loader, and the core
  imports no TPS module. At runtime, the fake model records every call and the tests assert the allowlist.
* `LoadPatient` is called with `AllowPatientUpgrade=False`.
* No network shares, site scripts, ARIA or other site infrastructure are required. Paths are local
  configuration.

## 7. What is tested (simulated RayStation API only)

`python -m unittest discover -s examples\raystation\tests -t examples\raystation\tests -v`

* Evaluator: every case of `PqmNumericEvaluatorTests.cs` (goal/variation/not met, unconfirmed scope,
  invalid values, unit conversion, complement volume, ratios).
* Objective syntax, CSV/JSON goal-table equivalence and rejection of malformed tables. Structure mapping:
  aliases, ambiguity, missing or uncontoured ROI, no semantic substitution.
* Quantities on a phantom-like fake plan (5 x 5 Gy; PTV 100 cm3, rectum 50 cm3, External; two beams)
  against hand-calculated values. Edge cases: missing ROI, ambiguous alias, zero ROI volume, volume >
  ROI, missing or old or future CT date, unknown CT device, fractionation mismatch, two beam sets, no dose,
  non-unique External, bad beam name.
* Writer: RFC 4180 escaping and round-trip, .NET "R" number format, UTC format, header check, refusal on a
  foreign header, repair of an interrupted line, lock contention, empty counts for failed runs.
  Pseudonyms are deterministic and match **C#-written test vectors** for the same test key. The headers
  equal `ReviewLogSchema.cs` and `schema.sql`.
* End to end: fake plan -> live and retrospective runs -> CSV -> `import_review_log.py` -> SQLite ->
  views (statuses, counts, denominators incl. OK rows, idempotent re-import, no clear identifiers).
* Merge (set `CLEARPLAN_REVIEW_LOG_EXAMPLE` to the C#-written example): C# and RayStation logs import into
  one database without duplicates or orphans. Python reproduces the C# `ruleset_sha256` and every
  numeric field of the C# files, and a shared key links the same user across both systems.

`tests/export_example_output.py <empty dir> [<C# example dir>]` regenerates the example output
(deterministic: fixed fake clock and review IDs, test-only key).

**Not tested and not claimed:** behaviour of a real RayStation (any version), correctness of the
`[verify]` calls below, dose-engine agreement, clinical suitability of the example rules, concurrent
writing with the C# writer on a real file share.

## 8. RayStation API calls and their provenance

`[example]` = already used by `clearplan_review_snapshot.py`. That example is itself unvalidated;
"from the example" does **not** mean verified. `[verify]` = assumption made here; check it in a
non-clinical installation of your RayStation version before use.

| Call / attribute | Used for | Provenance |
|---|---|---|
| `raystation` / `raystation.v2025` / `connect` -> `get_current(...)` | context accessor | [example] |
| `get_current("Patient" / "Case" / "Plan" / "BeamSet" / "Examination")` | live review | [example] |
| `Case.PatientModel.RegionsOfInterest`, `Roi.Name`, `Roi.Type` | ROI list, External detection | [example] |
| `Case.PatientModel.StructureSets[exam.Name].RoiGeometries[name]` | geometry of the planning exam | [example] |
| `RoiGeometry.HasContours()`, `RoiGeometry.GetRoiVolume()` (cm3) | contour presence, volume | [example] |
| `Plan.TreatmentCourse.TotalDose` | dose for all quantities | [example] |
| `TotalDose.GetDoseAtRelativeVolumes(RoiName, RelativeVolumes)` (cGy) | D x %, D x cc, D2/D98 | [example] |
| `TotalDose.GetRelativeVolumeAtDoseValues(RoiName, DoseValues)` (cGy in, fraction out) | V, CV, Paddick CI, GI | [example] |
| `TotalDose.GetDoseStatistic(RoiName, DoseType="Max" / "Average")` | Dmax, Dmean | [example] |
| `TotalDose.GetDoseStatistic(..., DoseType="Min")` | Dmin | [verify] |
| `TotalDose.DoseValues` (None = no dose) | `plan-has-dose` | [verify] |
| `BeamSet.FractionationPattern.NumberOfFractions` | fractions | [example] |
| `BeamSet.Prescription.PrimaryDosePrescription.DoseValue` (cGy, total) | Rx | [example] |
| `BeamSet.Prescription.PrimaryPrescriptionDoseReference` (fallback name) | Rx on newer API | [verify] |
| `...PrimaryDosePrescription.OnStructure.Name` | default target | [verify] |
| `BeamSet.Beams` | beams | [example] |
| `Beam.Number`, `Beam.Name`, `Beam.BeamMU` (assumed MU per fraction) | MU, field naming | [verify] |
| `BeamSet.MachineReference.MachineName`, `BeamSet.DeliveryTechnique` | `machine_id`, `technique` | [verify] |
| `BeamSet.DicomPlanLabel` | plan key | [verify] (present only in the example's fake objects) |
| `BeamSet.GetPlanningExamination()` | planning exam (fallback: current exam, run `partial`) | [verify] |
| `Examination.Name` | structure-set key | [example] |
| `Examination.GetExaminationDateTime()` (.NET/py datetime, local time) | CT age | [verify] |
| `Examination.GetAcquisitionDataFromDicom()["StudyModule"]["StudyDateTime"]` | CT age fallback | [verify] |
| `...["EquipmentModule"]["Manufacturer" / "ManufacturersModelName" / "StationName"]` | CT device | [verify] |
| `Examination.EquipmentInfo.ImagingSystemReference.ImagingSystemName` | CT device (imaging system) | [verify] |
| `Patient.PatientID`, `Patient.Name`, `Case.CaseName`, `Plan.Name`, `Plan.BeamSets` | pseudonyms, beam-set count | [verify] |
| `Plan.Review.ApprovalStatus`, `Plan.Review.ReviewTime` | approval fields | [verify] |
| `get_current("PatientDB").QueryPatientInfo(Filter={"PatientID": id})` (exact match re-checked) | mining | [verify] **verify locally** |
| `PatientDB.LoadPatient(PatientInfo=..., AllowPatientUpgrade=False)` | mining | [verify] **verify locally** |
| `Patient.Cases`, `Case.TreatmentPlans` (iteration) | mining | [verify] |
| (not RayStation) `USERNAME` environment variable | `user_key` | – |

Proposed additive schema changes (not implemented, so that the five tables stay byte-compatible):
`review.planning_system` (`eclipse` / `raystation`) and `review.plan_key_basis` (`dicom_uid` /
`composite`); a `source_kind` value for native TPS goals of other systems (e.g. RayStation clinical
goals, which this example does not log); a `day` unit code for `ct-age-days` (now `count`, as in the
Eclipse example).

## 9. Earlier example: identifier-free review-snapshot JSON

> **UNVALIDATED EXAMPLE — NOT FOR CLINICAL USE**
>
> This code demonstrates a read-only adapter boundary for RayStation v2025 SP2
> (scripting API 17.2.0). It has not been run or commissioned in a RayStation
> clinical environment. The generated JSON omits source identifiers but still
> contains **confidential clinical dosimetry**. Do not commit, publish, attach
> to an issue, or move the output outside an institution-approved local
> environment.

`clearplan_review_snapshot.py` reads the currently opened RayStation context and
writes a ClearPlan review-snapshot JSON file. It is an interoperability example,
not a claim that Eclipse/ESAPI rules and RayStation clinical goals are
semantically equivalent.

### Deliberate safety boundary

- Reads only the current `Patient` presence, `Case`, `Examination`, `Plan`, and
  `BeamSet`; it does not query or load other patients.
- Uses `TreatmentCourse.TotalDose` only. RayStation dose values are converted
  from cGy to Gy, relative volume fractions to percent, and ROI volume is kept
  in cm³.
- Samples cumulative DVHs at 1 Gy by default and rejects non-finite,
  out-of-range, or nonmonotone values instead of silently repairing them.
- Replaces ROI labels with deterministic role pseudonyms such as `target-01`
  and `oar-01`. Patient, case, examination, plan, beam-set, beam, ROI, and DICOM
  identifiers are not serialized.
- Transcribes numeric `DoseAtVolume`, `VolumeAtDose`, and `AverageDose` goal
  values, but marks every PQM row `not-evaluated`/`info`.
- Emits one explicit `not-evaluated` PlanCheck row because institution-specific
  RayStation checks have not been implemented or commissioned.
- Derives PlanFieldNamer-style suggestions from gantry, couch, and arc-direction
  geometry while omitting actual field IDs and names. The couch token precedes
  `UZ`/`GUZ`, including duplicate suffixes.
- Contains no RayStation save, create, delete, update, set, add, remove, import,
  or DICOM-export calls.
- Accepts only an absolute non-UNC output path. UNC destinations and implicit
  overwrite are rejected; mapped or redirected storage still requires local
  information-security review.

These safeguards reduce disclosure and modification risks; they do not make the
example clinically validated.

### Illustrative use inside RayStation

Run this as a reviewed file script or wrap it in an institution-approved
database script. RayStation command-line `-run` executes an approved database
script name; it is not a stand-alone API process.

```python
from clearplan_review_snapshot import export_current_review_snapshot

export_current_review_snapshot(
    r"C:\ClearPlan-local\raystation-review.json",
)
```

The local parent directory must already exist. Existing files are preserved.
If replacement is intentional, pass `overwrite=True`.

The adapter resolves the context accessor lazily in this order:

1. `raystation`
2. `raystation.v2025`
3. legacy `connect`

The fallback is present only to make the example legible across API generations;
it is not evidence of compatibility with an installed RayStation version.

### Tests outside RayStation

The unit tests use identifier-rich fake objects and verify that none of those
source values enters the JSON. They also check units, DVH monotonicity, supported
goal transcription, field geometry suggestions, path restrictions, lazy import
order, deterministic output, and an AST-based ban on write-like RayStation API
calls.

```powershell
python -m unittest discover -s examples\raystation\tests -v
```

Passing these synthetic tests does not validate RayStation runtime behavior.
The public build also exports one deterministic fake-context JSON and checks
that the C# `ReviewSnapshot` validator accepts it. The released simulator does
not import arbitrary adapter JSON, so this is a schema-contract check rather
than an end-to-end UI integration test.
Before any local evaluation, confirm the installed API version, run against
approved test patients only, inspect every output field, perform independent
commissioning, and keep the generated JSON outside version control.
