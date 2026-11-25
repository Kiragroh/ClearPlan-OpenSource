# Structured review log — data dictionary (schema version 1)

Supplement S1 draft. Applies to ClearPlan source with `ClearPlan.Core/ReviewLog` (portable writer) and
`tools/review-log` (SQLite schema, importer, example views).

## 1. Purpose and scope

Every physics plan review performed with ClearPlan can leave a structured, machine-readable record:
which goals were evaluated with which result, which checks fired (and which passed), which plan
descriptors were obtained, and which rule-set revision defined the review. The record is the data basis for
(i) local self-improvement (which goals end in *Variation* or *Not met*, which checks recur), (ii) the effect of
rule changes (before/after a rule-set revision) and (iii) a retrospective baseline obtained with identical rules.
The log does not establish review effectiveness by itself; it makes it measurable.

## 2. Principles

| Principle | Implementation |
|---|---|
| One row per entity | `review` (1 per evaluated plan), `goal_result` (1 per goal row), `plan_parameter` (1 per quantity), `check_finding` (1 per check result, **including OK**), `ruleset_revision` (1 per configuration document) |
| Append-only | Rows are only appended. Existing files must carry the exact expected header; otherwise the writer refuses to append (new schema version ⇒ new files). |
| Stable names | snake_case columns; order and names are fixed within a schema version (`ReviewLogSchema.cs`, `schema.sql`; a unit test compares both). |
| Encoding | UTF-8 without BOM, RFC 4180 CSV (comma separator, CRLF, double-quote escaping), invariant culture. |
| Time | ISO 8601 UTC, millisecond precision, literal `Z` (`2026-07-03T08:15:00.000Z`). |
| Numbers | Numeric value in its own column, unit in a separate column; round-trip precision; never strings such as `24.2 Gy`. |
| Missing values | Empty field = not available (SQL `NULL`). **Never 0.** NaN/Infinity are written as empty. |
| Not evaluated | A goal is `goal` only for an explicit pass. Any other or unknown state is `not_evaluated`; it is never counted as passed. |
| Same tables in all modes | `mode` = `live` (GUI), `headless` (report entry), `retrospective` (mining). |
| Concurrency | A directory lock file (`.review-log.lock`, exclusive, bounded retry) serialises writers across processes and workstations. Within one review the `review` row is written **last**; child rows without a `review` row (interrupted write) are ignored by the importer. |

## 3. Privacy

* Identifiers are pseudonymised by default with **HMAC-SHA256** keyed with a local secret (32 random bytes,
  base64, created once at `ReviewLogKeyFile`). Keys are truncated to 128 bit and prefixed: `pat_`, `crs_`, `pln_`, `usr_`.
  Pseudonyms are deterministic for one key (linkage across reviews) and cannot be recomputed without the key
  (unlike a plain or salted hash of a short patient ID, which is open to dictionary attack).
  The key is never written to the log. Store it with restricted access and separately from the log directory.
  All workstations that should produce linkable pseudonyms must share the same key file.
* `patient_key` = HMAC(patient ID, case-insensitive); `course_key` = HMAC(patient ID + course ID);
  `plan_key` = HMAC(DICOM plan UID) — for plan sums (no UID) HMAC(patient + course + plan-sum ID); `user_key` = HMAC(user ID without domain).
* Clear identifiers are written only if `reviewLog.logClearIdentifiers = true` (local policy); `identifier_mode` records which mode produced a row.
* Display labels that may contain names are not logged. `plan_label` (plan ID) is empty unless `reviewLog.logPlanLabel = true`.
* Free text (`message`, `observed_value`, `expected_value`, `note`, `run_message`) is scrubbed: known identifying terms of the review
  (patient ID and names, user ID/name, approver IDs/names) are replaced by `[redacted]`; text is truncated to 2000 characters.
  Free text from legacy checks can still contain unforeseen identifiers; set `reviewLog.logMessages = false` to omit it.
* Structure IDs (`template_structure`, `resolved_structure`, `entity` for targets) are logged; they are clinical labels, not person identifiers.
  Local data-protection review remains required.

## 4. Configuration

`settings.json` (options):

```json
"reviewLog": { "enabled": false, "logClearIdentifiers": false, "logPlanLabel": false, "logMessages": true }
```

`settings.ini` (`[Paths]`):

```ini
ReviewLogDirectory = Logs\ReviewLog
ReviewLogKeyFile = State\review-log.key
```

Logging is **off by default** and must be enabled by local policy. Relative paths resolve from the ClearPlan build folder.

## 5. Tables

Types refer to SQLite (`schema.sql`). "Empty" = not available.

### 5.1 `review` — one row per evaluated plan

| Column | Type | Unit | Definition |
|---|---|---|---|
| schema_version | INTEGER | – | Log schema version (1). |
| review_id | TEXT | – | GUID of this review (primary key). |
| mode | TEXT | – | `live` (GUI), `headless` (report entry), `retrospective` (mining batch). |
| started_utc | TEXT | UTC | Review start (GUI: review window opened; headless/retrospective: start of evaluation of this plan). |
| finished_utc | TEXT | UTC | Time the record was produced (report saved, review closed, or batch item finished). |
| clearplan_version | TEXT | – | File version of `ClearPlan.esapi.dll`. |
| build_id | TEXT | – | Module version ID (MVID) of the native assembly; identifies the exact build. |
| identifier_mode | TEXT | – | `hmac_sha256` or `clear`. |
| user_key | TEXT | – | Pseudonymised (or clear) reviewing user. |
| patient_key | TEXT | – | Pseudonymised (or clear) patient. |
| course_key | TEXT | – | Pseudonymised (or clear) course, patient-scoped. |
| plan_key | TEXT | – | Pseudonymised (or clear) plan, from the DICOM plan UID. |
| plan_label | TEXT | – | Plan ID; empty unless explicitly enabled. |
| planning_item_type | TEXT | – | `plan` or `plan_sum`. |
| plan_created_utc | TEXT | UTC | Plan creation time (TPS). |
| plan_approval_status | TEXT | – | TPS approval status at review time (e.g. `PlanningApproved`). |
| planning_approval_utc | TEXT | UTC | Planning approval time; empty if absent/unparsable. |
| treatment_approval_utc | TEXT | UTC | Treatment approval time; empty if absent/unparsable. |
| machine_id | TEXT | – | Distinct treatment-unit IDs of treatment fields, `|`-separated. |
| technique | TEXT | – | Distinct field techniques, `|`-separated. |
| mlc_model | TEXT | – | Distinct MLC models, `|`-separated. |
| total_dose_gy | REAL | Gy | Prescribed total dose of the active plan. |
| dose_per_fraction_gy | REAL | Gy | Dose per fraction. |
| fractions | INTEGER | – | Number of fractions. |
| goal_table_id | TEXT | – | Name of the selected constraint (goal) table. |
| goal_table_sha256 | TEXT | – | SHA-256 over the canonical definitions (template code, structure, objective, comparator, goal, variation, unit) of the constraint-table goal rows evaluated in this review. |
| ruleset_sha256 | TEXT | – | Combined rule-set fingerprint: SHA-256 over sorted `config_key=sha256` lines of `ruleset_revision` plus `goal_table=<goal_table_sha256>`. Changes whenever any rule document or the evaluated goal table changes. |
| n_goals | INTEGER | count | Goal rows (= rows in `goal_result`). |
| n_goal / n_variation / n_not_met / n_not_evaluated | INTEGER | count | Goal rows per status. |
| n_checks | INTEGER | count | Included check rows (`check_finding.included = 1`). |
| n_checks_excluded | INTEGER | count | Check results calculated but excluded by the configured check selection. |
| n_findings_open | INTEGER | count | Included checks with status `variation`, `fail` or `not_evaluated`. |
| report_written | INTEGER | 0/1 | 1 if this review produced a PDF report. |
| report_sha256 | TEXT | – | SHA-256 of the written PDF (links the record to the archived report). |
| report_filed | INTEGER | 0/1 | 1 if the report was filed to the OIS by this record (currently always 0; filing is a separate confirmed action). |
| run_status | TEXT | – | `ok`, `partial` (e.g. plan descriptors unavailable), `failed` (no evaluated content; counts empty). |
| run_message | TEXT | – | Short technical reason (scrubbed); never an exception message. |

### 5.2 `goal_result` — one row per clinical-goal row

| Column | Type | Unit | Definition |
|---|---|---|---|
| review_id | TEXT | – | → `review.review_id`. |
| row_no | INTEGER | – | 1-based row order in the review. |
| stable_id | TEXT | – | Stable row ID from the review snapshot. |
| source_kind | TEXT | – | `constraint_table` (local Excel/RefDB table), `eclipse_clinical_goal` (native TPS goals), `default_rule` (ClearPlan default target rule). |
| template_code | TEXT | – | Constraint/template code (TG-263 where configured). |
| template_structure | TEXT | – | Structure name in the rule. |
| resolved_structure | TEXT | – | Structure ID in the plan's structure set; empty if unresolved. |
| mapping_status | TEXT | – | `resolved`, `ambiguous`, `unresolved`. |
| objective | TEXT | – | Objective in ClearPlan syntax, e.g. `D0.035cc[Gy]`, `CV20Gy[cc]`, `D95%[Gy]`. |
| comparator | TEXT | – | `<`, `<=`, `>`, `>=`, … |
| goal_value | REAL | `unit` | Preferred (goal) threshold. |
| variation_value | REAL | `unit` | Acceptable-variation threshold; empty if the rule has no variation tier. |
| achieved_value | REAL | `unit` | Value achieved in the plan; empty if not measurable (never 0). |
| unit | TEXT | – | Unit of goal/variation/achieved value (`Gy`, `%`, `cm3`, …). |
| status | TEXT | – | `goal`, `variation`, `not_met`, `not_evaluated` (see §2). |
| priority | TEXT | – | Reserved; empty in schema version 1 (priority is not part of the detached snapshot). |
| source_label | TEXT | – | Configured source label of the rule (e.g. table/reference), not a patient label. |

### 5.3 `plan_parameter` — one row per quantity

| Column | Type | Unit | Definition |
|---|---|---|---|
| review_id | TEXT | – | → `review.review_id`. |
| row_no | INTEGER | – | 1-based row order (key). |
| scope | TEXT | – | `plan`, `target`, `beam`. |
| entity | TEXT | – | Target structure ID (`target`; for plan-level `pam` the PAM target) or `beam-<number>` (`beam`); empty for plan scope. |
| name | TEXT | – | Quantity name (table below). |
| value | REAL | `unit` | Value; empty if unavailable. |
| unit | TEXT | – | Unit (`1` = dimensionless). |
| method | TEXT | – | Definition/aggregation identifier (table below). |
| availability | TEXT | – | `available` (finite value present) or `unavailable` (value empty). |
| note | TEXT | – | Reason for unavailability or provenance (scrubbed). |

Parameter definitions (Supplement S2 draft):

| scope | name | unit | method | Definition |
|---|---|---|---|---|
| plan | total_mu | MU | sum_beam_meterset | Sum of treatment-field MU. |
| plan | mu_per_gy | MU/Gy | total_mu_per_fraction_dose | Total MU / dose per fraction. |
| plan | pam | 1 | pam_<weighting>_weighted | Plan aperture modulation (0 = target fully open … 1 = fully blocked); between fields weighted by `Beam.WeightFactor` (native) or MU (synthetic/import); trapezoidal cumulative-meterset sampling within fields. `entity` = PAM target. |
| plan | mean_aperture_cm2 | cm2 | mu_weighted_mean_cp_aperture | MU-weighted mean control-point aperture area at isocentre. |
| plan | small_aperture_fraction | 1 | fraction_cp_area_lt_<t>cm2_mu_weighted | MU-weighted fraction of control points with aperture area below threshold *t* (default 4 cm²). |
| target | paddick_ci | 1 | paddick_tv_rx2_over_tv_vrx | Paddick CI = TV_Rx² / (TV × V_Rx). |
| target | paddick_ci_reciprocal | 1 | inverse_paddick | 1 / Paddick CI. |
| target | gradient_index | 1 | body_v50_over_v100 | GI = V_50%Rx / V_Rx in EXTERNAL (whole plan). |
| target | homogeneity_index | 1 | d2_minus_d98_over_rx | HI = (D2% − D98%) / Rx. |
| target | target_volume_cm3 | cm3 | structure_volume | Target volume. |
| target | reference_dose_gy | Gy | plan_prescription | Rx used for the indices (displayed plan prescription). |
| beam | total_mu, pam, mean_aperture_cm2, small_aperture_fraction | as above | as above | Per treatment field. |
| beam | estimated_duration_s | s | dose_rate_profile_estimate | Beam-on time estimate from the configured dose-rate profile (not measured delivery). |

Values are descriptive; no threshold is implied.

### 5.4 `check_finding` — one row per check result, including OK

| Column | Type | Unit | Definition |
|---|---|---|---|
| review_id | TEXT | – | → `review.review_id`. |
| row_no | INTEGER | – | 1-based row order. |
| check_code | TEXT | – | Stable check identifier (legacy PlanCheck ID or ClearPlan check code; duplicates within a review receive a numeric suffix). |
| category | TEXT | – | Check family used by the check selection. |
| severity | TEXT | – | `none`, `info`, `warning`, `error`. |
| status | TEXT | – | `ok`, `variation`, `fail`, `info`, `not_evaluated`. OK rows are logged so that rates have denominators. |
| included | INTEGER | 0/1 | 1 = part of the review/report; 0 = calculated but excluded by the configured check selection. |
| observed_value / expected_value | TEXT | – | Values as produced by the check (text; scrubbed). |
| observed_numeric / expected_numeric | REAL | `unit` | Numeric form if the text is an invariant number and the unit is not `text`/`boolean`; else empty. |
| unit | TEXT | – | Unit of the check values. |
| message | TEXT | – | Check message (scrubbed; empty if `logMessages = false`). |

### 5.5 `ruleset_revision` — one row per configuration document used by a review

| Column | Type | Unit | Definition |
|---|---|---|---|
| review_id | TEXT | – | → `review.review_id`. |
| config_key | TEXT | – | `constraints-excel` or `constraints-refdb` (active source), `structure-aliases`, `default-review-rules`, `plancheck-selection`, `ct-compatibility`, `field-naming`, `mlc-geometry-profiles`, `plancheck-resources`. |
| category | TEXT | – | Document category. |
| revision_number | INTEGER | – | Managed revision number if the active file is a committed snapshot of the configuration history store and its SHA-256 matches; else empty. |
| sha256 | TEXT | – | SHA-256 of the file content actually configured at review time (authoritative). |
| revision_created_utc | TEXT | UTC | Creation time of the managed revision. |
| availability | TEXT | – | `available`, `unavailable` (configured but unreadable), `not_configured`, `not_hashed` (directory or > 64 MiB). |

## 6. SQL import and example views

```bash
python tools/review-log/import_review_log.py --log-dir <ReviewLogDirectory> [--log-dir <second>] --db review_log.sqlite --export-views views/
```

The importer (Python standard library only) creates the schema, checks headers, converts empty fields to `NULL`,
parses numbers strictly, imports idempotently (`INSERT OR IGNORE` on keys) and skips orphan child rows. Views:

| View | Question | Definition |
|---|---|---|
| `v_recurring_findings` | Which checks recur? | check_code × UTC month: `n_evaluated` (included rows, incl. OK), `n_open`, `n_fail`, `open_share`. |
| `v_goal_status_by_table` | Which goals end in Variation/Not met per protocol table? | goal table × source × structure × objective: counts per status; shares relative to evaluated rows (not_evaluated excluded from the denominator). |
| `v_parameter_trend` | Do descriptors drift? | name × scope × month: n, median, nearest-rank quartiles (q1 = rank ⌈n/4⌉, q3 = rank ⌈3n/4⌉), min, max; available values only. |
| `v_rule_change` | Did a rule change reduce findings? | check_code × rule-set epoch with the previous epoch's open share (`LAG`). |
| `v_ruleset_epoch`, `v_ruleset_config` | Which rule set was active when, and which document changed? | First/last use per `ruleset_sha256`; documents per epoch. |
| `v_latest_review_per_plan` | Plan-level analyses without double counting | Latest review per `plan_key` and mode. |

All views exclude `run_status = 'failed'`. Reviews of the same plan in different sessions are separate reviews;
use `v_latest_review_per_plan` for plan-level statistics.

## 7. When rows are written

| Mode | Trigger | report_written |
|---|---|---|
| live | Clinical PDF plan report saved from the review window; otherwise once when the review window closes (current clinical snapshot, if not yet logged). Synthetic/demo snapshots are never logged. | 1 / 0 |
| headless | After the read-only report entry has written PDF/HTML. | 1 |
| retrospective | Each row of an explicit plan list (`ClearPlan.RetrospectiveReviewMining.exe`); no report. Failures produce a `review` row with `run_status = failed`. | 0 |

Logging failures never interrupt the review; the native caller waits at most 3 s (live), 10 s (headless) or 30 s (retrospective).
The legacy `LiveMining.csv` writer is unchanged.

## 8. Known limitations (schema version 1)

* `priority` is reserved and empty. `report_filed` is always 0 (OIS filing is a separate action and not linked yet).
* Review duration, physicist overrides with reason and the link to subsequent plan revisions are not captured (next step: effectiveness study).
* Approval times come from TPS local-time strings; unparsable values stay empty.
* The rule-set fingerprint covers configuration files and the evaluated goal-table definitions, not compiled check logic; `build_id` identifies the code.
