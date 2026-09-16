# ClearPlan: receiving-clinic onboarding

[Documentation](index.md) · [Configuration guide](configuration-guide.md)

This development snapshot is not an approved clinical installation. The archived
v3.2.0 release and its evidence describe an earlier source state.
Record named local owners: MPE lead for acceptance, TPS administrator for
deployment, information-security contact for storage/access, and configuration
custodian for rules and rollback.

## 1. Choose and freeze the environment

The free, vendor-independent simulator needs no Eclipse, ESAPI or patient data;
training and software regression do not commission clinical semantics. Eclipse
mode requires licensed, compatible ESAPI assemblies and an approved runtime;
these assemblies are not redistributed. Follow the [build instructions](../README.md).
Record commit, binary hashes, TPS/API versions, dependencies and configuration
revisions. Retain the previous working package.

Starter PlanCheck and Stock 2024 are examples, not commissioned policy. Review
every adopted rule, comparator, unit, fraction range, scope and alias against
independently approved expectations. Table confirmation is not clinical
authorization. RayStation support is an illustrative, unvalidated snapshot adapter.

## 2. Set paths and activate controlled revisions

Clinical-build paths are in `built\settings.ini`, section `[Paths]`; non-path
options are in `built\settings.json`. Relative paths resolve from the loaded
ClearPlan assembly/settings directory, not the shell directory. Distribution
files are templates. Open **Settings → Paths & sources** in the English display.

| Setting | Purpose / distributed default |
|---|---|
| `RefDbJsonPath`, `ExcelWorkbookPath` | Blank RefDB; `ConstraintTemplates/ClearPlan_StockConstraints2024.xlsx` |
| `StructureAliasesJsonPath` | Blank: retain the chosen source's mappings |
| `ConfigurationDirectory` | `Configuration`: managed revisions |
| `MlcGeometryProfilesJsonPath` | `MachineGeometry\MlcGeometryProfiles.example.json` |
| `DoseRateProfilesJsonPath` | `MachineGeometry\DoseRateProfiles.example.json` |
| `CollisionProfilesJsonPath`, `SourceCollisionModelsJsonPath` | Blank: no supplied active collision configuration |
| `ReportsDirectory`, `CsvExportDirectory`, `LogsDirectory`, `StateDirectory` | `Reports`, `Exports`, `Logs`, `State` |
| `AriaUploadConfigJsonPath` | Blank: upload disabled |

Select the source mode, run **Test source**, and inspect path/fallback. Under
**Configurations & versions**, select a document, enter a
**Reason for change**, then use **Import source** or **Import another file …**.
Import validates a copy, preserving the source. JSON edits use
**Validate & save as version**; Excel uses **Open Excel draft**, then
**Apply changes** after saving and closing the separate draft.

Review version, UTC time, recorded Windows actor and SHA-256.
Only **Save settings** activates the candidate path; rebuild the review
afterward. **Restore selected version …** requires a reason and appends
a revision; save Settings to activate it. Back up the whole managed store.
Actor names and hashes are neither authenticated attribution nor tamper-proof
auditing. Stop on integrity errors; never repair checksums manually.
See [history behavior](CONFIGURATION_HISTORY_STORE.md).

## 3. Commission exact geometry and semantics

Validate exact machine identifiers, MLC model strings, leaf boundaries,
bank/layer indexing, jawless limits, both dual-layer openings, collimator rotation,
coordinates and isocenter scale. Unknown profiles remain unavailable; do not
guess. Collision envelopes require local measurements, coverage/accessory checks
and evidence, not merely a JSON commissioning flag. Source-model point screening
is not surface clearance. Validate dose-rate assumptions; estimates are not
measurements. See the [commissioning scaffold](local-commissioning-scaffold.md)
and [collision limitations](collision-preview.md).

## 4. Run a recorded, read-only acceptance matrix

Use authorized test-patient/phantom cases and detached synthetic faults; do not
alter clinical plans. Record expected/observed GUI, PDF and HTML behavior,
reviewer, version and discrepancies.

| Case | Required observation |
|---|---|
| Valid, independently reviewed plan | Correct patient/course/plan, source, fractions, mappings, DVHs and metrics; no TPS writes |
| Missing dose | Missing measurements stay unavailable, never a pass or invented zero |
| Missing constraint source/table | The missing source or table remains explicit; no unsupported constraint evaluation or silent pass |
| Missing support/table geometry; truncated CT | Absent support or incomplete anatomy cannot establish collision clearance |
| Stale context / plan switch | No previous-plan results or images presented as current; verify export context |
| Malformed or cross-side alias | Explicit configuration problem; no silent source change hiding the error |
| Optional UNC source/output paths | Test delay, denial and outage; inspect visible source/output fallback and actual protected destination |
| Single-/dual-layer synthetic baseline and faults | Exercise `publication-single-layer`, `publication-dual-layer` and baseline/fault scenarios; analytic dose is not aperture-derived validation |
| Selection and comparison | Shared DVH/CT visibility and hidden counts match reports; reference loading leaves active-plan exports unchanged |

## 5. Protect outputs and stage adoption

Treat reports, screenshots, logs, snapshots and pseudonymous outputs as
confidential. Name/ID masking is not full anonymization. Keep clinical data,
credentials and site-specific paths out of public artifacts. Reopen exports to
verify identity, scope and destination.

Leave optional ARIA upload disabled during read-only acceptance: this separate
patient-record write requires local approval and integration testing;
see [ARIA report upload](ARIA_REPORT_UPLOAD.md).

For optional document integration, ask Varian/MyVarian support to provision and
configure FHIR/VAIS if it is not available. Create and commission a local upload
profile, including the personal ESAPI-user/Practitioner mapping. Clinics may
instead use their configured Webservice/direct-upload workflow or an eDocPrinter
profile with exported reports; these routes are not automatic fallbacks of the
FHIR button. See [document routes and first-time setup](ARIA_REPORT_UPLOAD.md#first-time-setup-choose-and-commission-a-document-route).

Stage simulator training, protected read-only evaluation, supervised pilot, then
authorized routine use. The named MPE lead signs acceptance; the TPS administrator
deploys the evaluated package. Define stop-use/escalation criteria and rehearse
configuration restore and application rollback. Re-test after software, TPS,
profile, rule or path changes. Software tests, clinical acceptance and deployment
authorization remain separate records.
