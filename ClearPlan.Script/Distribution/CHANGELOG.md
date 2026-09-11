# Changelog

## [3.2.0.0] Integrated plan analysis and publication fixtures (unreleased candidate)
- Expanded the read-only review with target-specific PAM, physical single-/jawless dual-layer MLC geometry, Paddick CI and reciprocal, GI/HI, total MU and distinctly labeled estimated dose-rate trajectories. Automatic PAM compares eligible PTVs without changing TPS assignments.
- Added three-plane CT dose/structure overlays, optional field-start BEV/DRR report panels, compact single-plan PDF and offline HTML quicklook, separate comparison-plan selection and shared structure visibility.
- Improved focused DVH sizing, outside legends, hover contrast and reset behavior; unmatched rows are hidden by default but remain editable in structure mapping.
- Added versioned settings/configuration history for aliases, default constraints, field nomenclature, target rules and enabled checks, with safe restore-as-new-revision and explicit missing/ambiguous states.
- Added two patient-free publication fixtures using coherent analytic phantom definitions, while retaining the original seven scenarios and ESAPI manuscript sandbox. Phantom dose is not calculated from apertures and does not establish physical or clinical validation.
- Added optional, disabled-by-default ARIA report-document upload with patient/provider binding, durable audit receipts and verification; treatment-plan inspection remains read-only.
- Expanded portable regression coverage and release evidence collection to record actual executions, tested-input hashes and explicit stock-source skips. Final release, manuscript and clinical-commissioning gates remain pending; no test-count claim is inferred from registrations.

## [3.1.0.0] Simulator-backed technical-note release (GitHub v3.1.0)
- Added the standalone, vendor-free ClearPlan Simulator with seven deterministic synthetic scenarios, watermarked PDF reports, DVH export, reproducible captures, and no clinical source data.
- Added the versioned neutral `ReviewSnapshot` contract shared by the simulator and the read-only Eclipse/ESAPI host.
- Unified the clinical overview, focused PQM/PlanCheck/field/DVH pages, structure mappings, field-ID/name conformance, source status, and report action on the same validated review model.
- Added an integrated `mixed-review` publication scenario with analytically generated DVHs and explicit injected PlanCheck, field-name, mapping, and optional-source deviations.
- Bounded optional network source loading, retained the last valid catalog/workspace, and made PQM mapping and constraint-table changes transactional.
- Added an illustrative, unvalidated, read-only RayStation snapshot adapter with 20 Python safety and contract tests; it is not a commissioned clinical implementation.
- Expanded the C# executable harness to 118 tests and added simulator, path-timeout, transaction, shared-workspace, privacy, and release validators.
- Added the reproducible ZMP Short Communication manuscript, three patient-free figures, machine-readable evidence, and public v3.1.0 citation metadata.

## [3.0.0.4] Report and field-status clarity
- Added the real PDF report action directly to the integrated overall review.
- Split field conformance into visible `ID okay`/`ID ändern` and independent name status columns.
- Made the field-name token order explicit as angle, optional table angle, then UZ/GUZ direction.
- Added regression coverage for duplicate arc names with a non-zero table angle.
- Expanded the source-independent harness to 55 tests.

## [3.0.0.3] Integrated review surface
- Replaced the card-only quick overview with a vertically scrollable overall review containing plans, mapped PQMs, PlanCheck findings, field conformance, and a wide DVH.
- Made the resolved PQM structure visible as durable text in both overall and detail views.
- Added plan-aware expected field IDs while retaining the PlanFieldNamer-compatible name rules and reporting ID/name conformance separately.
- Moved DVH default selection out of realized WPF checkbox containers and synchronized separate overall/detail plot models.
- Placed the DVH legend inside the plot and removed the saturated blue navigation-selection banner.
- Expanded the source-independent harness to 54 tests.

## [3.0.0.2] PQM review stability
- Replaced visual-tree checkbox indexing with observable PQM view-model state during initialization, reporting, and CSV export.
- Prevented unrealized or filtered DataGrid rows from causing an `ArgumentOutOfRangeException` while opening a plan.
- Restricted the runner's path-failure classification to actual I/O, access, path-format, and security failures.
- Added a clinical build override and provenance manifest so the deployed DLL is demonstrably compiled from the preserved clinical PlanCheck source.
- Added a runtime-file lock preflight so a running shared Runner blocks deployment before any copy starts.
- Added regression coverage for wrapped I/O failures, collection-index failures, and unsafe PQM checkbox access.
- Expanded the executable harness to 48 tests and added a PQM review safety validator.

## [3.0.0.1] Central path settings and runner resilience
- Moved every file-system path into one editable `settings.ini`; `settings.json` now contains non-path options only.
- Exposed all RefDB, Excel, template, output, log, state, support-file, prescription, and PlanCheck resource paths on the ClearPlan settings page.
- Removed the obsolete `variancom` default from the clinical configuration and added deployment validation against its reintroduction.
- Added local writable fallbacks for optional logs, reports, exports, and state files when a configured network path is unavailable.
- Prevented optional plan-selection logging failures from closing ClearPlan.
- Added runner-level controlled error reporting with a local `Logs\RunnerErrors.log`.
- Expanded the executable harness to 46 tests and added runner/path resilience validation.

## [3.0.0.0] Constraint sources and Clinical Blueprint
- Added a source-independent constraint catalog with read-only RefDB JSON and Excel XLSX adapters.
- Added automatic RefDB-first loading with a visible Excel fallback plus strict explicit source modes.
- Added deterministic structure alias and laterality resolution with ambiguity reporting.
- Added plan-context constraint-table selection with a documented reason and confirmation for uncertain results.
- Replaced active CSV defaults with `ClearPlan_DefaultConstraints.xlsx`; legacy CSV files remain migration inputs only.
- Added a read-only PlanFieldNamer-compatible field-name page and executable examples for static, arc, couch-angle, UZ/GUZ, ordering, tolerance, and duplicate cases.
- Added the Clinical Blueprint navigation with a combined quick overview and complete PQM, PlanCheck, field-name, DVH, and settings pages.
- Added editable settings for source and operational paths with validate-before-save behavior.
- Styled the existing EsapiEssentials desktop runner with fail-open Clinical Blueprint resources without replacing its ESAPI lifecycle.
- Added deployment safeguards that preserve an independently governed clinical PlanCheck.
- Expanded build, migration, Office-content, UI-distinctiveness, and read-only validation.

## [2.9.3.1] Open-source starter refresh
- Added a human-readable `settings.json` with local default paths for logs, reports, exports, state files, changelog, and feedback guidance.
- Simplified the default `ErrorCalculator` profile to a small set of general starter checks instead of site-specific nomenclature and network-dependent rules.
- Added three lean starter PQM tables for conventional plans, hypofractionated plans, and plan sums.
- Reworked the Help menu so Change Log and Feedback open dedicated content instead of reusing the anonymize handler.
- Added privacy-safe defaults for logs, generated filenames, and local version tracking.
- Replaced site-specific user-domain handling with configurable identifier handling.
- Removed hard dependencies on clinic-specific network shares for the common onboarding workflows.

## [Compatibility notes]
- Existing path keys and legacy aliases in `[Paths]` are imported into the authoritative `settings.ini` model.
- Non-path options continue to be read from `settings.json`.
- Existing site-specific PQM tables can stay in `ConstraintTemplates`; the starter templates are just preferred defaults.
