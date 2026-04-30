# ClearPlan

ClearPlan is an open-source starter framework for automated radiation therapy plan review with plan quality metrics, DVH visualization, structured report generation, and failure-mode logging.

This repository has been cleaned up for onboarding and publication. The goal of the starter release is not to ship one clinic's full production rule set, but to give other teams a small, readable base that they can run, understand, and extend.

## What Is Included

- A simplified starter `ErrorCalculator` with general examples instead of clinic-specific nomenclature and network-dependent checks
- Three small starter PQM tables for conventional plans, hypofractionated plans, and plan sums
- A readable `settings.json` with local default paths
- Working Help menu entries for Change Log and Feedback
- Local default export folders for logs, reports, CSV exports, and app state

## Quick Start

1. Add institution-approved copies of `VMS.TPS.Common.Model.API.dll` and `VMS.TPS.Common.Model.Types.dll` to `ClearPlan-DLLs`.
2. Build `ClearPlan.sln` with `Debug|x64`.
3. Open `built/settings.json`.
4. Adjust paths or add a public feedback URL if needed.
5. Start `built/ClearPlan.Runner.exe` or load the built plugin in your ESAPI environment.

After the first build, the output folder contains:

- `built/ConstraintTemplates`: starter PQM tables
- `built/Logs`: usage logs and comparison logs
- `built/Reports`: generated PDF output
- `built/Exports`: CSV exports
- `built/State`: local app state such as version tracking

## Configuration

`settings.json` is the main configuration entry point. It is copied into the build output automatically.

Notes:

- Relative paths are resolved from the built application folder.
- Fields starting with `_comment` are only there to help humans and are ignored by the loader.
- Legacy `settings.ini` files are still supported as a fallback.

Important path settings:

- `constraintTemplatesDirectory`: folder containing PQM CSV files
- `defaultConventionalTemplate`: default starter template for standard fractionation
- `defaultHypofractionatedTemplate`: default starter template for high dose per fraction
- `defaultPlanSumTemplate`: default starter template for plan sums
- `logsDirectory`: diagnostic and comparison logs
- `reportsDirectory`: PDF report output
- `csvExportDirectory`: manual CSV export output
- `stateDirectory`: local state files
- `usageLogFile`: plan-selection usage log
- `activityLogFile`: structured activity/failure log
- `versionSeenUsersFile`: local version-tracking file for changelog popups
- `changeLogFile`: local markdown file opened from Help
- `feedbackFile`: local markdown fallback for feedback guidance

The optional `links.feedbackUrl` setting can point to a GitHub Issues page or another public feedback form. If it is left blank, ClearPlan opens `FEEDBACK.md` locally.

## Starter Templates

The repository ships these example CSVs:

- `Starter_Conventional.csv`
- `Starter_Hypofractionated.csv`
- `Starter_PlanSum.csv`

They are intentionally small. They are meant to demonstrate:

- structure aliasing
- common DVH objective patterns
- a clean table layout for extension

You can keep your own larger institutional tables in the same folder. The starter templates are simply preferred defaults for first-time users.

## Starter Plan Checks

The default starter checks focus on broadly understandable examples such as:

- missing or invalid dose
- missing treatment beams
- unapproved plans
- missing target structures
- missing or inconsistent primary reference points
- treatment-unit consistency
- dose-grid and image-resolution warnings for high-dose fractions
- basic plan-sum consistency checks

The open-source starter deliberately does not include clinic-specific nomenclature audits, internal network-share mining, or local workflow assumptions.

## Help Menu

The Help menu now works as a self-contained onboarding feature:

- `Change Log` opens the packaged changelog content
- `Feedback` opens a configured URL if present, otherwise the local `FEEDBACK.md`

This makes the release easier to distribute without wiring it to internal infrastructure.

## Repository Layout

- `ClearPlan.Script`: main ESAPI plugin and UI
- `ClearPlan.Reporting`: report data structures
- `ClearPlan.Reporting.MigraDoc`: PDF rendering
- `ClearPlan.Runner`: desktop launcher
- `ClearPlan-DLLs`: local third-party dependencies used by the starter build

Starter distribution files are maintained in `ClearPlan.Script/Distribution` and copied to `built/` during compilation.

## Build And Dependency Notes

This repository is prepared for public sharing, but external users still need an ESAPI-capable environment and must verify whether they are allowed to use or redistribute vendor-specific assemblies in their institution.

The public starter repository does not track these vendor-specific ESAPI files:

- `ClearPlan-DLLs/VMS.TPS.Common.Model.API.dll`
- `ClearPlan-DLLs/VMS.TPS.Common.Model.Types.dll`

These must be supplied locally before building.

## Community

- Use GitHub Issues for bugs, onboarding problems, and feature requests.
- Use the templates in `.github/ISSUE_TEMPLATE` to keep reports structured and anonymized.
- See `CONTRIBUTING.md` for guidance on starter-friendly changes.

For publication and collaboration, the important idea is:

- ClearPlan is a framework
- the starter package is intentionally lean
- production rule sets should be added by each institution on top of this base
