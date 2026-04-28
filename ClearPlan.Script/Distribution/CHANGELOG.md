# Changelog

## [2.9.3.1] Open-source starter refresh
- Added a human-readable `settings.json` with local default paths for logs, reports, exports, state files, changelog, and feedback guidance.
- Simplified the default `ErrorCalculator` profile to a small set of general starter checks instead of site-specific nomenclature and network-dependent rules.
- Added three lean starter PQM tables for conventional plans, hypofractionated plans, and plan sums.
- Reworked the Help menu so Change Log and Feedback open dedicated content instead of reusing the anonymize handler.
- Removed hard dependencies on clinic-specific network shares for the common onboarding workflows.

## [Compatibility notes]
- Existing `settings.ini` files are still read as a fallback when `settings.json` is missing.
- Existing site-specific PQM tables can stay in `ConstraintTemplates`; the starter templates are just preferred defaults.
