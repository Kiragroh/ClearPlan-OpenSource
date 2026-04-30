# Contributing To ClearPlan

Thank you for helping improve ClearPlan.

This repository is meant to be a starter framework for automated radiation therapy plan review. The most useful contributions are usually the ones that make the project easier to understand, easier to adapt, and safer to extend across different institutions.

## Good First Contributions

- simplify onboarding steps
- improve documentation
- add small starter PQM examples
- generalize checks that are currently too workflow-specific
- improve export, feedback, or changelog usability
- fix bugs that do not depend on one institution's internal infrastructure

## Before You Add New Checks

Please prefer checks that are:

- broadly understandable outside a single clinic
- configurable through templates or settings where possible
- safe when required structures or metadata are missing
- clearly worded in the UI and logs

Please avoid putting the following into the starter path unless they are optional:

- institution-specific naming rules
- assumptions about local network shares
- assumptions about local OIS or SQL access
- hidden dependencies on site-specific files

## PQM Templates

Starter templates live in `ClearPlan.Script/Distribution/ConstraintTemplates`.

If you add or revise a starter template:

- keep it small and readable
- prefer common structures and common DVH objectives
- use aliases where possible
- document any non-obvious assumptions in the CSV comments or in the README

## Settings

Use `settings.json` as the main entry point for configuration.

- Prefer relative paths in the starter release.
- Keep local default folders self-contained.
- If a feature needs a URL, make it optional and provide a local fallback.

## Dependencies

The public starter repository does not track vendor-specific ESAPI assemblies.

If your change touches build dependencies:

- keep redistributable third-party dependencies clearly separated
- do not add proprietary vendor binaries to version control
- update documentation when a manual dependency step is required

## Pull Requests

A strong pull request usually includes:

- a short explanation of what changed
- why the change helps users or developers
- any workflow or compatibility implications
- the build or validation steps used

If you changed plan checks or starter templates, it also helps to explain whether the change is intended as:

- a starter example
- a generally useful default
- or a site-specific extension that should stay optional
