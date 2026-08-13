# Proposed local commissioning scaffold

This checklist is a transfer aid for medical-physics groups adapting ClearPlan
or its snapshot contract. It is not a validated universal commissioning
protocol, a substitute for local risk analysis, or evidence that the public
examples are clinically suitable.

## 1. Freeze the evaluated system

- Record ClearPlan tag and commit, snapshot schema, TPS and scripting-API
  versions, licensed dependencies, operating system, and local catalog version.
- Keep the evaluated binaries, settings, rule sources, checksums, build log,
  rollback package, and approval record together.

## 2. Map local requirements and risks

- Map intended checks to the local process, TG-275 failure modes, applicable
  data-transfer paths, and the institution's FMEA or equivalent risk analysis.
- Mark every complete, partial, manual, and out-of-scope review item.
- Define row-state acknowledgment, escalation, override, audit, and stop-use
  criteria before testing.

## 3. Validate acquisition and semantics

- Verify API/runtime compatibility and plan, PlanSum, beam-set, fractionation,
  prescription, unit, coordinate, and dose-query semantics.
- Compare TPS-derived DVHs and metrics with independently approved results,
  including endpoint, plateau, rounding, and goal/variation boundaries.
- Confirm that unsupported or ambiguous states remain visible and unevaluated.

## 4. Validate rules and mappings

- Review every local constraint, comparator, unit, alias, laterality rule,
  structure code, table-selection rule, and field-naming rule.
- Demonstrate equivalent internal catalogs from approved RefDB JSON and Excel
  inputs where both source routes are intended.
- Test missing, duplicate, malformed, ambiguous, and version-incompatible
  source content.

## 5. Exercise known-good and altered controls

- Use approved test-patient or phantom plans representing normal workflows,
  boundary cases, uncommon techniques, and intentionally altered conditions.
- Compare every displayed and reported output with an independently approved
  expectation; do not use the public synthetic constraints as clinical goals.
- Record false flags, missed conditions, unavailable checks, and unresolved
  mappings.

## 6. Test failure and data-flow behavior

- Exercise unavailable and slow network sources, invalid settings, read-only
  permissions, report-output failures, refresh failures, and rollback.
- Verify the complete TPS-to-OIS/treatment-unit chain separately where it is in
  local scope; ClearPlan does not perform that verification.
- Confirm privacy controls and approved storage for logs, reports, and any
  adapter output containing confidential dosimetry.

## 7. Complete user acceptance

- Have physicists review the overview, focused pages, DVH behavior, reports,
  warning burden, and required human actions.
- Train users that pass/variation/fail are row states, never plan approval.
- Obtain independent clinical, information-security, and governance approval.

## 8. Control deployment and surveillance

- Deploy through a controlled Release build with documented permissions,
  backup, rollback, and version ownership.
- Re-run the regression and local acceptance sets after changes to the TPS,
  API, operating system, ClearPlan, settings, rule catalog, or report path.
- Review alerts, overrides, false and missed findings, user feedback, and
  incident-learning data periodically; suspend use when stop criteria are met.
