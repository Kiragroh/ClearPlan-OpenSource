# Proposed local commissioning scaffold

[Documentation](index.md) · [Configuration guide](configuration-guide.md)

This checklist is a transfer aid for medical-physics groups adapting ClearPlan
or its snapshot contract. It is not a validated universal commissioning
protocol, a substitute for local risk analysis, or evidence that the public
examples are clinically suitable.

Current development source is separate from the archived v3.2.0 release. Use the actual
evaluated source/artifact manifests and execution receipts, not an old release's
test totals, environment inventory or sample hashes. The
[reproducibility guide](reproducibility.md) describes public software tests;
the clinical acceptance activities below are additional local work.

## 1. Freeze the evaluated system

- Record ClearPlan tag and commit, snapshot schema, TPS and scripting-API
  versions, licensed dependencies, operating system, and local catalog version.
- Keep the evaluated binaries, settings, rule sources, checksums, build log,
  rollback package, and approval record together.
- Identify the exact source and dependency closure for the Eclipse build,
  including any independently governed clinical PlanCheck implementation.
  Public starter checks are not a substitute for that local rule set.
- Record tool/runtime versions from the evaluated environment. Verify local and
  server clocks when timestamp-dependent interfaces are in scope; do not treat
  a fixture timestamp or an earlier workstation record as that verification.

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
- Check all PTVs and the selected largest contained inner target against native
  geometry and prescription assignments. Review the two Paddick CI conventions,
  GI/HI definitions, target coverage and treatment of missing dose/volume inputs.
- Commission the exact MLC model/profile, physical leaf boundaries, layer/index
  mapping, jaws or fixed jawless limits, collimator frame and isocenter scale.
  For dual-layer MLCs, verify the intersection of both physical openings; neither
  a missing layer nor an inferred jaw is an acceptable substitute.
- Independently check target projections and PAM for representative control
  points and arcs. Automatic lowest-PAM selection must disclose eligible/valid
  candidates and must not be interpreted as clinical superiority. Verify explicit
  overrides and the unavailable state when no complete candidate exists.
- Distinguish the nominal MU/min field setting, separately supplied planned
  values and the dashed estimated segment trajectory. Review the configured
  speed assumption and model omissions. A variable estimate is not a measured
  delivery trace; use appropriate delivery evidence if that is the intended claim.

## 4. Validate rules and mappings

- Review every local constraint, comparator, unit, alias, laterality rule,
  structure code, table-selection rule, and field-naming rule.
- Demonstrate equivalent internal catalogs from approved RefDB JSON and Excel
  inputs where both source routes are intended.
- Test missing, duplicate, malformed, ambiguous, and version-incompatible
  source content.
- Review the editable Stock 2024 compilation as an example configuration, not a
  clinical recommendation. Its current public inventory is eight tables,
  495 executable rules and 70 canonical structures; inspect the exact locally
  adopted workbook and provenance rather than accepting it by row count.
- Validate configuration history, actor/time attribution, change reasons and
  restore-as-new-revision. Protect these records locally and re-evaluate the
  affected rules after any configuration change.

## 5. Exercise known-good and altered controls

- Use approved test-patient or phantom plans representing normal workflows,
  boundary cases, uncommon techniques, and intentionally altered conditions.
- Compare every displayed and reported output with an independently approved
  expectation; do not use the public synthetic constraints as clinical goals.
- Record false flags, missed conditions, unavailable checks, and unresolved
  mappings.
- Keep the seven public baseline/fault scenarios and two coherent publication
  phantoms as software regression controls. Their analytic dose is not calculated
  from the apertures, so they cannot replace physical or clinical validation.

## 6. Test failure and data-flow behavior

- Exercise unavailable and slow network sources, invalid settings, read-only
  permissions, report-output failures, refresh failures, and rollback.
- Verify the complete TPS-to-OIS/treatment-unit chain separately where it is in
  local scope; ClearPlan does not perform that verification.
- Confirm privacy controls and approved storage for logs, reports, and any
  adapter output containing confidential dosimetry.
- Exercise cross-tab DVH/image visibility, hidden unmatched rows, missing metrics
  and report options. Confirm that a compact or filtered report discloses its
  scope and does not turn omitted or unavailable evidence into a pass.
- Treat every native snapshot and report as confidential, including pseudonymous
  files. Public examples may use synthetic data or patient-free configuration
  screens; do not substitute masked native patient images. Any clinical publication
  material needs its own institutional privacy, ethics, and release review.

### Optional ARIA document transfer

This is a separate, explicitly confirmed write to the patient record, not a plan
modification. Leave it disabled unless the institution intends and approves the
integration. Verify the [ARIA configuration and interface contract](ARIA_REPORT_UPLOAD.md)
against the installed server profile and use authorized test records first.

- Check exact patient/provider binding, document type/category, preliminary
  status, report date and captured creation instant. A configured date margin
  is a disclosed compatibility setting, not clock synchronization.
- Verify certificate trust or explicitly approved exact certificate bindings,
  credential protection, no redirects and no automatic retry of a document POST.
- Check full and sparse readback responses. When FHIR returns only a stored PDF
  path, configure only approved `AttachmentReadbackRoots`; an empty list must
  disable filesystem readback. Reject out-of-root, traversal, local/device and
  redirected/reparse paths. Verify least-privilege read access and unavailable
  behavior without broadening the roots or writing to the file server.
- Require byte length and SHA-256 agreement with the prepared PDF for the
  filesystem route. Missing optional metadata must be disclosed as unchecked;
  metadata-only agreement is not proof of stored report integrity.
- Test rejected, timed-out and uncertain outcomes, durable receipts across
  restart, and GET-only reconciliation of an existing resource. Reconcile with
  ARIA before another explicitly authorized send; do not delete receipts merely
  to force a retry.

Keep local transfer evidence and receipts on approved protected storage. Synthetic
transport tests and source-level UI guards do not establish live interoperability.

## 7. Complete user acceptance

- Have physicists review the overview, focused pages, DVH behavior, reports,
  warning burden, and required human actions.
- Compare the GUI, PDF and HTML for the same active plan, selected structures,
  three-plane dose/contour overlays and optional exact field-start BEV/DRR. Verify
  that loading a comparison reference does not replace the active-plan report.
- Train users that pass/variation/fail are row states, never plan approval.
- Obtain independent clinical, information-security, and governance approval.

## 8. Control deployment and surveillance

- Deploy through a controlled Release build with documented permissions,
  backup, rollback, and version ownership.
- Re-run the regression and local acceptance sets after changes to the TPS,
  API, operating system, ClearPlan, settings, rule catalog, or report path.
- Review alerts, overrides, false and missed findings, user feedback, and
  incident-learning data periodically; suspend use when stop criteria are met.
- Keep public software test outcomes, approved clinical acceptance and production
  deployment authorization as separate records. None substitutes for the others.
