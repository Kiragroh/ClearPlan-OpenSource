# Review controls

[Documentation](../index.md) · [Configuration guide](../configuration-guide.md)

Source-level behavior as of 2026-09-11 for the existing native Clinical Blueprint
workspace. This is a scoped operation guide, not a replacement for the root design
system or a native acceptance/deployment record. All plan inspection remains
read-only; view selection is not treatment approval.

## Structure selection and report scope

The initial DVH selection combines resolved goal structures, native Eclipse DVH
selection, manual ClearPlan selections, all nonempty PTVs, and at most one largest
verified contained CTV/GTV/ITV overall. Unknown containment cannot establish that
an inner target is the largest. The full structure inventory stays available.

Use individual DVH or CT legend checkboxes to hide or show a structure. Targets
are default-selected, not locked. This session-local visibility is shared by the
DVH and CT views and survives tab changes and same-plan snapshot replacement; it
does not change native Eclipse selections, cached measurements or mapping options.
**Reset DVH** restores the captured initial selection. A different active
plan or synthetic/clinical mode does not inherit the previous plan's view state.

HTML/PDF export follows the current DVH/CT structure visibility, including initially
unselected structures and subsequent individual legend changes.
**Hide unmatched criteria (GUI + report)** is enabled by default in
new workspaces and direct PDF/HTML exports. It hides unresolved goals and
summary mappings, not matched rows whose measurement is unavailable. Hidden-row
counts remain visible and are not counted as passed. The PQM mapping editor stays
complete so an unresolved row can still be assigned.
Uncheck the option to include unmatched rows in the current GUI/report. Same-plan
refresh preserves that explicit choice; it does not alter the underlying review.

## Dose-rate presentation

The native ESAPI 16.1 adapter reads `Beam.DoseRate` as a single nominal field
setting. `ControlPoint.MetersetWeight` is cumulative MU weighting, not time.
The field setting remains a labeled scalar; it is never repeated as a substitute
rate curve. A separately labeled PlanCheck-style **estimated plan trajectory**
uses consecutive control-point MU changes, directed gantry-angle changes and an
explicit configured speed assumption. Each estimated segment duration is the
larger of its angle/speed time and its MU/nominal-rate time. This is a model, not
time-resolved information supplied by ESAPI or a measured delivery trace.

`DoseRateProfilesJsonPath` in `settings.ini` selects the editable JSON profile
catalog; **Dose rate · Estimation model** in Settings exposes its validated,
versioned configuration workflow. Matching requires the configured exact native
identifiers; missing or ambiguous profiles leave the estimate unavailable. A
blank configured path disables estimation. Verify the chosen speed assumption
against the local treatment system before interpreting the model.

GUI, PDF and HTML distinguish dashed `EstimatedDoseRateMuPerMin` samples from
solid, independently supplied `PlannedDoseRateMuPerMin` samples. Native ESAPI does
not fill the latter field with estimates. Profile, assumptions and availability
are disclosed alongside the plot; the horizontal axis is control-point index,
not measured time. The first control point has no preceding segment rate.
Missing or invalid values remain gaps; absent traces remain unavailable without
a nominal horizontal fallback. The publication simulator exercises the same
estimator with explicitly synthetic profiles and changing cumulative MU weights.

The estimate omits acceleration, MLC/jaw motion (including both dual-layer MLC
layers), dose ramping and beam holds. See the [dose-rate method](../DOSE_RATE_ESTIMATE.md)
for the formula, input checks, configuration and model limitations.

RapidArc can modulate dose rate and gantry speed, but this does not imply that
each individual arc must have a varying rate. An absent curve here is a
data-availability limit, not a finding about the delivered arc. Time-resolved
delivery/trajectory logs would be needed for a measured-rate trace.

Sources: [ESAPI Beam](https://docs.developer.varian.com/api/16.1/VMS.TPS.Common.Model.API.Beam.html),
[ESAPI ControlPoint](https://docs.developer.varian.com/api/16.1/VMS.TPS.Common.Model.API.ControlPoint.html),
[Ling et al., 2008](https://pubmed.ncbi.nlm.nih.gov/18793960/).

## Report and comparison options

**Include beam-start BEV / DRR in report** controls optional field-start panels in HTML and
PDF. Enabled panels use the same captured exact CP 0 image data; missing images
remain labeled unavailable. A field-start aperture is not arc-integrated fluence, and a CT-derived
DRR is an attenuation proxy, not a diagnostic image or independent dose check.
Goal tables use compact columns while retaining the comparator and units beside
numeric values. PlanCheck leads with message and status rather than repeated
observed/expected/unit columns.

Under plan comparison, select a plan and use **Load as reference**. The native host
loads a separate reference snapshot within the current patient while the active
plan and its review state remain unchanged. This is distinct from opening another
plan. HTML/PDF contain the current plan only; they do not export the comparison
reference or its deltas. Clinical and synthetic contexts are not interchangeable.

## Target rules and PAM

The shipped [DefaultReviewRules.json](../../ClearPlan.Script/Distribution/DefaultReviewRules.json)
uses strict `D98% > 95%` of assigned Rx for PTVs and `D98% > 100%` for CTV/GTV/ITV.
Equality fails. Each evaluated target requires exactly one matching native
`RTPrescription.Targets` volume assignment, explicit Gy/cGy dose and the same
positive fraction count as the active plan. No plan-total, containing-PTV or
partial-course Rx is substituted. Missing or ambiguous Rx stays NotEvaluated.

`DefaultReviewRulesJsonPath` remains editable through Settings. Existing external
or custom JSON is not migrated to the shipped values. Disabled, deleted or absent
rules are not recreated during review; malformed rules do not trigger a fallback.
Rebuild the review after saving a configuration change.

Automatic PAM calculates every eligible nonempty, segmented PTV and selects the
lowest complete finite available value. Partial, missing and nonfinite results
are excluded rather than treated as zero. The analysis retains per-candidate
availability/results and valid/eligible counts; review selection provenance names
the policy and its coverage. PDF and HTML also show the selected target and the
valid/eligible candidate count, with a compact automatic/explicit selection label.
A numerical minimum is not a clinical superiority claim. Equal values resolve deterministically by target ID. With no PTV candidates,
an eligible native plan target can still be used.

An explicit PAM target overrides automatic selection. An invalid explicit target
remains unresolved rather than silently falling back. If candidates exist but none
has complete finite PAM, automatic PAM remains unavailable. No target assignment
is written back to the TPS.

## PlanCheck inclusion settings

**Settings > PlanCheck selection** lists check families found in the current review or
already present in the selected configuration; it is not a complete algorithm
catalogue of the external engine. Repeated findings such as `74-2` belong to the
base family `74`. Native warnings, reference-point summaries and generated per-run
fallback IDs cannot be disabled through this policy.

The separate `PlanCheckSelectionJsonPath` file uses this schema:

```json
{
  "schemaVersion": 1,
  "checks": [
    { "code": "74", "enabled": false }
  ]
}
```

This example excludes family `74` from review/report; it does **not** prevent the
legacy calculator from executing. Unlisted families remain enabled. Only stable
codes and enabled flags are saved; descriptions shown in Settings stay in memory.

To persist a selection, use **Apply as versioned draft**, supply a change
reason under **Configurations & versions**, save the version, then save Settings.
The policy applies at the next review build. A blank/missing file retains all
findings without creating a file. Invalid, unreadable or unsupported JSON also
retains all findings and adds a configuration warning. Native Eclipse warnings
remain unaffected. Review and reports disclose the number of excluded findings;
this is not a complete all-checks pass.

## Source and verification boundary

The behavior above is grounded in `ReviewWorkspaceViewModel`,
`ClinicalReviewWorkspaceHost`, `DvhSelectionPolicy`, `UserTargetCoverageRule`,
`PamTargetSelection`, `PlanCheckSelectionConfiguration`, `SettingsViewModel`,
`ReviewReportDocument` and the HTML/PDF renderers. Relevant regression sources
include `ReviewControlTests`, `PamBestTargetTests`, `PlanCheckSelectionTests` and
`ReportPresentationTests`; naming these tests does not claim a fresh passing run.

See the public [plan-analysis method](../plan-analysis-method.md) and
[dose-rate method](../DOSE_RATE_ESTIMATE.md) for extraction boundaries,
approximations and unavailable states. Portable regression tests and synthetic
examples do not establish native clinical validation or deployment. Local
commissioning still requires controlled, read-only comparison with the treatment
planning system. Keep native images, reports and clinical evidence on protected
institutional storage; public repository figures and fixtures remain patient-free (synthetic data or
public-only configuration screens), not masked native patient captures.
