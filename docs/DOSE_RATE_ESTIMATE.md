# Estimated plan dose-rate trajectory

[Documentation](index.md) · [Configuration guide](configuration-guide.md)

This read-only development feature is a PlanCheck-style estimate, **not** a measured delivery trace or a time-resolved rate supplied by Eclipse. It does not change a plan, dose, field or machine setting.

For each consecutive control-point interval:

```
deltaMU = beamMU * (CMW[i] - CMW[i-1]) / finalCMW
estimatedSeconds = max(directedAngle / configuredGantrySpeed,
                       60 * deltaMU / nominalMUperMinute)
estimatedMUperMinute = 60 * deltaMU / estimatedSeconds
```

The nominal field setting is a model cap, never copied into a curve as a substitute for calculation. A constant estimate is possible if every interval is MU-limited. Variation is not manufactured. The first control point has no interval value. Moving zero-MU intervals yield zero; a zero-time duplicate has no rate. Missing profiles, ambiguous matches, nonfinite/invalid MU or weights, incomplete indices, unknown direction, or reversed/ambiguous angular intervals leave the whole beam estimate unavailable. Intervals of 180 degrees or more are rejected rather than guessed.

The model omits gantry acceleration, MLC and jaw dynamics (including both Halcyon layers), dose-rate ramping and beam holds. It provides optimistic kinematic segment timing, not a prediction validated against machine delivery. Dual-layer geometry continues to be used independently for aperture/PAM. MU/min is not absorbed-dose rate in Gy/s.

## Configuration and identity

`settings.ini` contains `DoseRateProfilesJsonPath`, pointing to JSON. An explicitly blank path disables estimation. In Settings the same file is accessible as **Dose rate · Estimation model**, with the existing validated import, change reason, author/time/hash history and restore-as-new-revision workflow. Loading is asynchronous, bounded to 64 KiB and ten seconds; optional profile failure does not erase other plan metrics. Refresh the active review after configuration changes.

Profiles match exact case-insensitive native identifiers, never fuzzy names. Native machine IDs, machine models and MLC models are retained without renaming. Explicit device-ID profiles can override a generic model profile. Multiple matches are unavailable. MLC alone cannot select a machine-speed model.

The example catalog maps the observed technical ESAPI labels together with the MLC metadata: `TDS + Varian High Definition 120` selects the **TrueBeam-HD120** reference profile; `RDS + SX2` selects the **Halcyon-SX2** reference profile. TDS/RDS are not presented as separate device families. Raw ESAPI metadata is preserved for diagnostics; the user-facing profile names identify the reference system. HD120 alone is not a unique treatment-unit identifier, so the exact combination is required and a locally configured machine-ID profile takes precedence. Exact `TrueBeam` and `Halcyon` model-name examples are also included. Profile speeds remain **declared assumptions**: 6 degrees/s for the PlanCheck-style C-arm example and 12 degrees/s for the Halcyon RapidArc example. Halcyon's 4-rpm mechanical maximum is not substituted for treatment timing; the vendor describes RapidArc delivery typically at 2 rpm. Installed machine/software settings need local verification. [Vendor description](https://cancercare.siemens-healthineers.com/products/radiotherapy/treatment-delivery/halcyon).

## Presentation and source separation

- `EstimatedDoseRateMuPerMin` and `EstimatedSegmentDurationSeconds` are separate from `PlannedDoseRateMuPerMin`.
- GUI: dashed **Estimated (segment)** series; visible **Estimated plan profile · Not measured delivery**, profile/speed and detailed tooltip. The horizontal axis is control-point index, not measured time.
- PDF/HTML: same shared rate/aperture renderer; dashed estimate versus solid supplied values; profile ID, assumed speed, estimated duration and a local-verification caveat. HTML includes escaped detailed assumptions/availability with a bounded scrollable legacy-browser fallback. No nominal trajectory.
- Native ESAPI input consists only of copied scalar machine metadata, gantry direction/angles, nominal rate and meterset weights. The [review-controls source discussion](review-controls/README.md#dose-rate-presentation) links the native API definitions and explains why actual delivery timing is not recovered here.

Analytic tests cover both speed profiles, MU conservation, CW/CCW wrap, zero/duplicate intervals, repeated invalidation, nonunit final weights, precision limits and strict profile matching. Engineering tests and the built-in consistency checks are not clinical commissioning. Public examples remain patient-free; clinical publication material requires separate institutional review.
