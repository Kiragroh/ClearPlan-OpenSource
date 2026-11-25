# ClearPlan v3.3.0: paper companion

This archive holds the material that the manuscript "Configurable plan review and reporting through treatment planning system scripting" points to. The application source is in `ClearPlan-OpenSource-v3.3.0-source.zip` (tag `v3.3.0`).

## S4_verification

Phantom verification of 29 October 2026, read-only via ESAPI.

- `goal-configurations/revision1.json`: baseline test goal tables (901 single-layer, 902 double-layer).
- `goal-configurations/revision2.json`: impossible volume requirement.
- `goal-configurations/revision3.json`: tightened complementary-volume goal.
- The thresholds are software-test inputs, not clinical recommendations.
- `entry-point/`: the verification entry point.
  - `HeadlessPlanReport.cs` and its helpers.
  - `PhantomValidation.cs` selects the explicit test goal table.
  - `build_validation.ps1` is the compile step against a built candidate.
  - `verify_native.py` is the field-by-field comparison of goal entries, plan parameters and check findings.
- Plan and patient selection is local and not included.

## S4_recalculation

Independent recalculation written without ClearPlan code.

- `esapi/ClearPlanK1ReferenceExport.cs`: read-only export of dose grid, contours and machine settings.
- `python/k1_reference.py`: recalculation of dose, aperture and modulation values.
- `python/k1_mlc_check.py`: leaf-geometry plausibility check.
- `python/test_k1_reference.py`: tests for the two scripts above.
- `tolerances.json`: the tolerances fixed before the run.
- `mlc_geometry.json`: vendor leaf geometry used.

## S6_routine_logs

`aggregate_plancheck.py` and the aggregated tables of Supplement S6. They contain no patient, course or plan identifiers and no free text.

## Not included

- Clinical data, phantom identifiers and local configuration.
- Credentials and proprietary TPS assemblies.
