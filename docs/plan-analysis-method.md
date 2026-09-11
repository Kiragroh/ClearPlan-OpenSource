# Plan parameters and aperture analysis

This is read-only research software. Unit tests, a successful build, and the native-outline consistency guard are not clinical commissioning. No machine settings, patient structures, plans, or dose are modified.

## Quantities and units

- Total MU is the sum of treatment-beam metersets in MU; setup and imaging-treatment fields are excluded. Missing/invalid beam MU makes the plan total unavailable, rather than producing a partial total.
- MU/Gy divides that per-fraction total by the prescribed dose per fraction in Gy. ESAPI cGy is converted explicitly; relative dose is not accepted as Gy.
- Plan normalization is the descriptive native TPS percentage, shown in plan comparison and the active-plan report. Missing normalization is unavailable, not zero; the value is not an independent dose check.
- Beam PAM is the cumulative-meterset-weighted fraction of the selected target's BEV area blocked by the aperture. Native plan PAM combines beam PAM using `Beam.WeightFactor / sum(Beam.WeightFactor)`, matching the inspected PlanCheck aggregation convention. Synthetic/imported snapshots explicitly retain the separate `MetersetMu` mode. PAM is target-specific and dimensionless, ranging from zero to one. It is not MCS, aperture variability relative to jaws, or an empirical MU surrogate. The implemented blocked-area definition follows Hernandez et al., equations 1-2. See [the original paper](https://aapm.onlinelibrary.wiley.com/doi/10.1002/mp.70144).
- Mean/minimum/maximum open area are in cm2. The small-aperture fraction is the MU-weighted fraction of sampled apertures below the explicitly displayed area threshold (default 4 cm2). This descriptive threshold is not a clinical pass/fail limit.

Every control point is retained. For an interval, its MU is beam MU multiplied by the change in cumulative meterset weight divided by the final weight. Half that interval MU is assigned to each endpoint for the aperture/PAM averages. This trapezoidal endpoint convention is explicit numerical sampling, not reconstruction of continuous leaf motion. Zero-MU transitions do not receive an artificial equal weight. The control-point MU field is the interval ending at that point. See [Varian's control-point definition](https://docs.developer.varian.com/api/17.0/VMS.TPS.Common.Model.API.ControlPoint.html).

Native PAM does not silently substitute MU for a missing or invalid beam weight. Missing/nonfinite/negative factors, zero total weight, or unavailable PAM for a positive-weight beam leave plan PAM unavailable. Other MU and aperture summaries retain their stated MU weighting. Field classification and weight factors are included in the native snapshot freshness checks and the BEV field list uses the same treatment-field scope.

## Aperture geometry

The calculation uses physical leaf boundaries and both banks of every layer, in isocenter-plane beam-limiting-device millimeters. Each layer is a union of leaf-strip openings. The effective aperture is the intersection of these layer unions and any real jaw rectangle. X- and Y-travel MLC layers are supported.

A fixed jawless field boundary is represented separately as `FixedBoundingBox`; it clips the opening without becoming a physical jaw. Native profile modes distinguish `Physical`, `FixedLimits`, and `None`; fixed limits must remain invariant across control points. A missing layer, unknown leaf widths, or malformed boundaries never become an assumed single-layer model or synthetic jaws. The optional legacy DICOM adapter remains separate from the current native ESAPI workflow.

The bundled ESAPI assembly is version 1.0.450.29. Its public MLC/control-point API does not expose the complete layer/physical-boundary metadata required here. The native adapter reads `ControlPoint.LeafPositions` and maps its indices using an explicit JSON profile with exact `Beam.MLC.Model`, physical leaf boundaries and a complete non-overlapping assignment to layers. Profile files are strict UTF-8, limited to 256 KiB and read off the owning dispatcher with a ten-second await limit. No leaf-count or model-name fallback guesses a layout. The current isolated preview includes an exact SX2 mapping: 28 distal and 29 proximal leaf pairs, 10 mm physical widths and a 5 mm stagger at isocenter. PAM uses the intersection of both physical layer openings, without invented jaws. Controlled read-only native exports exercise this mapping; that engineering verification is not clinical commissioning or a clinical release.

## Target projection and runtime boundary

An explicit user target selection takes precedence; an invalid explicit choice remains unavailable and is never replaced silently. In automatic mode, the detached projection workflow evaluates every eligible nonempty segmented PTV recognized by the shared DVH target policy and selects the lowest complete, finite PAM. Missing or partial candidate results are excluded, not treated as zero. The report retains each candidate's availability, the valid/eligible counts and selection provenance; equal PAM values use deterministic target-ID ordering. This numerical minimum does not establish clinical superiority and does not change the TPS target assignment. If no eligible PTV exists, a valid native `PlanSetup.TargetVolumeID` can supply the fallback target. See the [review controls](review-controls/README.md#target-rules-and-pam) for explicit overrides and unavailable states.

A native `GetStructureOutlines(target, false)` outline supplies the beam-start target in collimator coordinates; `true` would instead give the collimator-zero BEV frame. A static beam can reuse the native outline only when its orientation and table position remain unchanged. See [Varian's beam API](https://docs.developer.varian.com/api/17.0/VMS.TPS.Common.Model.API.Beam.html).

Full moving-gantry projection is an explicit asynchronous action, not part of routine review refresh. ESAPI mesh vertices/triangle indices, native outlines, and source coordinates are copied synchronously on the owning dispatcher. Only detached data enter the worker task. The active snapshot is deep-copied before enrichment.

The worker divergently projects the triangulated target surface using vendor source positions. It rasterizes the **union** of projected triangle spans at 0.625 mm; overlapping front/back faces do not cancel, and gaps between disconnected components are not filled by a convex hull. This resolution matches the C# implementation described in the source paper; it remains a numerical approximation.

The inspected local PlanCheck implementation instead switches between 1.25 mm and 0.625 mm at a target volume of 20 cc. Matching its blocked-area, endpoint-sampling and beam-aggregation conventions is not a claim of bitwise equality: ClearPlan retains its bounded mesh projection, fixed 0.625 mm sampling and native-outline consistency guard. The legacy binary was inspected, not run against the native cases as an independent numerical benchmark.

The collimator-coordinate rotation is measured from corresponding native collimator-zero/beam-coordinate outline points, rather than guessing its sign. At beam start the projected silhouette must differ from the native silhouette by at most 3% of their union area (symmetric difference / union, equivalent to one minus IoU). This is a numerical consistency guard, not a clinical accuracy or acceptance claim. Failure leaves moving PAM unavailable.

The current moving-projection envelope is HFS, zero fixed patient-support angle, fixed collimator, and fixed tabletop translation. Other orientations/motions are explicitly unavailable. The RTPLAN merge must reject nonzero pitch/roll/eccentric rotations not exposed by the bundled ESAPI. Limits are 250000 surface triangles, 4096 raster cells per dimension, four million triangle-row spans per control point, and a cancellable 30-second numeric-work budget. No control points are silently subsampled. A beam commits target projections only after all its control points complete.

## Dose-rate provenance and privacy

`Beam.DoseRate` is shown as nominal selected MU/min, never as a substitute trajectory. The native path now additionally calculates a separately labeled PlanCheck-style **estimated plan trajectory** from normalized control-point MU, directed angular travel and exact configured machine profiles. Estimated interval time is the maximum of angle/configured speed and 60 * MU/nominal rate; it is not actual delivery time. Missing/invalid inputs remain unavailable. Acceleration, MLC/jaw dynamics, ramping and holds are omitted. Supplied planned samples and estimates occupy separate DTO fields and solid/dashed plot series; synthetic rate examples remain labeled synthetic. See [method, settings and limits](DOSE_RATE_ESTIMATE.md).

Meshes are temporary memory-only inputs. Aperture geometry, target outlines/strips, and isocenter coordinates are excluded from review JSON and remain detached in-memory rendering data. Tests use synthetic geometry exclusively. No patient identifiers or clinical geometry belong in source fixtures or logs.

## CT-derived beam's-eye views

The clinical DRR reads the active planning CT via ESAPI `GetVoxels` and
`VoxelToDisplayValue` on the owning STA dispatcher. Grid axes, voxel-center
origin and spacing are copied with uniformly integer-strided samples (at most
256 per axis); a trailing border smaller than a stride can be omitted. Worker
tasks receive only detached numeric arrays. HU values are integrated along
divergent source-to-isocenter rays with clipped voxel support and trilinear
sampling. Positive-integral percentile windowing is for display only. This
attenuation proxy has no spectrum, scatter or beam-hardening model and is not
calibrated electron density, dose, diagnostic imaging or registration-grade DRR.

The native coordinate frame is calibrated from vendor source locations and
paired native structure outlines. The current envelope is HFS, zero fixed
couch, fixed collimator and fixed table translation. The 45-second budget is
cooperative and cannot interrupt an individual blocked vendor call. Report
beam starts share one detached CT capture. An absent or unsupported source is
explicitly unavailable; clinical mode never generates substitute anatomy.
The seven original simulator scenarios use an original synthetic head/torso BEV
phantom and separate orthogonal torso slices. Their reports explicitly disclose
that those slices are not the BEV source volume: they illustrate layout, not
cross-image anatomical correspondence.

The additional `publication-single-layer` and `publication-dual-layer` scenarios
use one deterministic mathematical ellipsoid phantom instead. The same 3D
structure and analytic dose definitions drive the three attached CT planes,
structure/isodose overlays, voxel-sampled target/OAR DVHs, numerical clinical-goal
evaluation and target CI/HI/GI. The DRRs use that phantom's synthetic attenuation
volume, and PAM uses the same ellipsoidal target with the respective single-layer
or jawless dual-layer apertures. Sampling and rasterization remain numerical
approximations. Crucially, the analytic dose is **not calculated from the MLC
apertures**: these fixtures demonstrate coherent displays and calculation paths,
not physical dose-model agreement, delivered dose or clinical validation. The
public `SyntheticPublicationScenarioFactory` supplies both GUI and report data;
the original seven fixtures remain unchanged.

Memory-only native-state fingerprints compare raw leaf/jaw arrays, MU, rates,
angles, table translations, accessories, treatment-unit identity, normalization,
prescription and CT/structure-set identity/grid metadata. A missing or changed
stamp refuses the report instead of retaining stale aperture values. Cancellation
is checked before any resumed native read after an asynchronous boundary. These
stamps do not hash HU voxels or contour content: unchanged metadata does not
prove unchanged anatomy. They are excluded from public review JSON.

One shared renderer draws DRR, layer boundaries, effective opening, physical
jaws, field/CP values and orientation scheme in the UI and PDF. Virtual field
bounds are not named X1/X2/Y1/Y2. The small gantry front view and couch top view
are orientation aids, not patient-position or collision checks. Report images
use CP0, even when closed, and are not integrated arc fluence.

## Verification

Build the `ClearPlan_Core_Tests` solution target in Release/x64, then run `artifacts/bin/Release/ClearPlan.Core.Tests.exe PlanAnalysis`. The targeted suite covers physical staggered layer intersection, jaw/fixed-boundary clipping, holes/concavity, target-area PAM, MU normalization, incomplete inputs, projection perspective/rotation, triangle union, cancellation, and private-geometry copying. Clinical commissioning still requires controlled Eclipse comparison for representative TrueBeam/Halcyon plans, including independent DICOM geometry checks and the chosen target.
