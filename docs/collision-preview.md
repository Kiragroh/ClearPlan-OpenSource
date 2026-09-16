# Collision / 3D: read-only development preview

[Documentation](index.md) · [Configuration guide](configuration-guide.md)

This feature visualizes detached body/support meshes and sampled treatment-field
positions. It is **not a clinical clearance decision, a machine interlock, or a
validated motion simulator**. No plan, structure, dose, or treatment parameter is
written back to the TPS.

## Use

1. Open **Collision / 3D**. In the simulator, select the explicit C-arm sphere or
   ring-bore demonstration. These dimensions do not describe a clinical device.
2. For ESAPI, configure an exact-machine, locally commissioned geometry catalog
   under **Settings → CollisionProfilesJsonPath**. The path is persisted in
   `settings.ini`; the existing configuration-history editor supports revisions,
   author/comment records, validation, and restoration. The distributed example
   has an intentionally empty `Profiles` array.
3. ESAPI does not expose table pitch/roll in the referenced API. The per-plan
   checkbox requires the operator to verify fixed zero pitch/roll across all
   fields. It is initially unchecked; changing it invalidates the scene and
   report image. This is operator assurance, not an API measurement. It does not
   capture treatment-day 6D corrections.
4. Load geometry and select a field. Its actual planned arc is sampled every 10°
   plus exact endpoints. **Play / Pause**, **Previous / Next**, the slider,
   and table rows use the same cached states. **All beams** advances to the
   next field (on by default); **Loop** loops the selected sequence (off
   by default). Field transitions are discrete, not simulated couch motion.
   Playback stops on manual selection or changing plans, tabs, or geometry.
   Static fields do not gain an invented rotation. No ESAPI queries or
   clearance recalculation run on ticks.
5. **10° trajectory** is enabled by default: for a moving head, status markers trace
   the sampled positions. **All envelopes** optionally adds overlapping contours
   from the other positions and is initially off. The active source-head no-fly
   model has a translucent shell with opaque status-colored contours; the
   contours are not a replacement for the evaluated volume. Green means the
   displayed model-sample pass, captured surface-clear or radial-clear state;
   yellow means incomplete/uncertain, red a model overlap. Read the displayed
   scope before interpreting the color.
   The stationary Halcyon bore is drawn once with opaque edges and a faint shell,
   without a fictitious axial end wall or stacked rotating cylinders. Its
   displayed source-model length follows the captured surface extent, not a
   measured housing length. All angle states remain in the table.
6. The table gives separate body/table distances in mm. `lower … upper` is a
   conservative lower bound and sampled upper estimate. A negative lower bound
   alone does not prove overlap. `≤ upper` means the lower bound is unknown.
   Finite RingBore radial values do not establish end/rim safety margins and
   therefore cannot produce global pass. Unknown, open, or CT-edge-truncated
   External coverage prevents global pass, including meshes artificially capped
   by the TPS. The separate Halcyon source-model radial review can show
   **Clear (captured)** (`clear`) when the captured body and support surfaces lie
   beyond the configured radial warning margin. This is an open-bore radial
   check at nominal setup, not axial clearance, complete anatomy coverage or
   treatment clearance. It does not turn the global uncertain result into pass;
   missing/invalid radial geometry remains uncertain.
7. Drag with the left mouse button to rotate; use the wheel to zoom or
   **Reset view** to restore the camera. Framing uses the whole field's
   sampled geometry, so changing the active angle does not refit the camera.
   The HFS room view applies one display-only rotation by the negative couch
   angle about the vertical axis through isocenter to body, support, source and
   model together. It does not modify native coordinates used for distances or
   establish noncoplanar clinical validity. Camera, **Body**, **Couch / support**,
   **Machine envelope** and trajectory/contour visibility change only the illustration,
   never the collision calculation or clinical evaluation.
8. Optionally select **3D figure in HTML / PDF**. Export uses an overview at the smallest displayed model distance across beams
   and a table of per-beam minima, restoring the GUI selection afterward. No new pages are appended to existing reports. Retained
   report PNGs cannot reconstruct the meshes or new CT-coverage metadata: those
   require a fresh read-only native capture.

## Required catalog data (schema version 1)

Each profile requires `MachineId` (exact treatment-unit ID), `Kind`, `Revision`,
`CommissionedBy`, `Evidence`, `Commissioned`, `SyntheticOnly`, and
`SafetyMarginMm`. No alias or MLC inference selects physical device geometry.
Duplicate IDs and JSON properties, including case variants, are rejected.

| Kind | Locally measured, verified dimensions in mm | Model interpretation |
|---|---|---|
| `CArmSphere` | `HeadCenterFromIsoMm`, `HeadRadiusMm` | A conservative sphere enclosing the relevant head assembly, centered along the native source ray. A nominal source distance alone is not a commissioned head envelope. |
| `RingBore` | `BoreRadiusMm`, `BoreHalfLengthMm` | A finite cylindrical clearance bore centered at isocenter along the HFS longitudinal axis. No insertion path or external gantry geometry. |

Real-machine commissioning requires measured envelopes, coordinate verification,
coverage/accessory review, boundary cases, and appropriate site-approved margins.
Checking a JSON flag does not perform those measurements. No real-machine example
dimensions or approval are supplied by this repository.

## Evaluation and unavailable states

### Captured surface result and assurance (2026-09-15)

For a supported nominal TrueBeam source model, `model-clear` describes only a
positive conservative distance bound beyond the configured warning margin for
all supplied surface primitives. `partial-clear` means one role is clear while
the other remains unresolved or absent. `model-hit` retains a nonpositive sampled
witness, including nominal setups. The original `CollisionSweepFrame.Status`
and all clinical topology, CT-coverage and coordinate gates remain unchanged.
GUI and report label this **surface** result separately from missing commissioning,
setup, CT or support evidence. Surface separation is not volume containment or
treatment clearance. Table rotation alone does not invalidate a computable
nominal surface distance.

Head-surface classification requires at least one proper triangle and valid
conservative bounds over every supplied primitive. Degenerate triangles are not
deleted: both exhaustive and spatial samplers include them. The bore retains its
separate radial geometry validation. A point-only mesh cannot become a surface
clear result.

The MLC inset uses a stored captured CP index and captured gantry/couch angles,
not the renumbered 10-degree row index. Interpolated angles explicitly label the
nearest bracketing captured CP as a reference. The aperture is not interpolated
and no other CP's DRR is underlaid. Native plan/beam fingerprints and finite
isocenter coordinates must match; missing or changed geometry clears the inset.
All capture stamps stay in memory. Playback never accesses ESAPI.

**All beams** continues across fields, and **Loop** optionally loops.
Transitions between fields are discrete jumps, not inferred couch trajectories.
Manual navigation, tab/context changes and failures stop playback. The MLC panel
stays fixed above the independently scrolling distance table.

- Native nominal capture requires one machine, HFS, one nonempty EXTERNAL
  structure, and known fixed table translations across fields. Couch yaw must
  remain fixed within each field but may differ between fields. Native source
  positions are retained in the patient frame; a patient-mesh transform is not
  guessed for evaluation. The stricter supported-coordinate/profile gate still
  requires zero couch yaw for every pose and operator-confirmed zero pitch/roll.
  Noncoplanar nominal display is not clinical validation; unsupported or ambiguous
  coordinates remain unavailable. Missing SUPPORT geometry cannot establish pass.
- Sphere evaluation uses full triangle distances, including interiors, and
  containment checks. Bore evaluation clips triangles to its finite axial slab
  before measuring radial clearance. Empty axial coverage is unavailable.
- Negative modeled clearance means overlap; positive clearance inside the
  configured margin means near. `sampled-clear` means only that the captured
  modeled positions did not intersect. It does not mean that delivery is safe.
- Missing geometry/profile/assurance, invalid meshes, coordinate uncertainty,
  canceled work, or exceeded workload limits cannot produce a pass. A single
  unavailable sub-result keeps the aggregate unavailable.
- Control-point samples do not cover continuous arc, between-field, setup, or
  insertion motion. CT coverage, accessories and mesh self-intersections remain
  limitations requiring local verification.
- Native objects are accessed on the owning dispatcher. Only copied numbers and
  triangles reach worker threads. Patient meshes and cached illustrations are
  excluded from review JSON; an explicitly requested report can contain the
  visible anatomy and must be handled as clinical data.

Current capture/sweep bounds are 10,000 input control points, 2,000 derived
angle states, and 250,000 vertices / 1,500,000 indices per captured surface.
The sweep also bounds surface–frame pairs at 20,000. Non-spatial evaluation
checks triangle and vertex counts times the number of distance evaluations
against 10 million each; the commissioned-profile evaluator retains its
triangle–pose budget and topology/coordinate gates. An analytically stationary
Halcyon source bore reuses distances and radial reviews at identical isocenters
without dropping displayed angle states or their global classifications.

The TrueBeam source-head distance path uses a spatial bounding-volume hierarchy
for larger meshes and a shared budget of 10 million charged distance evaluations
across the sweep. Conservative 1-Lipschitz bounds allow pruning only when a node
cannot improve an upper bound sampled on a supplied triangle. Repeated index
triples can share one distance-domain entry; the original mesh and its topology
classification remain unchanged. Smaller meshes use the bounded direct sampler.
Budget exhaustion remains uncertain, not pass. These are development limits,
not an accuracy guarantee; no mesh decimation or topology repair is performed
to obtain an answer.

## BEV display convention

Numeric ruler labels at ±5/±10 cm are removed; physical X1/X2/Y1/Y2 jaw values,
ticks and ISO remain. The effective MLC aperture is green. With an available
calibrated BLD-to-display transformation, the DRR is shown in an upright C0
gantry frame and the MLC/jaws rotate with the collimator. A missing calibration
stays explicitly BLD-aligned; the renderer does not guess a rotation sign from
the angle alone. Jawless dual-layer plans do not acquire fictitious jaws.

## Development verification

Build from the solution, then run `ClearPlan.Core.Tests.exe Collision` and
`ClearPlan.Core.Tests.exe BevCollimator`. `tools/collision/verify-collision-preview.ps1`
compiles a synthetic-only offscreen smoke probe using the built simulator DLLs.
It exercises desktop/compact layouts, ring switching, unavailable-state clearing,
and optional PDF/HTML output. It does not start ESAPI or access a patient.

Native and simulator compilation, analytic unit tests, and synthetic images do
not substitute for a live ESAPI acceptance test or local device commissioning.
This development change is not automatically deployed to the clinical runner.

## Methodological background

The need for explicit device geometry and independent validation is consistent
with published CT-based collision-prediction work:
[Frontiers in Oncology, 2021](https://www.frontiersin.org/journals/oncology/articles/10.3389/fonc.2021.617007/full)
and [the published collision-prediction study available in PMC](https://pmc.ncbi.nlm.nih.gov/articles/PMC4608969/).
These references do not validate this implementation or supply local dimensions.
