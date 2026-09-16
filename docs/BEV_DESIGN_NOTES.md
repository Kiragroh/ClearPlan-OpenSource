# Historical BEV surface design notes

[Documentation](index.md) · [Configuration guide](configuration-guide.md)

This is an archived layout record. Colors, rulers, and captions below describe an
earlier iteration, not the current renderer. Use the [analysis method](plan-analysis-method.md)
and [configuration guide](configuration-guide.md) for current behavior.

## Scope and authority

Mode: **OPERATE**. This is a scoped native WPF extension for selecting a field and control point (CP), then inspecting its projection, aperture and parameters. It does not introduce a new visual identity. [DESIGN.md](../DESIGN.md) records the public design system; this note adds no global tokens or product strategy.

The direction contract is at the top of [BeamEyeView.xaml](../ClearPlan.Presentation/Views/BeamEyeView.xaml). The implementation retains the Clinical Blueprint navigation, Segoe UI typography, teal actions and quiet white information surfaces. The dark viewport serves image inspection within that system.

## Layout and selection

- Field selection sits above the large BEV panel. Previous/next commands and a discrete CP slider support deliberate movement through the selected field; the DRR calculation action is an explicit action for the selected CP.
- The image container stretches with the available area and has a **640 DIP minimum height**. In compact windows its inner vertical scroll area preserves that height instead of shrinking the entire panel into a thumbnail. Horizontal scrolling is disabled. This is a container constraint, not a guarantee of a 640-pixel anatomical image.
- CP controls, the DRR action and status text occupy separate rows outside that scroll area, keeping them adjacent and pinned while the panel content scrolls.
- The shared renderer composes a square image viewport and an adjacent white field-information panel. It shows field/CP identity, machine and MLC model, energy, technique, meterset, planned and nominal rates, aperture area, and current/field-start angles. Missing values remain explicitly unavailable; a later CP must not be presented as CP0.

## Image and aperture semantics

[BeamEyeViewRenderer.cs](../ClearPlan.Rendering/BeamEyeViewRenderer.cs) maps projection and aperture into one measured coordinate frame: beam-limiting-device geometry in millimetres at the isocentre plane, with centimetre rulers and named X/Y axes. Here, the coordinate scale does not imply measured delivery data or validated clinical geometry.

- MLC layers use distinct cyan and amber outlines with matching named legends in the reviewed one-/two-layer examples. Light leaf-body tint leaves anatomy visible; drawn bodies extend to the viewport edge and do not establish physical leaf-body dimensions.
- Physical jaws use solid yellow boundaries and separate X1/X2/Y1/Y2 values. Y-jaw labels reserve a clear gap from the ruler's tick-label band.
- Fixed virtual field limits, when present, use dashed white boundaries and an explicit virtual-boundary label. A jawless field is identified as jawless; virtual bounds are never relabelled as physical jaws.
- DRR availability and aperture-geometry availability are separate states. Missing or mismatched image data is withheld with an explanation; it is not silently replaced with a valid-looking projection.
- Front-room gantry and couch-top schematics remain separate, explicitly schematic orientation aids. They are not collision checks or patient-coordinate verification.

## PDF continuity

[ReportPdf.cs](../ClearPlan.Reporting.MigraDoc/ReportPdf.cs) uses the same renderer for the per-field **field-start** panel, preserving the image, boundary/layer semantics and adjacent parameters. The report describes the exact first control point, including a closed start aperture when present, rather than arc-integrated fluence. Its field-start purpose is distinct from the viewer's currently selected CP.

## Evidence and limits

The historical local visual review and captures are not distributed with this public source. Their former paths are not public evidence or current validation.

| Capture | Recorded coverage |
|---|---|
| Dual-layer BEV (historical local capture) | 1600 × 1000, synthetic jawless/two-layer example |
| Single-layer BEV (historical local capture) | 1600 × 1000, synthetic physical-jaw example |
| Compact BEV (historical local capture) | 1180 × 720 window / 1164 × 681 client capture, inner scroll area and visible CP controls |

These examples deliberately use synthetic phantoms and generic 20/21-pair fixtures, not commissioned machine geometry. The PDF overview uses a separate illustrative torso fixture, **not the source volume for the BEV DRRs**; the image sets do not demonstrate anatomical correspondence. Preserve the conspicuous simulation/not-for-clinical-use warning.

This note records source and supplied visual evidence, not a new runtime test. Real ESAPI operation remains unvalidated. The finish review does not establish CP/selector interaction, scrolling behavior, loading/error-state behavior, keyboard or screen-reader access, performance, other window sizes, dosimetric or geometric accuracy, commissioning, clinical use, or accessibility conformance.
