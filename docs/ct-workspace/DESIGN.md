---
name: ClearPlan Clinical Blueprint - CT workspace
description: Scoped record of the built read-only native three-plane CT surface; inherits the existing Clinical Blueprint design authority.
---

# Design System: ClearPlan CT workspace

## Overview

**Creative North Star: "Clinical Blueprint"**

This is a scoped implementation record for the native **Schnittbilder** surface, observed on 2026-09-09. The established [root DESIGN.md](../../DESIGN.md) remains the visual authority. This addition does not replace its identity, redefine its tokens, or establish a new global page template. Its surface mode is **Operate**: inspect three orthogonal overviews, enlarge one plane, and understand the displayed evidence and its limits.

The built surface extends the incumbent paper-like workspace with dark image fields. It keeps the native controls and quiet slate/teal hierarchy around physically scaled images. Availability is stated in words and counts; image, contour, and dose availability remain separate concepts.

**Key Characteristics:**

- Three simultaneous planes by default: Transversal, Koronal, Sagittal.
- Explicit enlargement with native scrolling and persistent selected-plane context.
- Display-only switches for detached contours, isodoses, and viewport framing.
- Source-specific limitations, fixed CT windowing, and conspicuous synthetic labeling.

Primary evidence is [PlanImagesView.xaml](../../ClearPlan.Presentation/Views/PlanImagesView.xaml), [PlanImagesViewModel.cs](../../ClearPlan.Presentation/ViewModels/PlanImagesViewModel.cs), and [PlanImageRenderer.cs](../../ClearPlan.Rendering/PlanImageRenderer.cs). Geometry comes from [PlanImageViewport.cs](../../ClearPlan.Core/Review/PlanImageViewport.cs); native source bounds come from [EsapiPlanImageBuilder.cs](../../ClearPlan.Script/Review/EsapiPlanImageBuilder.cs). The delivery scope is recorded in the [CT workspace and target-consistency plan](../superpowers/plans/2026-09-09-ct-target-consistency.md). This document describes the built CT surface, not every intention in that plan.

No additional token primitives are declared here. Inherited color, type, spacing, shape, and button tokens stay in the root authority and [ClinicalBlueprint.xaml](../../ClearPlan.Presentation/Styles/ClinicalBlueprint.xaml). There is no scoped sidecar: this record adds neither a new token system nor HTML/CSS component previews.

## Colors

### Primary

The inherited ClinicalBlueprintTeal marks loading progress and the shared button interaction states. ClinicalBlueprintCobaltDark carries the section heading. Display controls use the incumbent pale-teal secondary-button treatment, not a new action palette.

### Secondary

Contour colors come from each detached overlay. Native structure colors are copied from the source structure; native isodose colors identify explicitly labeled levels relative to prescribed total dose. They are image evidence, not plan-quality or approval colors.

Orientation letters and the isocenter cross use the renderer's existing bright aqua. The synthetic renderer places a contrasting band labeled **SIMULATION - NOT PATIENT DATA** inside each image. These are scoped rendering details, not new Clinical Blueprint brand tokens.

### Neutral

The surrounding surface uses inherited text, muted text, and border resources. The image field is the renderer's dark slate, matched by the XAML image container. Missing-plane messages remain readable on that dark field. Outside acquired CT coverage stays unfilled; it is not a replacement background anatomy.

## Typography

The surface inherits Segoe UI. The section title uses the existing section-title style (21 DIPs, semibold); plane names use a smaller semibold heading (16 DIPs). Availability and loading status use compact text (12 DIPs); field-of-view context, orientation explanations, and enlarge labels use supporting text (11 DIPs). These are observed local roles, not additions to the root type scale.

Plane names, HU, millimeters, percentage of Rx, and Gy labels identify what is being shown. The image itself carries orientation letters; the adjacent explanatory line expands R/L, A/P, S/I and identifies the cross as the isocenter. Truncated legend labels retain their source description in a tooltip.

## Layout

The surface uses the incumbent horizontal/vertical inset (24/22 DIPs). The order is title and reload action, wrapping display controls, availability and status, image viewport, then orientation key and a collapsible legend. The first three areas remain outside the image scroll region.

The default view is a single row of three equal-width plane cells. Each contains a name and **Vergrößern** action, a stretching dark image container, and its own availability and field-of-view readout. The image area has a minimum content height (180 DIPs); the layout does not silently switch into a mobile stack.

Enlarging selects one cell and one column. Its minimum content height becomes **520 DIPs**, so it is genuinely larger at compact window sizes. A native vertical ScrollViewer exposes the remaining image content; horizontal scrolling is disabled. The surface title changes to **Schnittbilder · [selected plane]**, and **Drei Ebenen** remains above the scroll region. The selected plane is therefore still named when its in-cell heading has scrolled away. This is viewport scrolling, not CT slice navigation.

Images use uniform stretching and nearest-neighbor bitmap scaling. The renderer maps both axes in physical units into a square viewport, preserving non-square source pixel spacing. Extra space is a dark margin, not stretched anatomy. The legend expands below the image area and has its own bounded vertical scroll region (maximum 112 DIPs).

The four inspected screenshots use mathematical synthetic images only:

| State | Larger window | Compact window |
| --- | --- | --- |
| Three-plane overview | [1600 × 1000](../../.impeccable/review/ct-three-planes-1600x1000.png) | [1180 × 720](../../.impeccable/review/ct-three-planes-1180x720.png) |
| Enlarged coronal plane | [1600 × 1000](../../.impeccable/review/ct-enlarged-coronal-1600x1000.png) | [1180 × 720](../../.impeccable/review/ct-enlarged-coronal-1180x720.png) |

Sizes name the tested windows, not a promise about exported PNG pixel dimensions. These images document the layout and explicit simulation states; they do not contain actual patient CT data and do not establish clinical commissioning, mobile support, or accessibility conformance.

## Elevation & Depth

The CT surface is flat: the inherited light workspace surrounds thinly bordered dark image fields. It does not apply the theme's optional card shadow. Spatial image content supplies depth; the controls do not introduce floating panels or decorative elevation.

## Shapes

Display actions retain the shared modestly rounded button shape. CT image fields are square-cornered, clipped to their bounds, and separated by thin inherited borders. The view does not convert the three planes into raised cards. The legend uses compact rectangular color swatches beside text.

## Components

### Display controls

**Strukturen** and **Isodosen** are native CheckBox controls, both checked by default and disabled when no valid image is available. They rebuild only the detached display and legend; they neither recapture anatomy nor edit a clinical structure or dose. Buttons retain inherited hover, keyboard-focus, and disabled states. Control labels, image names, loading progress, and selected detail text have explicit automation names or help text in XAML; this implementation evidence is not a complete assistive-technology test.

**CT laden** requests a new capture for the current plan and is disabled while loading. Loading clears the old detached image set and shows explicit per-plane placeholders plus indeterminate progress. Failure exposes a retry message instead of retaining substitute images. Image replacement accepts only exact, nonempty matching plan keys; duplicate captures of a plane yield an ambiguous/unavailable plane, not arbitrary selection.

### Windowing and framing

**W 400 / L 40 HU** is a focusable informational readout, not a window/level editor. Its tooltip explains that the snapshot already contains windowed 8-bit pixels. Native capture uses vendor HU conversion and this fixed window; simulation uses an HU-equivalent mathematical phantom.

**Iso-Fokus** is the initial framing mode. With a valid in-bounds isocenter and captured dose region, the physical viewport is centered on that isocenter and fits the sampled **at least 2% Rx** region plus a margin (10 mm). Source-limited dose coverage is named explicitly. If the dose region is unavailable, the entire CT extent remains visible; if the isocenter is also unavailable, framing falls back to image center. The local readout identifies missing isocenter or absent dose zoom.

**Gesamte CT** fits the whole CT around image center. It does not remove the captured isocenter or dose region from the snapshot. Overlay visibility and framing are independent: hiding isodose lines does not discard the captured dose-region framing data. Neither mode extrapolates beyond the captured CT.

### Availability, legend, and sources

The summary states **n / 3 Ebenen verfügbar**. Each valid plane reports **CT verfügbar** or **Simulation**, followed by displayed contour and isodose counts and any unavailable-source count. A count reflects available overlays with drawable paths, not the total structure inventory or a clinical completeness verdict. A valid source with no intersection in the selected plane contributes no displayed contour and is distinct from an unavailable source.

The collapsible legend contains currently displayed overlays, deduplicated across planes by kind, label, and color. Source descriptions remain attached to legend items. The focusable field-of-view readout exposes its caption, capture summary, framing explanation, and unavailable-source reasons through a tooltip and automation help text. The ordinary status line stays short rather than repeating this technical detail.

### Captured image and overlay boundaries

The view displays one captured orthogonal overview per plane, near the treatment isocenter. It is not an interactive volume viewer and has no slice slider, contour editing, window/level adjustment, dose calculation, clinical save, or approval action.

Native capture supports the verified single external-beam planning context with HU image data and a single unambiguous treatment isocenter. Unsupported geometry, absent data, invalid dose, and failed source reads remain unavailable; the renderer does not manufacture replacements. Native objects remain on the owning STA during capture; the display receives detached pixels, coordinates, labels, and overlay paths.

The native structure display is bounded to **up to 24 nonempty segmented structures**, ranked with targets first. The source inventory and capture work also have bounds. This is a display subset, not a target assignment or proof that every structure is represented. Transverse contours come from native image-plane contours; coronal and sagittal contours are native-mesh plane intersections and retain that approximation in provenance.

Native dose overlays require valid total-plan dose and a positive absolute prescription. The capture samples at most 192 × 192 positions per plane and traces levels at 2, 20, 50, 80, 95, 100, and 107% Rx; labels also carry Gy. Only drawable available levels appear in the plane count and legend. Native dose-grid coverage is respected, and the 2%-Rx focus is based on sampled extent, not a claim of exact dose-volume containment. Isodoses are drawn before structure contours; both are clipped to CT coverage.

## Do's and Don'ts

These guardrails describe this CT surface and its existing safety boundaries; they do not add global design policy.

### Do:

- **Do** preserve the inherited Clinical Blueprint resources and native controls.
- **Do** retain simultaneous three-plane inspection, explicit enlargement, and the selected-plane heading outside the scroll region.
- **Do** keep physical aspect ratio, orientations, fixed windowing, displayed counts, and source limitations understandable together.
- **Do** distinguish missing CT, missing overlay sources, and a source with no intersection in the current plane.
- **Do** retain simulation warnings and the read-only meaning of overlay and framing controls.

### Don't:

- **Don't** equate available images, displayed contours, or isodose levels with plan approval or clinical completeness.
- **Don't** present the bounded 24-structure display as the entire native structure set.
- **Don't** fabricate anatomy, contours, dose, or coverage outside the detached source.
- **Don't** describe image-region scrolling as slice navigation, or fixed W400/L40 as an adjustable window.
- **Don't** treat synthetic screenshot coverage as native clinical acceptance or commissioning.
