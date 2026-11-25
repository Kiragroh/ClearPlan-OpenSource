---
name: ClearPlan Clinical Blueprint
description: Observed native WPF review workspace; slate navigation, teal actions, white evidence surfaces.
colors:
  ClinicalBlueprintCanvas: "#F3F5F7"
  ClinicalBlueprintSurface: "#FFFFFF"
  ClinicalBlueprintSurfaceMuted: "#EEF1F4"
  ClinicalBlueprintTeal: "#0F766E"
  ClinicalBlueprintCobaltDark: "#16324A"
  ClinicalBlueprintNavigation: "#17212B"
  ClinicalBlueprintNavigationSelected: "#313A46"
  ClinicalBlueprintAmber: "#D99B36"
  ClinicalBlueprintGreen: "#15803D"
  ClinicalBlueprintRed: "#B42318"
  ClinicalBlueprintText: "#202B38"
  ClinicalBlueprintTextMuted: "#667585"
  ClinicalBlueprintBorder: "#D6DDE4"
  ClinicalBlueprintPassSurface: "#EAF6EE"
  ClinicalBlueprintVariationSurface: "#FFF5DD"
  ClinicalBlueprintFailSurface: "#FDEDEC"
  ClinicalBlueprintInfoSurface: "#EAF2F7"
typography:
  section-title:
    fontFamily: "Segoe UI"
    fontSize: "21px"
    fontWeight: 600
  subsection-title:
    fontFamily: "Segoe UI"
    fontSize: "17px"
    fontWeight: 600
  summary-value:
    fontFamily: "Segoe UI"
    fontSize: "23px"
    fontWeight: 600
  table-body:
    fontFamily: "Segoe UI"
    fontSize: "12px"
    fontWeight: 400
  navigation-label:
    fontFamily: "Segoe UI"
    fontSize: "14px"
    fontWeight: 600
rounded:
  navigation: "3px"
  button: "4px"
  card: "6px"
  badge: "10px"
spacing:
  compact: "8px"
  card: "16px"
  section: "24px"
components:
  button-primary:
    backgroundColor: "{colors.ClinicalBlueprintTeal}"
    textColor: "{colors.ClinicalBlueprintSurface}"
    rounded: "{rounded.button}"
    padding: "8px 14px"
  button-secondary:
    backgroundColor: "#E5ECEC"
    textColor: "#175D59"
    rounded: "{rounded.button}"
    padding: "8px 14px"
  card:
    backgroundColor: "{colors.ClinicalBlueprintSurface}"
    rounded: "{rounded.card}"
    padding: "{spacing.card}"
  navigation:
    backgroundColor: "{colors.ClinicalBlueprintNavigation}"
    typography: "{typography.navigation-label}"
    rounded: "{rounded.navigation}"
    width: "184px"
    height: "46px"
  status-badge:
    backgroundColor: "{colors.ClinicalBlueprintInfoSurface}"
    rounded: "{rounded.badge}"
    padding: "2px 8px"
---

# Design System: ClearPlan

## Overview

**Creative North Star: "Clinical Blueprint"**

September 14 three-axis follow-up: the BEV information rail retains a separate
gantry front view, couch top view and collimator beam-axis schematic. Each has its
own angle and fixed reference; asymmetric BLD axes make collimator rotation
readable even without usable CT pixels. With validated upright display calibration,
the collimator inset follows the same screen rotation as the aperture and is labelled
BLD/screen (fixed screen-up reference). Otherwise it is explicitly BLD/nominal with
a zero-angle reference. The numeric caption retains the native machine angle.
This schematic is not a collision clearance result; gantry/couch remain separately
labelled nominal room/top views. No image calibration is inferred from a number.

September 14 native finish review: report title is always "Plan Quality Report";
software/version and identity remain in the footer. Decorative red warnings are
removed, while named clinical result states and missing-input disclosures remain.
Physical MLC layers use green (single/distal) and blue (proximal), reinforced by
solid/dashed line styles. The fixed nominal Halcyon field envelope is dashed white,
distinct from movable jaws. BEV fits the whole panel by default; a native checkbox
enables enlarged scrollable details. The collision view shows actual captured
External/Support surfaces and explicitly uncommissioned source geometry, without
converting a nominal illustration into a clear/passed collision finding.

Native Citrix/ESAPI observations covered Halcyon 3D and BEV activation at the
available desktop window size; native TrueBeam PDF capture verified five field-start
DRRs. Local PDF review covered private/masked 14-page TrueBeam and 12-page Halcyon
reports. This is a scoped desktop review, not mobile or accessibility certification.

September 2026 review refinement: the native header measures two automatic-height rows; plan switching is visible next to the sandbox action. PTV quality precedes delivery traces, with explicit plan-Rx/body scope and no pass/fail decoration. DVH legends remain outside the right edge of the plot area.

Interaction follow-up: DVH hover reads structure/dose/volume without clicking; reset only restores detached selection and axes. Parameter charts have aligned headers, a direct jump action and an explicitly nominal MU/min readout. Native geometry/PAM populates after initial rendering, with current selections preserved; loaded BEV panels reactivate on snapshot replacement. Both CI forms are labeled Paddick CI and 1 / Paddick CI.

Feedback and report follow-up: routine completion uses restrained teal instead of warning color; unavailable results remain explicitly named. Both DVH trackers own opaque white/dark-text contrast. Missing CT projection detail moves into a keyboard-accessible information tooltip. HTML-Quicklook extends the existing report identity with offline tables, outside-legended DVHs and cached imagery; it is an in-app snapshot view, with an explicit HTML save action.

This documents the built native Windows review workspace, not a new brand concept. Slate navigation frames paper-white evidence surfaces; restrained teal actions and traces support inspection without competing with numeric results.

The implementation is the source of truth: `ClearPlan.Presentation/Styles/ClinicalBlueprint.xaml`, `Views/ReviewWorkspaceView.xaml`, `Views/PlanParametersView.xaml` and `Views/PlanComparisonView.xaml`. Extend that incumbent world; do not introduce a separate web-dashboard identity.

**Key Characteristics:**

- Dense, readable numeric tables alongside wide plots.
- Native Windows selection, scrolling and command behavior.
- Explicit units, provenance, availability and synthetic-mode labels.
- Quiet boundaries; color supports a named finding, not an inferred overall verdict.

Token lengths use portable CSS px notation for the observed WPF device-independent units. They are not physical-pixel guarantees. The sidecar's HTML/CSS snippets and synthesized tonal strips are explanatory panel previews only, not runtime implementations or additional approved design tokens.

## Colors

Teal carries actions and evidence traces; cool slate neutrals establish the workspace hierarchy. Resource names are retained, including the historical `CobaltDark` name; `ClinicalBlueprintCobalt` is an existing alias of the same teal value.

### Primary

- **ClinicalBlueprintTeal:** primary buttons, selected-cell focus and principal parameter traces.
- **ClinicalBlueprintCobaltDark:** section headings and contextual emphasis.

### Secondary

- **ClinicalBlueprintAmber:** navigation selection indicator and named variation/fallback states.
- **ClinicalBlueprintGreen / ClinicalBlueprintRed:** explicitly labeled pass/available and fail/unavailable states, respectively.
- **Pass / Variation / Fail / Info surfaces:** pale backgrounds behind the corresponding textual status.

### Neutral

- **Navigation / NavigationSelected:** persistent slate rail and selected destination.
- **Canvas / Surface / SurfaceMuted:** page background, evidence areas and subdued supporting regions.
- **Text / TextMuted / Border:** primary reading, secondary explanation and separators.

**The Named Status Rule.** Bind status color to the stated finding or availability category. A positive row status is not plan approval or an overall clinical pass.

## Typography

Segoe UI is the native workspace family. Table values stay normal weight; titles and navigation use semibold. The observed hierarchy is recorded in the frontmatter, not a newly invented type scale.

Summary values provide a quick entry point; subsection headings separate charts, beam tables and definitions. Keep units beside numbers or in column/axis labels. A missing value is an em dash with context where needed, never an implied zero. Preserve German characters, superscripts and delta symbols.

## Layout

The incumbent workspace has a fixed left navigation rail (196 units), a top context row (82 units) and a stretching content area. Its declared minimum is 960 × 600 units; this is a minimum workspace size, not a responsive breakpoint.

The two analysis views use a content inset of 24 horizontally and 22 vertically. Content scrolls vertically; tables retain their own scrolling when necessary. Do not shrink all lower content into the first viewport.

The parameter surface reads totals, beam selection, paired wide plots, beam table, then metric definitions and target provenance. Its four-column summary and equal-width plot pair are specific to that surface; they are not global page templates.

The comparison surface uses paired plan identity labels, one wide DVH and separate constraint/parameter tables. Reference is dashed; current plan is solid. Keep identity, units, goal differences and interpretation adjacent to comparison values.

Built screenshot coverage includes 1600 × 1000 and compact 1180 × 720 windows. See `.impeccable/review/finish-review.md` for the exact limits; this evidence does not establish mobile support or accessibility conformance.

## Elevation & Depth

Depth is predominantly flat: light canvas, white plotting areas, tonal headers and thin borders. The theme defines a card shadow (WPF blur radius 12, depth 2, opacity 0.09), but the default card style does not apply it and the two new views do not use it. Do not infer a general shadow system from that unused resource.

## Shapes

Buttons have modest rounded corners, cards slightly softer corners, and status badges compact pill-like corners; frontmatter records the observed radii. Navigation uses a narrow left selection/focus stroke. Plot containers and summary strips are flatter, frequently square and separated by hairlines. Do not convert every evidence region into a raised card.

## Components

### Buttons

Primary actions use teal and white; secondary actions use a pale teal surface and dark teal text. Both share the same native Button template. Hover darkens the background to the existing darker teal; keyboard focus darkens it further. Disabled opacity is 0.45. Preserve command availability and disabled explanatory tooltips.

### Inputs / Fields

Beam, target and current-plan selectors are native ComboBox controls with explicit accessible names. The DVH filter is a native TextBox with a descriptive tooltip. Their platform behavior is retained; the project does not define a bespoke input theme. Do not substitute web-like fake controls.

### Navigation

A left TabControl presents stable review destinations. Selected tabs use the lighter slate surface, warm pale text and an amber left stroke. Unselected hover lightens the slate; keyboard focus also exposes the amber stroke. Selection must remain understandable without relying on color alone.

### Cards / Containers

The shared card is white, thinly bordered and padded. Analysis summaries use top/bottom rules; plots use white containers with restrained padding. Choose the implemented container appropriate to the content, not a uniform card grid.

### Status / Mode Badges

Status badges combine category-specific pale fill, border and text. The same colors serve distinct result and availability vocabularies; keep their labels explicit. The separate mode badge makes synthetic demonstrations conspicuous and carries the not-for-clinical-use warning.

### Data Tables

Native read-only DataGrid rows have a minimum height of 29, horizontal rules, pale alternating rows and semibold headers. Selected rows and cells use distinct muted fills; focused cells gain a teal border. Numeric data stays compact and consistently formatted. Preserve visible units, missing-value markers, source limitations and explanatory columns.

### Evidence Plots

Dose-rate follow-up (2026-09-11): nominal MU/min is a scalar readout only, never a
substitute trajectory. Separately calculated estimates are dashed and explicitly
labeled as estimated, not measured delivery. The assumed machine profile/speed
is visible; detailed assumptions move to a tooltip. Supplied planned samples
remain solid with missing-value gaps; when neither source is available a quiet
wrapped message replaces the plot. New workspaces and direct reports hide
unmatched goals by default; the complete mapping editor and explicit opt-out stay
available. Hidden counts remain disclosed, not counted as passed.

OxyPlot views occupy wide white surfaces. Titles, legends, axes and units identify the measurement; dashed/solid line semantics supplement color. Planned and nominal dose-rate data must not be presented as measured delivery data. Complexity metrics and signed deltas are descriptive, not standalone quality verdicts.

## Do's and Don'ts

### Isodose and collision display extension (2026-09-15)

The CT image keeps its physical framing. A separate dark scale strip below each
plane labels only present contours in % Rx and Gy; it never overlays anatomy.
The GUI has one deduplicated native-text legend beneath all three planes, so
fitting cannot shrink labels into unreadable raster lettering. Per-page export
scales remain embedded for portability. Explicit add/remove actions and the
per-level visibility checkbox are committed together with Apply.
Preserve 2/20/50/80% context levels. The default high-dose sequence is green 95%,
mint 100%, yellow 105%, pink 107%, red 110%, purple 125%. The native collapsible
editor uses explicit Apply and keeps session changes separate from versioned
defaults. A reload preserves the current display; Loaded defaults is deliberate.
For report collision geometry, the minimum-pose overview is unobscured; a separate
strip underneath holds the miniature room reference and two independent angle
circles. Dashed zero references and numeric angle labels supplement position.
The GUI reserves a separate right-hand orientation column, not an overlay on the
patient. The heading, angle labels and geometry advance together on each rendered
frame; playback controls may already request the next frame.

### Do:

- **Do** extend the existing Clinical Blueprint resources and native Windows behavior.
- **Do** keep numeric tables and wide evidence plots readable at the compact tested window size.
- **Do** show units, provenance, availability and synthetic-mode context beside the evidence.
- **Do** distinguish reference/current and planned/nominal/measured concepts in words and line styles.
- **Do** retain explicit finding-level results in reports and review screens.

### Don't:

- **Don't** infer treatment approval, clinical superiority or an aggregate pass from a raw status, complexity metric or delta.
- **Don't** hide absent data behind zero, decorative success color or unexplained blanks.
- **Don't** replace native controls or the incumbent slate/teal identity during a scoped extension.
- **Don't** treat preview HTML/CSS, synthesized tonal strips or screenshot review as runtime or clinical validation.
