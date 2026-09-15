# Development source: 3.2.0-dev.20260915.7

This source update follows public commit `d21dbd50fb1b02ce7180cbf44a2531eda48ce8d3`
and the archived v3.2.0 release. It is a development snapshot, not a new approved
clinical package. Reproductions should record the exact source commit as well as
the local configuration revisions and native TPS/API versions.

## Included

- Shared, editable isodose palette: retain lower levels and control each level's
  percent, color and visibility. Add/remove levels in the GUI; apply to the three
  image planes and active-plan report. The GUI has one common legend; PDF image
  pages retain their own scale. Dose samples are detached read-only data, not a
  dose recalculation or TPS change.
- More compact report field pages and MLC/DRR presentation, with physical field
  limits and separate dual-layer colors.
- Optional collision/geometry review with cached angle playback, per-field
  body/support distances, CP-linked MLC inset, and separate gantry/couch
  orientation panels. Color labels describe the captured model scope only;
  missing geometry, incomplete CT, unsupported coordinates and uncommissioned
  models do not become treatment clearance.
- Updated configuration-path editing, clinical-report document metadata and
  receiving-clinic guidance.

## Public/private boundary

The public source retains the institution-neutral starter `ErrorCalculator.cs`.
It does not include a clinic's complete private check implementation, native
headless automation, clinical captures, reports, screenshots, local settings,
credentials, proprietary TPS assemblies or measured/source-derived device
catalogs. `CollisionProfiles.example.json` and
`SourceCollisionModels.example.json` contain empty catalogs. A clinic must supply
and commission its own profiles; no default machine is silently inferred.

The public catalog regression asserts this empty-default boundary. A separate
local headless-report regression remains part of private deployment testing;
the portable test checks the public snapshot adapter rather than requiring that
private tool. Remaining geometry tests use explicitly artificial analytic
fixtures, not clinical patient data.

## Verification and limitations

Run `tools/test-portable-review.ps1` from a Windows workstation with the .NET
Framework 4.8 developer tools. This builds only the vendor-free core tests and
simulator. It does not compile or execute licensed ESAPI integration, import
clinical data, upload an ARIA document or establish clinical performance.

The earlier release's paper files and evidence remain historical versioned
artifacts; they must not be described as validation of this development update.
Current clinical manuscript drafts and figures are deliberately not distributed
with this source update. The RayStation example remains an illustrative,
unvalidated snapshot adapter, not a commissioned equivalent of the ESAPI host.
