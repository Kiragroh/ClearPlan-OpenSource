# Read-only RTPLAN enrichment

`ClearPlan.Dicom` reads an explicitly selected DICOM file into detached numerical
geometry. It does not open a patient, write to a treatment-planning system,
calculate dose, deliver radiation, or alter clinical constraints. A successful
file parse alone does not authorize use with the active plan.

## Integration contract

```csharp
ParsedRtPlan RtPlanReader.Read(string path, int? fractionGroupNumber = null);
ReviewPlanAnalysis RtPlanReader.MatchAndApply(
    ReviewPlanAnalysis active, ParsedRtPlan parsed,
    string expectedSopUid, string expectedFrameUid = null);
```

Both methods are static in namespace `ClearPlan.Dicom`. The host obtains the
exact active plan's SOP Instance UID on its ESAPI owner thread, presents a native
file dialog, and runs parsing/calculation away from that owner thread. It must
reject an active-plan/snapshot change across any await before publishing the
result. Do not search a folder and pick its first RTPLAN or match by plan label.
If available, pass the active frame UID as a second identity check.

`Read` requires one fraction group or an explicitly selected group number. Its
`ParsedRtPlan.Analysis` uses neutral beam ordinals and contains no imported plan,
patient, or machine labels. SOP class, SOP instance, study, and frame UIDs remain
in memory under Newtonsoft `[JsonIgnore]`; do not log the parsed object through
another serializer, the input path, or raw DICOM datasets. `RtPlanImportException`
contains a fixed safe `Code` and message, without a raw inner exception.

`MatchAndApply` first validates the entire plan and returns a new detached
snapshot. Failure leaves the caller's snapshot unchanged. It requires:

- Exact nonempty SOP Instance UID match and, when supplied, frame UID match.
- Equal uniquely numbered treatment-beam sets, matched by `BeamNumber`.
- Beam MU within `max(0.01 MU, source MU * 1e-6)` and identical CP counts.
- Identical CP indices; circular gantry, collimator, and couch-angle differences
  at most 0.01 degrees; normalized cumulative MU-weight difference at most 1e-6.
- Every isocenter component in patient coordinates within 0.1 mm.

These are serialization/rounding tolerances, not clinical acceptance thresholds.
The deep copy preserves ESAPI-derived dose-per-fraction, selected target,
outlines, projected target strips, and projection-quality metadata. The import
replaces physical aperture geometry and planned dose rate, then recalculates
descriptive aperture/PAM/MU metrics. Missing target coverage remains unavailable.
Neither dose nor target contours are reconstructed from RTPLAN.

## Supported geometry and provenance

The SOP whitelist contains standard RT Plan Storage
`1.2.840.10008.5.1.4.1.1.481.5` and documented Varian Private RT Plan Storage
`1.2.246.352.70.1.70`. The latter is explicitly listed in the primary
[Vision RT conformance statement](https://visionrt.com/wp-content/uploads/2026/04/1016-0711-DICOM-Conformance-Statement-1.0.pdf).
The modality string alone is never sufficient; file and dataset SOP identifiers
must agree. Only PATIENT-coordinate static/dynamic photon treatment beams with
MU dosimetry are accepted; setup beams are excluded.

The parser reads referenced beam metersets and normalizes each cumulative weight
by `FinalCumulativeMetersetWeight`. Each layer requires its own exact `N+1`
ordered leaf boundaries and `2N` bank positions. Later CPs inherit omitted
positions and beam parameters; CP0 must supply every physical device. Planned
`DoseRateSet` applies to the segment beginning at that CP and is inherited until
changed. Absent rate stays unavailable. These semantics follow the
[DICOM RT Beams Module](https://dicom.nema.org/medical/dicom/current/output/chtml/part03/sect_C.8.8.14.html).

Supported device definitions are MLCX, MLCY, and the complete MLCX1/MLCX2 dual-layer
extension, plus X/Y or ASYMX/ASYMY limits. No leaf width or layout is guessed from
leaf count. A partial dual-layer definition is rejected. Both physical layers
are intersected in the isocenter plane; different leaf boundaries are retained.

For the dual MLCX1/MLCX2 extension, X/Y tags are conservatively represented as a
separate `FixedBoundingBox`, not physical movable jaws. They clip the aperture
and must stay constant within 0.001 mm across CPs. `HasJaws` stays false.
Asymmetric jaw tags alongside this dual-layer extension are rejected because
that hardware interpretation is not established by this adapter. These are
explicit adapter safeguards, not a claim that a SOP class identifies all machine
hardware. Standard single-layer X/Y or ASYMX/ASYMY jaws remain physical geometry.

## Rejection boundaries and limitations

Malformed/missing arrays, reversed banks or boundaries, inconsistent CP/fraction
counts, duplicate/unknown beam references, nonmonotonic or inconsistent final
weights, unknown devices/classes, enhanced device definitions and block-defined
apertures are rejected. Files above 128 MiB and sequences above 20,000 items are
rejected. Leaf-pair count is bounded at 2,000 per physical layer.

Moving isocenters and changed tabletop translations are rejected. Explicit
nonzero tabletop pitch/roll/eccentric and gantry-pitch rotations are rejected;
they are outside the target-projection geometry available to the host. Target
projection has additional ESAPI-side orientation/quality gates. No RTSTRUCT,
measured dose rate, treatment log, inter-CP delivery reconstruction, or clinical
pass/fail threshold is supplied by this adapter.

## Verification

Run `tools/build-and-test-dicom.ps1`; it restores pinned dependencies in locked
mode, builds .NET Framework 4.8, and executes synthetic `.dcm` byte files written
with fo-dicom. Fixtures contain no patient attributes and exist only in a unique
temporary directory removed after the run. They test both layers, inherited
positions/rates, final-weight normalization, real jaws/fixed limits, malformed
geometry, missing/ambiguous groups, unsupported classes/orientations, exact-plan
match failure, input immutability, target retention, and identifier suppression.

For a separately authorized local technical inspection, run:

```text
ClearPlan.Dicom.Tests.exe --inspect <exact RTPLAN path> [--fraction-group <number>]
```

This mode is read-only and writes no fixtures, logs, or artifacts. Output is
restricted to class category, technical beam/CP/layer counts, jaw/limit flags,
and planned-rate/aperture availability. It excludes UIDs, filenames, labels,
isocenters, raw positions, angles, and MU values. It does not validate an active
ESAPI plan match. Host runtime and independently reviewed clinical-case checks
remain separate release gates; synthetic tests do not replace them.

See [dependency and packaging notes](DICOM_DEPENDENCIES.md).
