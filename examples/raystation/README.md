# Illustrative RayStation adapter

> **UNVALIDATED EXAMPLE — NOT FOR CLINICAL USE**
>
> This code demonstrates a read-only adapter boundary for RayStation v2025 SP2
> (scripting API 17.2.0). It has not been run or commissioned in a RayStation
> clinical environment. The generated JSON omits source identifiers but still
> contains **confidential clinical dosimetry**. Do not commit, publish, attach
> to an issue, or move the output outside an institution-approved local
> environment.

`clearplan_review_snapshot.py` reads the currently opened RayStation context and
writes a ClearPlan review-snapshot JSON file. It is an interoperability example,
not a claim that Eclipse/ESAPI rules and RayStation clinical goals are
semantically equivalent.

## Deliberate safety boundary

- Reads only the current `Patient` presence, `Case`, `Examination`, `Plan`, and
  `BeamSet`; it does not query or load other patients.
- Uses `TreatmentCourse.TotalDose` only. RayStation dose values are converted
  from cGy to Gy, relative volume fractions to percent, and ROI volume is kept
  in cm³.
- Samples cumulative DVHs at 1 Gy by default and rejects non-finite,
  out-of-range, or nonmonotone values instead of silently repairing them.
- Replaces ROI labels with deterministic role pseudonyms such as `target-01`
  and `oar-01`. Patient, case, examination, plan, beam-set, beam, ROI, and DICOM
  identifiers are not serialized.
- Transcribes numeric `DoseAtVolume`, `VolumeAtDose`, and `AverageDose` goal
  values, but marks every PQM row `not-evaluated`/`info`.
- Emits one explicit `not-evaluated` PlanCheck row because institution-specific
  RayStation checks have not been implemented or commissioned.
- Derives PlanFieldNamer-style suggestions from gantry, couch, and arc-direction
  geometry while omitting actual field IDs and names. The couch token precedes
  `UZ`/`GUZ`, including duplicate suffixes.
- Contains no RayStation save, create, delete, update, set, add, remove, import,
  or DICOM-export calls.
- Accepts only an absolute non-UNC output path. UNC destinations and implicit
  overwrite are rejected; mapped or redirected storage still requires local
  information-security review.

These safeguards reduce disclosure and modification risks; they do not make the
example clinically validated.

## Illustrative use inside RayStation

Run this as a reviewed file script or wrap it in an institution-approved
database script. RayStation command-line `-run` executes an approved database
script name; it is not a stand-alone API process.

```python
from clearplan_review_snapshot import export_current_review_snapshot

export_current_review_snapshot(
    r"C:\ClearPlan-local\raystation-review.json",
)
```

The local parent directory must already exist. Existing files are preserved.
If replacement is intentional, pass `overwrite=True`.

The adapter resolves the context accessor lazily in this order:

1. `raystation`
2. `raystation.v2025`
3. legacy `connect`

The fallback is present only to make the example legible across API generations;
it is not evidence of compatibility with an installed RayStation version.

## Tests outside RayStation

The unit tests use identifier-rich fake objects and verify that none of those
source values enters the JSON. They also check units, DVH monotonicity, supported
goal transcription, field geometry suggestions, path restrictions, lazy import
order, deterministic output, and an AST-based ban on write-like RayStation API
calls.

```powershell
python -m unittest discover -s examples\raystation\tests -v
```

Passing these synthetic tests does not validate RayStation runtime behavior.
The public build also exports one deterministic fake-context JSON and checks
that the C# `ReviewSnapshot` validator accepts it. The released simulator does
not import arbitrary adapter JSON, so this is a schema-contract check rather
than an end-to-end UI integration test.
Before any local evaluation, confirm the installed API version, run against
approved test patients only, inspect every output field, perform independent
commissioning, and keep the generated JSON outside version control.
