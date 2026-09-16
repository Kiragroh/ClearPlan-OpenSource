# Architecture and extension boundaries

[Documentation](index.md) · [Configuration](configuration-guide.md) · [Developer handoff](developer-handoff.md)

ClearPlan is a review application, not a replacement dose engine or a plan-approval
system. The clinic supplies its accepted workflow, constraints, mappings, and checks.

## Data flow

```text
Native TPS context + local configuration
                 |
    read-only, vendor-specific adapter
                 |
     detached review/analysis snapshot
        /              |             \
 shared WPF UI   detached calculations   PDF / offline HTML
```

Eclipse reads stay on the owning ESAPI STA/dispatcher. Workers receive copied
numeric data, not live patient, plan, structure, or image objects. After asynchronous
work, cancellation and active-plan guards run before publishing results. Clinical
rendering geometry remains in memory; it is not ordinary public JSON interchange data.

## Code ownership

| Project | Responsibility |
|---|---|
| `ClearPlan.Core` | Normalized constraints, mappings, settings, contracts, detached calculations, configuration history |
| `ClearPlan.Script` | Read-only ESAPI acquisition, native UI integration, settings, public starter checks |
| `ClearPlan.Presentation` | Shared WPF workspace, review state, selection, comparison, plots |
| `ClearPlan.Rendering` | CT/BEV rendering from detached data |
| `ClearPlan.Reporting` / `.MigraDoc` | Captured report document, PDF, offline HTML |
| `ClearPlan.Simulator` | Patient-free scenarios using the same review boundary |
| `ClearPlan.Core.Tests` | Portable execution and source-boundary contracts |
| `examples/raystation` | Illustrative snapshot producer and fake-context tests |
| `ClearPlan.Dicom` | Optional RTPLAN enrichment; not required by native ESAPI review |

The desktop runner manages the ESAPI session through EsapiEssentials; it does not
provide a replacement TPS runtime. Licensed vendor assemblies are supplied locally.

## Policy and acquisition are separate

RefDB JSON and Excel normalize into one catalog. Consumers should not branch on
source-specific fields. Explicit aliases preserve laterality; ambiguity is not
resolved by choosing the first structure. Table compatibility precedes ranking.
Confirmation of a public stock table is not clinical policy approval.

Versioned documents cover target rules, field naming, aliases, check inclusion,
CT approvals, display defaults, and machine profiles. Import copies a source;
restore appends a revision. Actor/time/hash history is local provenance, not a
tamper-proof authenticated audit trail.

Public `ErrorCalculator.cs` contains general examples, not the complete private
clinical implementation. Inclusion settings filter findings but do not necessarily
stop legacy calculation. Exact CT approvals have a [narrow adapter scope](CT_COMPATIBILITY.md).

## Metrics and output

PAM measures the fraction of a target's BEV projection blocked by the effective
aperture. Physical leaf-strip openings define each layer. Single-layer geometry
uses that opening; dual-layer geometry uses the **intersection of both openings**.
Real jaws or explicit fixed jawless limits clip the result. Missing layers are not
replaced with invented jaws or a single-layer approximation. Projection, sampling,
and beam weighting are specified in the [analysis method](plan-analysis-method.md).

Measurements retain units, provenance, and unavailable states. Estimated dose rate
is not measured delivery. A sampled collision scene is not continuous motion
clearance. GUI and active-plan reports share detached review data and visibility;
a comparison reference must not silently become the exported active plan.

Optional report BEV capture has a builder-owned cooperative budget. Timeout can
mark optional images unavailable; requested cancellation or stale context still
rejects the result. A blocked vendor call cannot be forcibly interrupted by that budget.

## Another TPS and the research aim

A new adapter must explicitly map units, fractionation, target prescription,
coordinates, DVH queries, warnings, and missing-data semantics. Similar API property
names do not establish equivalence. The RayStation example withholds clinical
pass/fail classifications and is not a complete PAM/CT/BEV/collision port.

The Technical Note aims to describe configurable workflow integration that can be
adapted to scripting-capable TPSs. Portable tests, native API verification, local
commissioning, and clinical performance studies remain separate evidence scopes.
