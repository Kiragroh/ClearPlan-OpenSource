# Documentation

[Project home](../README.md) · [Configuration](configuration-guide.md) · [Architecture](architecture.md) · [Developer handoff](developer-handoff.md)

ClearPlan separates software behavior, local clinical policy, and clinical acceptance.
Start with your task, then consult the detailed method when you need its exact boundary.

## Get started

- [Configuration guide](configuration-guide.md): paths, sources, aliases, defaults, and controlled changes.
- [Clinic onboarding](CLINIC_ONBOARDING.md): owners, acceptance, rollout, privacy, and rollback.
- [Developer handoff](developer-handoff.md): builds, tests, extension points, and open limits.
- [Architecture](architecture.md): native acquisition, detached data, calculations, and reporting.
- [Display language](LANGUAGE.md): German/English behavior and identity preservation.
- [Actual Settings gallery](settings-gallery.md): patient-free screenshots of editable configuration.

## Methods and review behavior

| Topic | Detailed reference |
|---|---|
| Structure visibility, comparison, reports, and check inclusion | [Review controls](review-controls/README.md) |
| PAM, physical apertures, target projection, MU, and BEV/DRR | [Plan-analysis method](plan-analysis-method.md) |
| Estimated dose rate and machine assumptions | [Dose-rate estimate](DOSE_RATE_ESTIMATE.md) |
| Sampled collision geometry and model limits | [Collision preview](collision-preview.md) |
| CT, dose, and contour rendering | [Historical CT workspace design](ct-workspace/DESIGN.md); current controls in the [configuration guide](configuration-guide.md#display-and-report-scope) |
| Native BEV rendering | [Current method](plan-analysis-method.md#ct-derived-beams-eye-views) and [historical design notes](BEV_DESIGN_NOTES.md) |

PAM cites [Hernandez et al.](https://aapm.onlinelibrary.wiley.com/doi/10.1002/mp.70144).
The method document specifies target, sampling, projection, and aggregation conventions;
different implementations are not automatically numerically interchangeable.
[RT Complexity Lens](https://rt-complexity-lens.lovable.app/metrics) provides additional conceptual explanations.

## Configure local policy

- [Structure aliases](STRUCTURE_ALIASES.md) and [field-name defaults](DEFAULT_FIELD_NAMING.md).
- [Exact CT/HU approvals](CT_COMPATIBILITY.md).
- [Configuration history](CONFIGURATION_HISTORY_STORE.md).
- [Stock workbook maintenance](../tools/stock-catalog/README.md).
- [Optional ARIA report upload](ARIA_REPORT_UPLOAD.md).

## Connect and validate

- [Illustrative RayStation adapter](../examples/raystation/README.md): limited, unvalidated snapshot example.
- [Optional RTPLAN import](DICOM_RTPLAN_IMPORT.md) and [dependencies](DICOM_DEPENDENCIES.md): separate from native ESAPI review, which needs no DICOM export.
- [Local commissioning scaffold](local-commissioning-scaffold.md).
- [Reproducibility and evidence](reproducibility.md).

## Version and research records

Documented code baseline:
[`98aed7bc2fb3ecbced7cc2a520f691da36c53bfa`](https://github.com/Kiragroh/ClearPlan-OpenSource/tree/98aed7bc2fb3ecbced7cc2a520f691da36c53bfa)
(`3.2.0-dev.20260916.1`). Documentation-only commits can follow this baseline;
record their own commit separately when reproducing an entire checkout.

- [Current development notes](releases/2026-09-16-development-source.md).
- Historical [v3.2.0](releases/v3.2.0.md) and [v3.1.0](releases/v3.1.0.md) records.
- [Archived synthetic paper materials](../paper/README.md), not the current local author draft.

Public examples must be patient-free. Clinical figures and reports remain on institution-approved storage.
