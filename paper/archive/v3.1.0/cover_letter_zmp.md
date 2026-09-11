Dear Editors,

We are pleased to submit the Short Communication (technical note) entitled
“ClearPlan: a simulator-backed contract for reproducible verification of
radiotherapy plan-review software” for consideration in
*Zeitschrift für Medizinische Physik*.

ClearPlan addresses a practical reproducibility problem in radiotherapy
software: local review rules and treatment-planning-system interfaces vary,
whereas software verification and training material should be shareable
without patient data. The framework separates read-only TPS acquisition,
external RefDB JSON or Excel rule sources, a validated vendor-neutral review
snapshot, and a shared interface and report layer. It combines plan quality
metrics, PlanCheck findings, structure mappings, field-identifier and
field-name conformance, source status, and dose-volume histograms.

The public release includes a standalone vendor-free simulator with seven
versioned deterministic scenarios. The manuscript reports only synthetic
software verification: 125 C# tests and 20 tests of an illustrative,
unvalidated RayStation adapter. It does not claim clinical sensitivity,
specificity, workflow improvement, error reduction, or safety benefit.
RayStation is presented only as an example of the adapter boundary, not as a
commissioned clinical implementation.

We believe the manuscript is relevant to the journal because it provides a
clinically grounded, open, and reproducible method for inspecting and
regression-testing multi-domain plan-review software while retaining local
commissioning and human oversight. The manuscript and figures contain
exclusively analytically generated synthetic data; no patient data,
patient-derived values, human participants, or patient-identifiable information were
used. Ethics approval and informed consent were therefore not applicable.

The reproducible software version corresponding to the manuscript is ClearPlan
v3.1.0:
https://github.com/Kiragroh/ClearPlan-OpenSource/releases/tag/v3.1.0

The authors declare no competing interests and received no funding for this
work. The corresponding author will confirm final approval from all coauthors
before submission. The manuscript is not currently under consideration
elsewhere.

Sincerely,

Maximilian Grohmann
University of Leipzig Medical Center, Department of Radiation Oncology
Stephanstraße 9a, Building 5.2, 04103 Leipzig, Germany
maximilian.grohmann@medizin.uni-leipzig.de · +49 341 97 18226
ORCID: https://orcid.org/0000-0002-6909-811X
