# ClearPlan v3.1.0 verification traceability

This matrix links the technical-note requirements and principal software
hazards to deterministic scenarios, oracles, executable checks, and retained
evidence. It applies to synthetic software verification only.

| ID | Requirement or hazard | Exercised case | Oracle | Test or guard | Retained evidence |
| --- | --- | --- | --- | --- | --- |
| T01 | Public demonstrations must contain no clinical record, private path, or DICOM UID. | All seven scenarios and the exact staged release files. | Every scenario is explicitly synthetic, uses generic labels, and is byte-identical to its reviewed source file. | `ReviewSnapshot.Synthetic`, `ReviewSnapshot.PublishSafeLabels`, `ReviewSnapshot.EmbeddedPrivateContent`, `SyntheticScenario.Provenance`, `test-public-simulator-package-input.ps1`. | Scenario JSON, package-input pass, release SHA-256. |
| T02 | PQM classifications must follow the checked-in DVH, stated interpolation, one-decimal rounding, and goal/variation boundaries. | Baseline, target-underdose, OAR-overdose, and mixed-review PQMs. | Independent interpolation, integration, rounding, and comparator implementation in the test harness. | `SyntheticScenario.PqmDvhConsistency`; `SyntheticDvhMetrics.Interpolation`, `.MeanDose`, `.Thresholds`, `.Invalid`. | `technical_note_evidence.json`, Table 3, Figure 3. |
| T03 | Deliberately injected PlanCheck failures must be distinguishable from calculated PQMs. | `metadata-plancheck` and `mixed-review`. | Fixed expected row identities and status counts; injected non-pass messages carry an explicit synthetic prefix. | `SyntheticScenario.RowsAndStatuses`, `.VisibleFaults`, `.MixedReview`. | Scenario JSON, Table 3, Table 4. |
| T04 | Field identifier and field-name conformance must remain separate, and couch angle must precede direction. | `field-and-mapping` and `mixed-review`. | Expected identifier prefix and exact five suggested names. | `FieldNameSuggester.*`, `SyntheticScenario.FieldAndMapping`, `validate-readonly-field-preview.ps1`. | Figure 1b, evidence JSON suggested-name list. |
| T05 | Ambiguous aliases must remain visible and require human resolution. | One ambiguous mapping in `field-and-mapping` and `mixed-review`. | Two available structures yield one `variation`, not an arbitrary pass. | `StructureAliasResolver.Ambiguity`, `SyntheticScenario.FieldAndMapping`. | Scenario JSON and Table 4. |
| T06 | An optional network source may fail without crashing the runner or hiding the source state. | `optional-path-fallback`; relative paths resolved below a UNC base. | Embedded synthetic defaults remain active and the source row reports `fallback`; UNC work is time-bounded. | `SyntheticScenario.OptionalFallback`, `validate-network-source-bounds.ps1`, `validate-runner-resilience.ps1`. | Figure 1a and simulator smoke log. |
| T07 | The tested clinical view-model transaction must restore report and last-good review state after a failed refresh. | Constraint-table change and manual structure mapping outside TPS runtime. | Capture, complete preparation, guarded refresh, and rollback of PQMs, objectives, overview, DVH selection, and mapping metrics. | `validate-pqm-transaction-safety.ps1`, `ReviewWorkspaceHost.RefreshFallback`. | Validator pass and source-level transaction boundary; no live Eclipse fault-injection evidence. |
| T08 | Valid zero-valued OAR metrics and ordinary names such as `Trachea` must remain visible by default. | Zero achieved value, delimited partial-helper token, unavailable target/OAR. | Only explicit unavailable states or delimited helper tokens may auto-ignore a non-target. | `PqmDefaultReviewPolicy.Valid`, `.Partial`, `.Unavailable`. | Executable test log. |
| T09 | The simulator and its package must not reference proprietary TPS assemblies. | Public Release binaries and ZIP input. | Project and assembly references exclude `VMS.TPS.*` and EsapiEssentials; no symbols are packaged. | `validate-vendor-free-assemblies.ps1`, `test-simulator.ps1`, `package-public-simulator.ps1`. | Release ZIP inventory and SHA-256. |
| T10 | RayStation is an illustrative read-only contract adapter, not a commissioned implementation. | Fake current-context objects only. | Unmapped judgments remain `not-evaluated`; anomalous DVHs are rejected; output omits clinical identifiers; write-like calls are rejected by AST inspection; one Python export must pass C# schema validation. | Twenty Python tests and `test-raystation-contract.ps1`. | Test log, generated fake-context JSON, and zero-issue C# validator output; no runtime or UI-ingestion claim. |
| T11 | The synthetic PDF must contain all mapped review domains and an unambiguous safety mark. | `mixed-review` report action. | Snapshot-to-report mapping plus fixed synthetic watermark. | `ReviewReportMapping.Sections`, `.Watermark`, `.PdfSmoke`, `.PdfContent`; simulator smoke. | Generated synthetic PDF. |
| T12 | The public release archive must be reproducible from the tagged source. | Two consecutive package builds and a clean tagged checkout. | Sorted ZIP entries, fixed entry timestamp, deterministic Release compilation, byte-identical source scenarios. | `build-public-release.ps1`. | Release `.sha256` asset and clean-checkout comparison. |

## Interpretation

Passing these checks means that the declared synthetic inputs produce the
declared software outputs in the documented environment. It does not validate
institutional constraints, Eclipse or RayStation runtime compatibility,
clinical error-detection performance, workflow efficiency, or patient-safety
benefit.
