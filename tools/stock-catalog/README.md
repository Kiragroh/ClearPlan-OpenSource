# Stock 2024 constraint compilation

The stock catalog is a maintainable fallback when an institutional RefDB is not connected. It is an editable example transcription of a user-authorized 2024 compilation, not a clinical recommendation or a newly validated guideline. The current distributed workbook contains eight tables, 495 executable OAR rules and 70 canonical structures. All stock tables require explicit reviewer confirmation before evaluation. Matching the fraction count does not establish clinical applicability.

## Data contract

The extractor produces JSON arrays of objects for the existing ClearPlan Excel schema:

- `tables` -> `Tables`, with the optional `requires_confirmation` column set to `true`.
- `constraints` -> `Constraints`, preserving source comparator, metric input, output unit, goal and variation. `Mean`, `Max` and `Min` are represented by the existing parser's `Dmean`, `Dmax` and `Dmin` names without changing units or values.
- `structures` -> `Structures`, with verified TG263 primary-name spellings and explicitly reviewed local aliases. Codes are intentionally blank for matching.
- `reviewRows` -> `ReviewRows`, a non-executable audit queue. It must not be merged into `Constraints` by a workbook builder.
- `sources` -> `Sources`, a non-executable provenance ledger. The first three sheets are the only sheets imported as executable configuration by the existing loader.

`metadata.headers` gives the exact ordered column names, including required empty columns. The first data row must directly follow the header row on the three machine-readable sheets. Additional presentation sheets may precede these sheets without changing their names. Do not insert titles above their headers.

The released XLSX is the public maintenance artifact; no extraction intermediate is required to use or edit it. In the optional local source-audit workflow, `stock2024.public.json` is the label-neutral intermediate, while `stock2024.normalized.json` retains the original source filename and institution codes and stays in protected, ignored storage. Neither intermediate is assumed to exist in a clean public checkout. The public variant changes display/provenance labels to `LOCAL2024`; it retains the source hash and does not change a dose value, threshold, anatomical identifier or primary-reference code. `LOCAL2024` means local institutional constraints, not a publication.

## Included and excluded data

Eight current source sheets are included: `T_1Fx`, `T_3Fx`, `T_5Fx`, `T_8Fx`, `T_10Fx`, `T_15Fx`, `T_20Fx`, `T_Conv`. For the reviewed snapshot, 813 input rules reconcile to 495 executable OAR rules and 318 review-only rules. The counts by table are 58, 59, 67, 47, 48, 48, 50 and 118, respectively.

The conventional table has no invented fraction-count, dose-per-fraction or total-dose bounds. It must be selected and confirmed as appropriate. Site-specific `Mamma_5Fx`, `Mamma_15Fx` and `CHHiP_20Fx` are outside the first stock scope. Their indications, prescription definitions and contralateral/ipsilateral mappings require a distinct reviewed configuration. Legacy `z*`, poster layouts, sum tables and intermediate 2/4-fraction sheets are not treated as current executable inputs.

The review-only queue preserves original row locations and expressions for:

- Generic `Target` and `Tumor` roles with placeholder doses/tolerances, and generic `MG` target/body templates.
- Legacy CI/GI/HI expressions reporting dimensionless indices as percentages, DC and malformed objectives such as `V50Gy%[]`.
- Disease, sex, one-side, ipsilateral/contralateral, suborgan or cropped-volume conditions that are not represented by the selected whole-organ contour.
- Unclear anatomical definitions, missing source codes and unsupported/non-scalar goals.

No source value is silently corrected. In particular, a strict `<` comparator stays strict, complementary volume retains its lower-bound comparator and cc output, and a dose input in `V26Gy[%]` does not turn the volume output into Gy.

## Aliases and naming

`stock2024.mapping.json` is the declarative maintenance point. Canonical spellings were checked against column F of the official [AAPM TG263 2017-08-15 nomenclature worksheet](https://www.aapm.org/pubs/reports/rpt_263_supplemental/TG263_Nomenclature_Worksheet_20170815.xls). This spelling check does not prove equivalence of a local contour's anatomical extent.

Aliases come from `Glossar` and included executable source rows. A normalized alias shared by different canonical structures is excluded from all of them. Cropped or combined names are not collapsed to a whole-organ name. Examples explicitly excluded are `Lungs-PTV` -> `Lungs`, `Liver-PTV` -> `Liver`, `Normal Brain` -> `Brain`, spinal canal -> spinal cord, renal cortex -> both kidneys and chest wall -> ribs. Generic PTV/GTV/CTV aliases are never used to choose an arbitrary target. Laterality is retained.

Source anatomical codes are not enabled for matching: the source includes mixed or inconsistent codes. For example, reusing a whole-organ code for a PRV would create an unsafe fallback. Local maintainers may supply audited codes in their own workbook, but must review their code system, laterality and anatomical definition.

## Provenance and reproducibility

The following hashes identify historical extraction inputs, not the current release archive or workbook. The current distributed-XLSX hash belongs in the freshly collected software receipt and final release manifest.

Historical source SHA256: `b4624a60ec074635e181c8661b48b9698971134ae090c0989916944dd65df9d3`.

Official nomenclature worksheet SHA256: `5ff0b9e2ebf578793f6fa8f59c2357b3feb61fdc0f87171e9103a0495d93d150`.

The original source-code legend is DrawingML text, not worksheet cells. The extractor reads this real legend: `T` refers to Timmerman's *A Story of Hypofractionation and the Table on the Wall*, DOI `10.1016/j.ijrobp.2021.09.027`; `C_[A-D]_number` refers to CORSAIR and the source's numbered reference list. Original reference text is preserved. Ambiguous ranges such as `C_A_5-39` are not expanded by guessing. Reference metadata has not been independently corrected or primary-literature validated.

For an optional, locally authorized source audit, use Python with `openpyxl` and the exact approved input workbook:

```powershell
python tools/stock-catalog/extract_stock_catalog.py 'C:\path\authorized-source.xlsx' --output artifacts/stock-catalog/stock2024.normalized.json --public-output artifacts/stock-catalog/stock2024.public.json
python tools/stock-catalog/test_stock_catalog.py -v
```

The source workbook is never saved or modified. Its hash is checked before and after extraction. A changed hash fails closed until the mapping and source differences have been explicitly reviewed. XLSX creation belongs to the workbook authoring step, not this read-only extractor.

For routine maintenance, edit a local copy of the distributed Excel workbook's `Tables`, `Constraints` and `Structures` sheets and select it in ClearPlan settings. Keep stable IDs, explicit units, source references and `requires_confirmation=true` for unvalidated stock-derived tables. Use the configuration workspace to validate and version changes with actor, time and reason; restoring creates a new revision rather than erasing history. Adding a row to `ReviewRows` does not activate it. Promotion to `Constraints` requires a separately reviewed anatomical definition, numerical expression and applicable prescription scheme.

## Verification

Parser tests run without the original workbook or its private intermediate. They check metrics, units, comparator direction, unsupported expressions and conservative alias normalization. The separate `SnapshotTests` class additionally checks source accounting, provenance and label-only transformation, but requires the private normalized intermediate. A clean public checkout reports that class as skipped. Its unexecuted cases are not counted as passed and must not be reported as a fully reproduced source transcription.

The application-level `ExcelConstraintSource.Stock2024` test loads the distributed XLSX, checks its executable rows, the explicit-confirmation gate, unsafe alias exclusions and usable-RefDB priority. This route remains available in the public checkout independently of the skipped private source-audit class. Run from the repository root after a public Core build:

```powershell
python tools/stock-catalog/test_stock_catalog.py -v
.\artifacts\bin\Release\ClearPlan.Core.Tests.exe ExcelConstraintSource.Stock2024
```

The [public release workflow](../../docs/reproducibility.md) records actual execution totals, skipped tests/classes and current workbook hashes. Registration totals are not execution evidence. Final clinical applicability and local rule acceptance are separate from these parser/loader checks.

## Authoring and round-trip limits

The historical XLSX authoring helper used `@oai/artifact-tool`. That environment is not supplied as a reproducible public dependency route, and running that helper is not a prerequisite for the public build or routine maintenance. Do not describe a fresh public checkout as reproducing the original workbook-authoring process merely because it contains the finished XLSX.

An optional read-only `validate_workbook.py` can compare a workbook cell by cell with its matching label-neutral extraction JSON when both are available. It checks scalar values/types, headers, filters, frozen panes and public-label separation. This is distinct from routine validation of an edited local workbook and cannot be claimed on a checkout missing that matching oracle. Generated layouts require their own visual review; an earlier six-sheet review is not evidence for a newly edited version.

Routine users need neither Node nor Python to edit a local XLSX copy and select it in ClearPlan. Keep the reviewed workbook, configuration revision and local acceptance record together. Do not publish the source workbook, private normalized intermediate or institution-specific configuration to make an optional test group pass.
