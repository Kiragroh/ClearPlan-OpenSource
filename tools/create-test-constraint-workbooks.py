#!/usr/bin/env python3
"""Create public-safe XLSX fixtures for the package-free C# reader tests."""

from __future__ import annotations

import argparse
from pathlib import Path

import xlsxwriter


TABLE_HEADERS = [
    "table_id",
    "display_name",
    "active",
    "is_plan_sum",
    "fx_min",
    "fx_max",
    "dpf_min_gy",
    "dpf_max_gy",
    "total_dose_min_gy",
    "total_dose_max_gy",
    "site",
    "regime",
    "source",
]

CONSTRAINT_HEADERS = [
    "constraint_id",
    "table_id",
    "structure_id",
    "structure_name",
    "metric",
    "unit",
    "comparator",
    "goal",
    "variation",
    "priority",
    "source",
    "comment",
    "legacy_metadata",
]

STRUCTURE_HEADERS = [
    "structure_id",
    "canonical_name",
    "active",
    "laterality",
    "aliases",
    "side_aliases_left",
    "side_aliases_right",
    "dicom_type",
    "codes",
]


def write_book(
    path: Path,
    table_headers: list[str],
    table_rows: list[list[object]],
    constraint_rows: list[list[object]],
    structure_rows: list[list[object]],
) -> None:
    workbook = xlsxwriter.Workbook(path)
    header_format = workbook.add_format(
        {"bold": True, "bg_color": "#DCE6F1", "border": 1}
    )
    for name, headers, rows in [
        ("Tables", table_headers, table_rows),
        ("Constraints", CONSTRAINT_HEADERS, constraint_rows),
        ("Structures", STRUCTURE_HEADERS, structure_rows),
    ]:
        worksheet = workbook.add_worksheet(name)
        worksheet.write_row(0, 0, headers, header_format)
        for row_index, row in enumerate(rows, start=1):
            worksheet.write_row(row_index, 0, row)
        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, max(len(rows), 1), len(headers) - 1)
    readme = workbook.add_worksheet("README")
    readme.write("A1", "ClearPlan Testkatalog")
    readme.write("A2", "Größe, Rückenmark, ≤, ≥ und cm³ prüfen Unicode.")
    workbook.close()


def create_fixtures(output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    valid_tables = [
        [
            "fixture-table",
            "Öffentlicher Test",
            True,
            False,
            3,
            5,
            5.0,
            8.0,
            24.0,
            40.0,
            "Beispiel",
            "Hypofraktionierung",
            "Fixture",
        ]
    ]
    valid_constraints = [
        [
            "c-1",
            "fixture-table",
            "SpinalCord",
            "SpinalCord",
            "D0.03cc",
            "Gy",
            "≤",
            18.5,
            20.0,
            1,
            "Fixture",
            "Größe des Rückenmarks",
            '{"symbol":"≤","unit":"cm³"}',
        ],
        [
            "c-2",
            "fixture-table",
            "Kidney_L/R",
            "Kidney_L/R",
            "V20Gy",
            "%",
            "<=",
            30,
            35,
            2,
            "Fixture",
            "Nierenprüfung",
            '{"aliases":"Niere_L|Niere_R"}',
        ],
    ]
    valid_structures = [
        [
            "SpinalCord",
            "SpinalCord",
            True,
            "",
            "Rückenmark|Myelon",
            "",
            "",
            "ORGAN",
            "T-A7010",
        ],
        [
            "Kidney_L/R",
            "Kidney_L/R",
            True,
            "L/R",
            "Kidney|Niere",
            "Kidney_L|Niere_L",
            "Kidney_R|Niere_R",
            "ORGAN",
            "T-71000",
        ],
    ]
    write_book(
        output_dir / "excel-catalog-fixture.xlsx",
        TABLE_HEADERS,
        valid_tables,
        valid_constraints,
        valid_structures,
    )
    write_book(
        output_dir / "excel-catalog-missing-header.xlsx",
        ["wrong_header", *TABLE_HEADERS[1:]],
        valid_tables,
        valid_constraints,
        valid_structures,
    )
    write_book(
        output_dir / "excel-catalog-invalid-links.xlsx",
        TABLE_HEADERS,
        [valid_tables[0], valid_tables[0]],
        [
            [
                "c-broken",
                "unknown-table",
                "SpinalCord",
                "SpinalCord",
                "Dmean",
                "Gy",
                "<=",
                20,
                25,
                1,
                "Fixture",
                "Broken relationship",
                "{}",
            ]
        ],
        valid_structures,
    )


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    create_fixtures(args.output_dir.resolve())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
