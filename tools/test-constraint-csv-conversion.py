#!/usr/bin/env python3
"""Small regression tests for the legacy CSV to XLSX normalization."""

from __future__ import annotations

import csv
import importlib.util
from pathlib import Path
from tempfile import TemporaryDirectory


def load_converter():
    converter_path = Path(__file__).with_name(
        "convert-constraint-csvs-to-xlsx.py"
    )
    spec = importlib.util.spec_from_file_location(
        "clearplan_constraint_converter", converter_path
    )
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load converter: {converter_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def write_legacy_csv(path: Path, structure_id: str, alias: str, code: str) -> None:
    headers = [
        "Structure IDs",
        "Structure Codes",
        "IDAliases",
        "CodeAliases",
        "DVH Objective",
        "Evaluation Point",
        "Variation",
        "Priority",
        "Source",
        "ZusatzInfo",
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=headers)
        writer.writeheader()
        writer.writerow(
            {
                "Structure IDs": structure_id,
                "Structure Codes": code,
                "IDAliases": alias,
                "CodeAliases": "ORGAN",
                "DVH Objective": "Dmax [Gy]",
                "Evaluation Point": "<= 42",
                "Variation": "45",
                "Priority": "1",
                "Source": "Regression test",
                "ZusatzInfo": "",
            }
        )


def test_case_insensitive_structure_merge() -> None:
    converter = load_converter()
    with TemporaryDirectory(prefix="clearplan-constraint-test-") as temp:
        root = Path(temp)
        first = root / "T_First.csv"
        second = root / "T_Second.csv"
        write_legacy_csv(first, "Cauda Equina", "Cauda", "T-D1100")
        write_legacy_csv(second, "Cauda equina", "CaudaEq", "T-D1101")

        tables, constraints, structures = converter.convert([first, second])

    assert len(tables) == 2
    assert len(constraints) == 2
    assert len(structures) == 1
    assert structures[0]["structure_id"] == "Cauda Equina"
    assert structures[0]["aliases"] == "Cauda|CaudaEq"
    assert structures[0]["codes"] == "T-D1100|T-D1101"
    assert {
        constraint["structure_id"] for constraint in constraints
    } == {"Cauda Equina"}


def main() -> int:
    test_case_insensitive_structure_merge()
    print("PASS case-insensitive legacy structure merge")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
