#!/usr/bin/env python3
"""Collect bounded, patient-free evidence used by the ClearPlan technical note."""

from __future__ import annotations

import hashlib
import json
import re
import subprocess
import sys
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCENARIO_DIR = ROOT / "ClearPlan.Simulator" / "Scenarios"
MIXED_SCENARIO = SCENARIO_DIR / "mixed-review.json"
WORKBOOK = (
    ROOT
    / "ClearPlan.Script"
    / "Distribution"
    / "ConstraintTemplates"
    / "ClearPlan_DefaultConstraints.xlsx"
)
OUTPUT = ROOT / "paper" / "evidence" / "technical_note_evidence.json"


def run_checked(command: list[str]) -> str:
    completed = subprocess.run(
        command,
        cwd=ROOT,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    combined = "\n".join(
        part.strip() for part in (completed.stdout, completed.stderr) if part.strip()
    )
    if completed.returncode != 0:
        raise RuntimeError(
            "Evidence command failed ({}):\n{}".format(
                completed.returncode, combined
            )
        )
    return combined


def match_int(pattern: str, text: str, label: str) -> int:
    match = re.search(pattern, text, flags=re.IGNORECASE)
    if not match:
        raise RuntimeError("Could not extract {} from command output.".format(label))
    return int(match.group(1))


def status_counts(rows: list[dict], field: str = "status") -> dict[str, int]:
    counts = Counter(str(row.get(field, "missing")).lower() for row in rows)
    return {key: counts[key] for key in sorted(counts)}


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def main() -> None:
    scenario_paths = sorted(SCENARIO_DIR.glob("*.json"))
    scenarios = [json.loads(path.read_text(encoding="utf-8")) for path in scenario_paths]
    if not scenarios or not all(item.get("synthetic") is True for item in scenarios):
        raise RuntimeError("Every publication scenario must be explicitly synthetic.")

    mixed = json.loads(MIXED_SCENARIO.read_text(encoding="utf-8"))

    test_executable = (
        ROOT / "artifacts" / "bin" / "Release" / "ClearPlan.Core.Tests.exe"
    )
    if not test_executable.is_file():
        raise RuntimeError(
            "Build ClearPlan.Core.Tests in Release configuration before "
            "collecting paper evidence."
        )

    csharp_output = run_checked([str(test_executable)])
    csharp_total = match_int(r"(\d+)\s+tests?,\s+0\s+failures", csharp_output, "C# test count")

    raystation_output = run_checked(
        [
            sys.executable,
            "-m",
            "unittest",
            "discover",
            "-s",
            str(ROOT / "examples" / "raystation" / "tests"),
            "-v",
        ]
    )
    raystation_total = match_int(
        r"Ran\s+(\d+)\s+tests?", raystation_output, "RayStation test count"
    )
    if not re.search(r"\bOK\b", raystation_output):
        raise RuntimeError("RayStation adapter tests did not report OK.")

    workbook_output = run_checked(
        [sys.executable, str(ROOT / "tools" / "validate-constraint-workbook.py")]
    )
    workbook_tables = match_int(r"tables=(\d+)", workbook_output, "workbook table count")
    workbook_constraints = match_int(
        r"constraints=(\d+)", workbook_output, "workbook constraint count"
    )
    workbook_structures = match_int(
        r"structures=(\d+)", workbook_output, "workbook structure count"
    )

    evidence = {
        "schemaVersion": 1,
        "generatedUtc": datetime.now(timezone.utc).replace(microsecond=0).isoformat(),
        "scope": (
            "Software verification and deterministic synthetic demonstration only; "
            "no patient or clinical outcome data."
        ),
        "repositoryRelease": "v3.1.0",
        "softwareVerification": {
            "csharpTests": {
                "command": "artifacts/bin/<configuration>/ClearPlan.Core.Tests.exe",
                "configuration": "Release",
                "total": csharp_total,
                "failures": 0,
            },
            "illustrativeRayStationAdapterTests": {
                "command": "python -m unittest discover -s examples/raystation/tests -v",
                "total": raystation_total,
                "failures": 0,
                "runtimeValidation": False,
            },
        },
        "constraintWorkbook": {
            "path": WORKBOOK.relative_to(ROOT).as_posix(),
            "sha256": sha256(WORKBOOK),
            "tables": workbook_tables,
            "constraints": workbook_constraints,
            "structures": workbook_structures,
        },
        "syntheticScenarios": {
            "count": len(scenarios),
            "scenarioIds": [item["scenarioId"] for item in scenarios],
            "allExplicitlySynthetic": True,
            "sourceFiles": [path.relative_to(ROOT).as_posix() for path in scenario_paths],
        },
        "mixedReviewScenario": {
            "scenarioId": mixed["scenarioId"],
            "seed": mixed["seed"],
            "scenarioGeneratedUtc": mixed["generatedUtc"],
            "sourceCount": len(mixed["sources"]),
            "sourceStatusCounts": status_counts(mixed["sources"]),
            "planCount": len(mixed["plans"]),
            "pqm": {
                "rowCount": len(mixed["pqmRows"]),
                "statusCounts": status_counts(mixed["pqmRows"]),
                "achievedValues": [
                    {
                        "templateCode": row["templateCode"],
                        "value": row["achievedValue"],
                        "unit": row["unit"],
                        "status": row["status"],
                    }
                    for row in mixed["pqmRows"]
                ],
            },
            "planCheck": {
                "rowCount": len(mixed["planCheckRows"]),
                "statusCounts": status_counts(mixed["planCheckRows"]),
                "injectedNonPassRows": sum(
                    1
                    for row in mixed["planCheckRows"]
                    if row["status"] != "pass"
                    and row["message"].startswith("Injected synthetic finding:")
                ),
            },
            "fields": {
                "rowCount": len(mixed["fieldRows"]),
                "idStatusCounts": status_counts(mixed["fieldRows"], "idStatus"),
                "nameStatusCounts": status_counts(mixed["fieldRows"], "nameStatus"),
                "suggestedNames": [row["suggestedName"] for row in mixed["fieldRows"]],
            },
            "mappings": {
                "rowCount": len(mixed["structureMappings"]),
                "statusCounts": status_counts(mixed["structureMappings"]),
            },
            "dvh": {
                "seriesCount": len(mixed["dvhSeries"]),
                "selectedSeriesCount": sum(
                    1 for series in mixed["dvhSeries"] if series["selected"]
                ),
                "pointsPerSeries": {
                    series["stableId"]: len(series["points"])
                    for series in mixed["dvhSeries"]
                },
            },
        },
        "figureInputs": {
            "overview": "real simulator capture from mixed-review",
            "architecture": "editable vector diagram generated from repository source",
            "comparison": "derived only from checked-in synthetic scenarios",
        },
    }

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT.write_text(
        json.dumps(evidence, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
    )
    print("Wrote {}".format(OUTPUT))


if __name__ == "__main__":
    main()
