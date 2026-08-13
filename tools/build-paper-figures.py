#!/usr/bin/env python3
"""Build all reproducible figures for the ClearPlan Technical Note."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPT_DIR = ROOT / "tools" / "paper_figures"
SCRIPTS = (
    "capture_figure_1.py",
    "generate_figure_2_architecture.py",
    "generate_figure_3_comparison.py",
)


def main() -> None:
    for script_name in SCRIPTS:
        subprocess.run(
            [sys.executable, str(SCRIPT_DIR / script_name)],
            cwd=ROOT,
            check=True,
        )
    print("Built three ClearPlan Technical Note figures.")


if __name__ == "__main__":
    main()
