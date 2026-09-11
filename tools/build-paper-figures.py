#!/usr/bin/env python3
"""Build all reproducible figures for the ClearPlan Technical Note."""

from __future__ import annotations

import subprocess
import argparse
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPT_DIR = ROOT / "tools" / "paper_figures"
def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--capture-dir", type=Path, required=True,
                        help="Verified 1600 x 1000 publication-dual-layer application captures.")
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--node-executable", default="node")
    args = parser.parse_args()
    captures = args.capture_dir.resolve(strict=True)
    output = args.output_dir.resolve()
    subprocess.run([args.node_executable, str(SCRIPT_DIR / "compose_workspace.cjs"),
                    "--capture-dir", str(captures), "--output-dir", str(output)],
                   cwd=ROOT, check=True)
    subprocess.run([args.node_executable, str(SCRIPT_DIR / "generate_architecture.cjs"),
                    str(output)], cwd=ROOT, check=True)
    print("Built two Technical Note figures. Final release provenance and visual inspection remain separate gates.")


if __name__ == "__main__":
    main()
