#!/usr/bin/env python3
"""Capture the real mixed-review overview for Figure 1."""

from __future__ import annotations

import subprocess
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[2]
FIGURE_DIR = ROOT / "paper" / "figures"
CAPTURE_DIR = ROOT / "artifacts" / "paper-captures" / "mixed-review"
OUTPUT = FIGURE_DIR / "Figure_1_ClearPlan_overview.png"
PANEL_SPECS = (
    (
        "overview.png",
        (205, 140, 1585, 750),
        "a  Integrated overview: sources, active plan and PQM",
    ),
    (
        "fields.png",
        (205, 150, 1585, 430),
        "b  Read-only field identifier and name conformance",
    ),
    (
        "dvh.png",
        (205, 150, 1585, 980),
        "c  Wide dose-volume histogram review",
    ),
)
OUTPUT_WIDTH = 1995
SIDE_MARGIN = 24
PANEL_GAP = 18
LABEL_HEIGHT = 52


def label_font(size: int):
    candidates = (
        Path("C:/Windows/Fonts/arialbd.ttf"),
        Path("C:/Windows/Fonts/Arial.ttf"),
    )
    for candidate in candidates:
        if candidate.is_file():
            return ImageFont.truetype(str(candidate), size=size)
    return ImageFont.load_default()


def compose_publication_figure() -> Image.Image:
    inner_width = OUTPUT_WIDTH - (2 * SIDE_MARGIN)
    panels = []
    for filename, crop_box, label in PANEL_SPECS:
        source_path = CAPTURE_DIR / filename
        if not source_path.is_file():
            raise FileNotFoundError(f"Simulator capture is missing: {source_path}")
        with Image.open(source_path) as source:
            if source.size != (1600, 1000):
                raise RuntimeError(
                    f"Expected a 1600 x 1000 capture, found {source.size!r}: "
                    f"{source_path}"
                )
            panel = source.convert("RGB").crop(crop_box)
        scaled_height = round(panel.height * inner_width / panel.width)
        panels.append(
            (
                panel.resize(
                    (inner_width, scaled_height),
                    resample=Image.Resampling.LANCZOS,
                ),
                label,
            )
        )

    output_height = (
        2 * SIDE_MARGIN
        + sum(LABEL_HEIGHT + panel.height for panel, _ in panels)
        + PANEL_GAP * (len(panels) - 1)
    )
    composite = Image.new("RGB", (OUTPUT_WIDTH, output_height), "white")
    draw = ImageDraw.Draw(composite)
    font = label_font(30)
    navy = "#17324D"
    divider = "#CBD5E1"

    y = SIDE_MARGIN
    for index, (panel, label) in enumerate(panels):
        draw.text((SIDE_MARGIN, y + 7), label, font=font, fill=navy)
        y += LABEL_HEIGHT
        composite.paste(panel, (SIDE_MARGIN, y))
        draw.rectangle(
            (
                SIDE_MARGIN,
                y,
                SIDE_MARGIN + panel.width - 1,
                y + panel.height - 1,
            ),
            outline=divider,
            width=2,
        )
        y += panel.height
        if index < len(panels) - 1:
            y += PANEL_GAP
    return composite


def find_simulator() -> Path:
    candidates = [
        ROOT / "artifacts" / "simulator" / "Release" / "ClearPlan.Simulator.exe",
        ROOT / "artifacts" / "simulator" / "Debug" / "ClearPlan.Simulator.exe",
    ]
    for candidate in candidates:
        if (
            candidate.is_file()
            and (candidate.parent / "Scenarios" / "mixed-review.json").is_file()
        ):
            return candidate
    raise FileNotFoundError(
        "Build ClearPlan.Simulator before generating the paper figures."
    )


def main() -> None:
    simulator = find_simulator()
    CAPTURE_DIR.mkdir(parents=True, exist_ok=True)
    subprocess.run(
        [
            str(simulator),
            "--scenario",
            "mixed-review",
            "--capture-all",
            str(CAPTURE_DIR),
        ],
        cwd=simulator.parent,
        check=True,
    )

    image = compose_publication_figure()
    FIGURE_DIR.mkdir(parents=True, exist_ok=True)
    image.save(OUTPUT, dpi=(300, 300), optimize=True)

    if OUTPUT.stat().st_size < 250_000:
        raise RuntimeError("Figure 1 is unexpectedly small or empty.")
    print(f"Wrote real simulator composite: {OUTPUT}")


if __name__ == "__main__":
    main()
