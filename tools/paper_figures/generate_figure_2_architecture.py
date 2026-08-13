#!/usr/bin/env python3
"""Generate the editable ClearPlan architecture diagram for Figure 2."""

from __future__ import annotations

from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch


ROOT = Path(__file__).resolve().parents[2]
FIGURE_DIR = ROOT / "paper" / "figures"
STEM = FIGURE_DIR / "Figure_2_ClearPlan_architecture"

NAVY = "#17324D"
BLUE = "#277DA1"
TEAL = "#2A9D8F"
AMBER = "#E9C46A"
PALE_BLUE = "#EAF3F8"
PALE_TEAL = "#E8F5F2"
PALE_AMBER = "#FFF6DC"
GRAY = "#5E6B75"
PALE_GRAY = "#F3F5F7"


def box(ax, x, y, width, height, text, face, edge, dashed=False):
    patch = FancyBboxPatch(
        (x, y),
        width,
        height,
        boxstyle="round,pad=0.015,rounding_size=0.025",
        linewidth=1.0,
        edgecolor=edge,
        facecolor=face,
        linestyle="--" if dashed else "-",
    )
    ax.add_patch(patch)
    ax.text(
        x + width / 2,
        y + height / 2,
        text,
        ha="center",
        va="center",
        fontsize=7.2,
        color=NAVY,
        linespacing=1.25,
    )


def arrow(ax, start, end, dashed=False, connectionstyle="arc3"):
    ax.add_patch(
        FancyArrowPatch(
            start,
            end,
            arrowstyle="-|>",
            mutation_scale=9,
            linewidth=0.9,
            color=GRAY,
            linestyle="--" if dashed else "-",
            connectionstyle=connectionstyle,
            shrinkA=4,
            shrinkB=4,
        )
    )


def main() -> None:
    FIGURE_DIR.mkdir(parents=True, exist_ok=True)
    plt.rcParams.update(
        {
            "font.family": "serif",
            "font.serif": ["Times New Roman", "DejaVu Serif"],
            "font.size": 9,
            "svg.fonttype": "none",
            "pdf.fonttype": 42,
        }
    )
    fig, ax = plt.subplots(figsize=(6.65, 4.25))
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")

    box(ax, 0.02, 0.76, 0.15, 0.10, "RefDB JSON\nstructures + aliases", PALE_BLUE, BLUE)
    box(ax, 0.02, 0.61, 0.15, 0.10, "Excel XLSX\nstructures + tables + rules", PALE_BLUE, BLUE)
    box(ax, 0.22, 0.68, 0.16, 0.14, "Read-only source adapters\nschema + relationship\nvalidation", PALE_BLUE, BLUE)
    box(ax, 0.43, 0.68, 0.17, 0.14, "Validated internal catalog\nnormalized units,\ncomparators + provenance", PALE_TEAL, TEAL)
    box(ax, 0.65, 0.68, 0.16, 0.14, "Deterministic resolution\naliases, laterality +\nplan-context table selection", PALE_TEAL, TEAL)

    box(ax, 0.02, 0.39, 0.22, 0.10, "Eclipse / ESAPI\nclinical read-only host", PALE_AMBER, "#B7791F")
    box(ax, 0.02, 0.24, 0.22, 0.10, "Patient-free simulator\nchecked-in scenarios", PALE_AMBER, "#B7791F")
    box(
        ax,
        0.02,
        0.09,
        0.22,
        0.10,
        "Illustrative RayStation adapter\nunvalidated example",
        PALE_GRAY,
        GRAY,
        dashed=True,
    )

    box(
        ax,
        0.29,
        0.39,
        0.20,
        0.10,
        "Clinical snapshot projector\nfrom existing local results",
        PALE_TEAL,
        TEAL,
    )
    box(
        ax,
        0.29,
        0.24,
        0.20,
        0.10,
        "Synthetic scenario builder\nanalytic DVHs + fixed findings",
        PALE_TEAL,
        TEAL,
    )
    box(
        ax,
        0.29,
        0.09,
        0.20,
        0.10,
        "Illustrative JSON producer\nno released UI loader",
        PALE_GRAY,
        GRAY,
        dashed=True,
    )
    box(
        ax,
        0.55,
        0.20,
        0.20,
        0.23,
        "Common ReviewSnapshot v1\ncontract + validation\nPQM · PlanCheck · fields\nmappings · DVH · sources",
        PALE_TEAL,
        TEAL,
    )
    box(ax, 0.80, 0.34, 0.18, 0.12, "Shared review UI\noverview + focused pages", PALE_BLUE, BLUE)
    box(ax, 0.80, 0.20, 0.18, 0.09, "Simulator snapshot-to-PDF\nmapper", PALE_BLUE, BLUE)
    box(ax, 0.80, 0.06, 0.18, 0.09, "Clinical legacy ReportData\nPDF path", PALE_BLUE, BLUE)

    arrow(ax, (0.17, 0.81), (0.22, 0.78))
    arrow(ax, (0.17, 0.66), (0.22, 0.72))
    arrow(ax, (0.38, 0.75), (0.43, 0.75))
    arrow(ax, (0.60, 0.75), (0.65, 0.75))
    arrow(
        ax,
        (0.73, 0.68),
        (0.39, 0.49),
        connectionstyle="arc3,rad=0.15",
    )

    arrow(ax, (0.24, 0.44), (0.29, 0.44))
    arrow(ax, (0.24, 0.29), (0.29, 0.29))
    arrow(ax, (0.24, 0.14), (0.29, 0.14), dashed=True)
    arrow(ax, (0.49, 0.44), (0.55, 0.37))
    arrow(ax, (0.49, 0.29), (0.55, 0.31))
    arrow(ax, (0.49, 0.14), (0.55, 0.25), dashed=True)
    arrow(ax, (0.75, 0.37), (0.80, 0.40))
    arrow(ax, (0.75, 0.26), (0.80, 0.245))
    arrow(
        ax,
        (0.49, 0.41),
        (0.80, 0.105),
        connectionstyle="arc3,rad=0.28",
    )

    ax.text(
        0.02,
        0.93,
        "Constraint configuration",
        fontsize=8.6,
        fontweight="bold",
        color=NAVY,
    )
    ax.text(
        0.02,
        0.015,
        "Solid arrows: implemented path     Dashed: contract-producing example only; no released UI ingestion     No adapter writes to a treatment plan.",
        fontsize=6.4,
        color=GRAY,
    )

    fig.subplots_adjust(left=0.015, right=0.985, top=0.985, bottom=0.03)
    for extension in ("svg", "pdf", "png"):
        output_path = Path(f"{STEM}.{extension}")
        fig.savefig(
            output_path,
            dpi=300,
            bbox_inches="tight",
            pad_inches=0.03,
            facecolor="white",
        )
        if extension == "svg":
            svg = output_path.read_text(encoding="utf-8")
            output_path.write_text(
                "\n".join(line.rstrip() for line in svg.splitlines()) + "\n",
                encoding="utf-8",
            )
    plt.close(fig)
    print(f"Wrote editable architecture figure: {STEM}.svg")


if __name__ == "__main__":
    main()
