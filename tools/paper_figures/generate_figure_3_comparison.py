#!/usr/bin/env python3
"""Generate the deterministic DVH and status comparison for Figure 3."""

from __future__ import annotations

import json
from collections import Counter
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
from matplotlib.lines import Line2D


ROOT = Path(__file__).resolve().parents[2]
SCENARIO_DIR = ROOT / "ClearPlan.Simulator" / "Scenarios"
FIGURE_DIR = ROOT / "paper" / "figures"
STEM = FIGURE_DIR / "Figure_3_Synthetic_comparison"

COLORS = {
    "PTV_60": "#D1495B",
    "SpinalCord": "#00798C",
    "Parotid_L": "#E69F00",
}
STATUS_COLORS = {
    "pass": "#009E73",
    "variation": "#E69F00",
    "fail": "#D55E00",
}


def load(name: str) -> dict:
    return json.loads((SCENARIO_DIR / f"{name}.json").read_text(encoding="utf-8"))


def series_by_id(snapshot: dict) -> dict:
    return {row["structureId"]: row for row in snapshot["dvhSeries"]}


def status_counts(snapshot: dict) -> dict[str, Counter]:
    return {
        "PQM": Counter(row["status"] for row in snapshot["pqmRows"]),
        "PlanCheck": Counter(row["status"] for row in snapshot["planCheckRows"]),
        "Field names": Counter(row["nameStatus"] for row in snapshot["fieldRows"]),
        "Mappings": Counter(row["status"] for row in snapshot["structureMappings"]),
    }


def main() -> None:
    baseline = load("baseline-pass")
    mixed = load("mixed-review")
    baseline_series = series_by_id(baseline)
    mixed_series = series_by_id(mixed)
    counts = status_counts(mixed)

    plt.rcParams.update(
        {
            "font.family": "serif",
            "font.serif": ["Times New Roman", "DejaVu Serif"],
            "font.size": 7.5,
            "axes.labelsize": 8,
            "xtick.labelsize": 7,
            "ytick.labelsize": 7,
            "legend.fontsize": 6.4,
            "axes.spines.top": False,
            "axes.spines.right": False,
            "svg.fonttype": "none",
            "pdf.fonttype": 42,
        }
    )
    fig, (dvh_ax, status_ax) = plt.subplots(
        1,
        2,
        figsize=(6.65, 3.4),
        gridspec_kw={"width_ratios": [1.45, 1.0]},
    )

    labels = {
        "PTV_60": "PTV",
        "SpinalCord": "Spinal cord",
        "Parotid_L": "Parotid left",
    }
    for structure_id in ("PTV_60", "SpinalCord", "Parotid_L"):
        color = COLORS[structure_id]
        before = baseline_series[structure_id]["points"]
        after = mixed_series[structure_id]["points"]
        dvh_ax.plot(
            [point["doseGy"] for point in before],
            [point["volumePercent"] for point in before],
            color=color,
            linewidth=1.0,
            linestyle="--",
            alpha=0.75,
        )
        dvh_ax.plot(
            [point["doseGy"] for point in after],
            [point["volumePercent"] for point in after],
            color=color,
            linewidth=1.5,
            linestyle="-",
        )
    dvh_ax.set_xlabel("Dose (Gy)")
    dvh_ax.set_ylabel("Relative volume (%)")
    dvh_ax.set_xlim(0, 70)
    dvh_ax.set_ylim(0, 101)
    dvh_ax.grid(True, color="#D9E1E8", linewidth=0.5, alpha=0.8)
    structure_handles = [
        Line2D(
            [0],
            [0],
            color=COLORS[structure_id],
            linewidth=1.5,
            label=labels[structure_id],
        )
        for structure_id in ("PTV_60", "SpinalCord", "Parotid_L")
    ]
    scenario_handles = [
        Line2D(
            [0],
            [0],
            color="#475569",
            linewidth=1.0,
            linestyle="--",
            label="Baseline",
        ),
        Line2D(
            [0],
            [0],
            color="#475569",
            linewidth=1.5,
            linestyle="-",
            label="Mixed review",
        ),
    ]
    structure_legend = dvh_ax.legend(
        handles=structure_handles,
        title="Structure",
        frameon=False,
        loc="lower left",
        bbox_to_anchor=(0.0, 0.01),
        handlelength=1.8,
    )
    dvh_ax.add_artist(structure_legend)
    dvh_ax.legend(
        handles=scenario_handles,
        title="Synthetic scenario",
        frameon=False,
        loc="lower left",
        bbox_to_anchor=(0.34, 0.01),
        handlelength=1.8,
    )
    dvh_ax.text(-0.11, 1.03, "a", transform=dvh_ax.transAxes, fontweight="bold", fontsize=9.5)

    layer_names = list(counts)
    x_values = np.arange(len(layer_names))
    bottom = np.zeros(len(layer_names))
    for status, label, hatch in (
        ("pass", "Pass", ".."),
        ("variation", "Variation", "///"),
        ("fail", "Fail", "xx"),
    ):
        values = np.array([counts[layer][status] for layer in layer_names])
        bars = status_ax.bar(
            x_values,
            values,
            bottom=bottom,
            width=0.64,
            color=STATUS_COLORS[status],
            edgecolor="#263238",
            linewidth=0.35,
            hatch=hatch,
            label=label,
        )
        for bar, value, base in zip(bars, values, bottom):
            if value:
                status_ax.text(
                    bar.get_x() + bar.get_width() / 2,
                    base + value / 2,
                    str(int(value)),
                    ha="center",
                    va="center",
                    color="white" if status != "variation" else "#3B2F00",
                    fontsize=7,
                    fontweight="bold",
                )
        bottom += values
    status_ax.set_ylabel("Number of rows")
    status_ax.set_xticks(x_values)
    display_layer_names = [
        "Field names\n(IDs pass)" if name == "Field names" else name
        for name in layer_names
    ]
    status_ax.set_xticklabels(display_layer_names, rotation=15, ha="right")
    status_ax.set_ylim(0, max(bottom) + 1)
    status_ax.set_yticks(range(0, int(max(bottom)) + 2, 2))
    status_ax.legend(frameon=False, loc="upper right")
    status_ax.text(-0.14, 1.03, "b", transform=status_ax.transAxes, fontweight="bold", fontsize=9.5)

    fig.subplots_adjust(left=0.075, right=0.99, top=0.96, bottom=0.20, wspace=0.34)
    FIGURE_DIR.mkdir(parents=True, exist_ok=True)
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
    print(f"Wrote synthetic comparison figure: {STEM}.pdf")


if __name__ == "__main__":
    main()
