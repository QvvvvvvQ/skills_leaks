#!/usr/bin/env python3
"""
diagrams.py - Raster chart generation for document embedding.

Intended for chart types that Word cannot render natively (heatmaps, radar,
waterfall, multi-axis composites, etc.).  For standard pie / bar / line charts,
prefer the native OpenXML chart APIs demonstrated in Blueprint.cs.

All public functions accept explicit data and styling arguments so that the
caller controls content entirely.  When *palette* is omitted the Sage
scheme from chromatics.py is used by default.
"""

import os
import sys
import textwrap

import matplotlib
matplotlib.use("Agg")          # headless backend, no GUI dependency
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np

# ---------------------------------------------------------------------------
# Internal: resolve palette colours
# ---------------------------------------------------------------------------

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

def _load_palette(name: str | None = None) -> list[str]:
    """Return a list of CSS-hex accent colours from *chromatics.py*."""
    sys.path.insert(0, _SCRIPT_DIR)
    try:
        import chromatics
    finally:
        sys.path.pop(0)
    palette = chromatics.fetch_palette(name or "Sage")
    orn = palette["Ornamental"]
    return [
        "#" + orn["Subtle"]["Primary"],
        "#" + orn["Subtle"]["Secondary"],
        "#" + orn["Subtle"]["Highlight"],
        "#" + orn["Bold"]["Primary"],
        "#" + orn["Bold"]["Secondary"],
        "#" + orn["Bold"]["Highlight"],
    ]

# ---------------------------------------------------------------------------
# Shared styling helpers
# ---------------------------------------------------------------------------

_RC_DEFAULTS = {
    "font.sans-serif": ["Helvetica Neue", "Arial", "sans-serif"],
    "axes.unicode_minus": False,
    "figure.facecolor": "white",
}

def _apply_rc():
    for k, v in _RC_DEFAULTS.items():
        plt.rcParams[k] = v

def _minimal_axes(ax, grid_axis="y"):
    """Strip chrome to leave only the data visible."""
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    for side in ("left", "bottom"):
        ax.spines[side].set_color("#D0D0D0")
    if grid_axis:
        getattr(ax, f"{grid_axis}axis").grid(
            True, linestyle=":", linewidth=0.6, alpha=0.45, color="#B0B0B0"
        )
        ax.set_axisbelow(True)

def _save(fig, output_path: str, resolution: int = 200):
    fig.savefig(output_path, dpi=resolution, bbox_inches="tight",
                facecolor="white", edgecolor="none")
    plt.close(fig)
    return output_path

# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def grouped_bars(
    categories: list[str],
    series: dict[str, list[float]],
    *,
    ylabel: str = "",
    output_path: str = "chart_grouped_bars.png",
    palette: str | None = None,
    figsize: tuple[float, float] = (9, 5),
) -> str:
    """Grouped vertical bar chart.  *series* maps label → values."""
    _apply_rc()
    colors = _load_palette(palette)
    fig, ax = plt.subplots(figsize=figsize)

    n_groups = len(categories)
    n_series = len(series)
    bar_w = 0.7 / n_series
    x = np.arange(n_groups)

    for idx, (label, values) in enumerate(series.items()):
        offset = (idx - n_series / 2 + 0.5) * bar_w
        bars = ax.bar(x + offset, values, bar_w, label=label,
                      color=colors[idx % len(colors)], edgecolor="none")
        for b in bars:
            h = b.get_height()
            ax.text(b.get_x() + b.get_width() / 2, h,
                    f"{h:g}", ha="center", va="bottom",
                    fontsize=8, color="#555555")

    ax.set_xticks(x)
    ax.set_xticklabels(categories, fontsize=9)
    if ylabel:
        ax.set_ylabel(ylabel, fontsize=10, color="#404040")
    ax.legend(frameon=False, fontsize=9)
    _minimal_axes(ax)
    fig.tight_layout()
    return _save(fig, output_path)


def trend_lines(
    x_labels: list[str],
    series: dict[str, list[float]],
    *,
    ylabel: str = "",
    output_path: str = "chart_trend.png",
    palette: str | None = None,
    figsize: tuple[float, float] = (9.5, 4.5),
    fill_first: bool = True,
) -> str:
    """Multi-series line chart with optional area fill on the first series."""
    _apply_rc()
    colors = _load_palette(palette)
    fig, ax = plt.subplots(figsize=figsize)

    markers = ("o", "D", "^", "s", "v", "P")
    for idx, (label, values) in enumerate(series.items()):
        c = colors[idx % len(colors)]
        ax.plot(x_labels, values, marker=markers[idx % len(markers)],
                markersize=4, linewidth=1.8, color=c, label=label)
        if fill_first and idx == 0:
            ax.fill_between(x_labels, values, alpha=0.10, color=c)

    if ylabel:
        ax.set_ylabel(ylabel, fontsize=10, color="#404040")
    ax.legend(frameon=False, fontsize=9, loc="upper left")
    _minimal_axes(ax)
    plt.xticks(rotation=40, ha="right", fontsize=8)
    fig.tight_layout()
    return _save(fig, output_path)


def stacked_areas(
    x_labels: list[str],
    series: dict[str, list[float]],
    *,
    ylabel: str = "",
    output_path: str = "chart_stacked.png",
    palette: str | None = None,
    figsize: tuple[float, float] = (9, 4.5),
) -> str:
    """Stacked area chart."""
    _apply_rc()
    colors = _load_palette(palette)
    fig, ax = plt.subplots(figsize=figsize)

    labels = list(series.keys())
    data = list(series.values())
    ax.stackplot(x_labels, *data, labels=labels,
                 colors=colors[:len(labels)], alpha=0.80)

    if ylabel:
        ax.set_ylabel(ylabel, fontsize=10, color="#404040")
    ax.legend(loc="upper left", frameon=False, fontsize=9)
    _minimal_axes(ax)
    fig.tight_layout()
    return _save(fig, output_path)


def horizontal_bars(
    categories: list[str],
    series: dict[str, list[float]],
    *,
    xlabel: str = "",
    output_path: str = "chart_hbars.png",
    palette: str | None = None,
    figsize: tuple[float, float] = (9, 4.5),
) -> str:
    """Grouped horizontal bar chart."""
    _apply_rc()
    colors = _load_palette(palette)
    fig, ax = plt.subplots(figsize=figsize)

    n_cat = len(categories)
    n_ser = len(series)
    bar_h = 0.7 / n_ser
    y = np.arange(n_cat)

    for idx, (label, values) in enumerate(series.items()):
        offset = (idx - n_ser / 2 + 0.5) * bar_h
        ax.barh(y + offset, values, bar_h, label=label,
                color=colors[idx % len(colors)], edgecolor="none")

    if xlabel:
        ax.set_xlabel(xlabel, fontsize=10, color="#404040")
    ax.set_yticks(y)
    ax.set_yticklabels(categories, fontsize=9)
    ax.legend(frameon=False, fontsize=9, loc="lower right")
    _minimal_axes(ax, grid_axis="x")
    fig.tight_layout()
    return _save(fig, output_path)


def donut(
    labels: list[str],
    values: list[float],
    *,
    output_path: str = "chart_donut.png",
    palette: str | None = None,
    figsize: tuple[float, float] = (6.5, 5.5),
) -> str:
    """Donut (ring) chart — a modern alternative to the basic pie chart."""
    _apply_rc()
    colors = _load_palette(palette)
    fig, ax = plt.subplots(figsize=figsize)

    wedges, texts, pcts = ax.pie(
        values, labels=labels, colors=colors[:len(labels)],
        autopct="%1.0f%%", startangle=140, pctdistance=0.78,
        wedgeprops={"width": 0.45, "edgecolor": "white", "linewidth": 1.5},
    )
    for t in texts:
        t.set_fontsize(9)
        t.set_color("#404040")
    for p in pcts:
        p.set_fontsize(8)
        p.set_color("#333333")

    ax.set_aspect("equal")
    fig.tight_layout()
    return _save(fig, output_path)


# ---------------------------------------------------------------------------
# Demonstration entry-point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    out = os.path.join(_SCRIPT_DIR, "demo_charts")
    os.makedirs(out, exist_ok=True)

    grouped_bars(
        ["Jan", "Feb", "Mar", "Apr"],
        {"Region A": [220, 245, 260, 310], "Region B": [180, 200, 190, 230]},
        ylabel="Revenue (K USD)",
        output_path=os.path.join(out, "demo_bars.png"),
    )
    trend_lines(
        [f"W{i}" for i in range(1, 13)],
        {"Signups": [50, 62, 71, 80, 95, 110, 108, 122, 135, 140, 155, 170],
         "Active":  [30, 38, 48, 55, 68, 78, 75, 88, 100, 105, 118, 130]},
        ylabel="Users",
        output_path=os.path.join(out, "demo_trend.png"),
    )
    stacked_areas(
        ["H1", "H2", "H3", "H4"],
        {"Cloud": [40, 55, 65, 80], "On-prem": [60, 50, 45, 35]},
        ylabel="Infra spend (K)",
        output_path=os.path.join(out, "demo_stacked.png"),
    )
    horizontal_bars(
        ["Latency", "Throughput", "Reliability", "Cost"],
        {"Current": [68, 72, 90, 55], "Goal": [85, 90, 95, 75]},
        xlabel="Score",
        output_path=os.path.join(out, "demo_hbars.png"),
    )
    donut(
        ["Enterprise", "SMB", "Consumer", "Partner"],
        [42, 26, 20, 12],
        output_path=os.path.join(out, "demo_donut.png"),
    )
    print(f"Demo charts written to {out}/")
