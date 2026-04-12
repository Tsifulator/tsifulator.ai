"""
Plot Service — generates charts server-side using matplotlib.
Works on Railway without R installed. Excel users can get
professional plots without having R or any local tools.

Usage: Called from chat.py when Claude emits a 'create_plot' action
from an Excel context. The plot is base64-encoded and sent back
through the transfer system for the add-in to insert.
"""

import matplotlib
matplotlib.use("Agg")  # Non-interactive backend (no display needed)
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import numpy as np
import io
import base64
import json
from typing import Optional


# Professional styling
plt.rcParams.update({
    "figure.facecolor": "white",
    "axes.facecolor": "white",
    "axes.grid": True,
    "grid.alpha": 0.3,
    "font.family": "sans-serif",
    "font.size": 11,
    "axes.titlesize": 14,
    "axes.labelsize": 12,
})


def create_plot(
    plot_type: str,
    data: dict,
    title: str = "",
    x_label: str = "",
    y_label: str = "",
    width: int = 8,
    height: int = 5,
    style: str = "default",
    options: dict = None,
) -> dict:
    """Create a chart and return it as base64 PNG.

    Args:
        plot_type: "scatter", "bar", "line", "histogram", "pie", "box", "heatmap"
        data: {
            "x": [1,2,3,...],           # x values
            "y": [4,5,6,...],           # y values
            "labels": ["A","B",...],    # category labels (for bar/pie)
            "series": [                 # multiple series
                {"name": "Sales", "x": [...], "y": [...]},
                {"name": "Costs", "x": [...], "y": [...]},
            ],
            "values": [10, 20, 30],    # single series values (for histogram/pie)
        }
        title: Chart title
        x_label: X-axis label
        y_label: Y-axis label
        width: Figure width in inches
        height: Figure height in inches
        style: matplotlib style ("default", "seaborn-v0_8", "ggplot", "dark_background")
        options: Additional options dict

    Returns:
        {"success": True, "image_base64": "...", "mime_type": "image/png"}
    """
    if not options:
        options = {}

    try:
        if style != "default":
            plt.style.use(style)

        fig, ax = plt.subplots(figsize=(width, height))

        if plot_type == "scatter":
            _plot_scatter(ax, data, options)
        elif plot_type == "bar":
            _plot_bar(ax, data, options)
        elif plot_type == "line":
            _plot_line(ax, data, options)
        elif plot_type == "histogram":
            _plot_histogram(ax, data, options)
        elif plot_type == "pie":
            _plot_pie(ax, data, options)
        elif plot_type == "box":
            _plot_box(ax, data, options)
        else:
            # Default to scatter
            _plot_scatter(ax, data, options)

        if title:
            ax.set_title(title, fontweight="bold", pad=15)
        if x_label:
            ax.set_xlabel(x_label)
        if y_label:
            ax.set_ylabel(y_label)

        # Tight layout to avoid clipping
        fig.tight_layout()

        # Save to buffer
        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
        buf.seek(0)
        img_base64 = base64.b64encode(buf.read()).decode("utf-8")

        plt.close(fig)

        return {
            "success": True,
            "image_base64": img_base64,
            "mime_type": "image/png",
            "width": width * 150,  # pixels at 150 DPI
            "height": height * 150,
        }

    except Exception as e:
        plt.close("all")
        return {
            "success": False,
            "error": str(e),
            "image_base64": "",
        }


def _plot_scatter(ax, data, options):
    """Scatter plot."""
    if "series" in data:
        for series in data["series"]:
            ax.scatter(
                series.get("x", []),
                series.get("y", []),
                label=series.get("name", ""),
                alpha=options.get("alpha", 0.7),
                s=options.get("point_size", 50),
            )
        ax.legend()
    else:
        colors = options.get("color", "#2196F3")
        ax.scatter(
            data.get("x", []),
            data.get("y", []),
            c=colors,
            alpha=options.get("alpha", 0.7),
            s=options.get("point_size", 50),
        )

    if options.get("trend_line"):
        x = np.array(data.get("x", []), dtype=float)
        y = np.array(data.get("y", []), dtype=float)
        if len(x) > 1:
            z = np.polyfit(x, y, 1)
            p = np.poly1d(z)
            ax.plot(sorted(x), p(sorted(x)), "r--", alpha=0.8, label=f"Trend (R²={np.corrcoef(x,y)[0,1]**2:.3f})")
            ax.legend()


def _plot_bar(ax, data, options):
    """Bar chart."""
    labels = data.get("labels", [])
    if "series" in data:
        n_series = len(data["series"])
        bar_width = 0.8 / n_series
        x = np.arange(len(labels))
        for i, series in enumerate(data["series"]):
            offset = (i - n_series / 2 + 0.5) * bar_width
            ax.bar(x + offset, series.get("y", series.get("values", [])),
                   bar_width, label=series.get("name", ""), alpha=0.85)
        ax.set_xticks(x)
        ax.set_xticklabels(labels, rotation=45 if len(labels) > 6 else 0, ha="right")
        ax.legend()
    else:
        values = data.get("y", data.get("values", []))
        colors = options.get("colors", None)
        if options.get("horizontal"):
            ax.barh(labels, values, color=colors, alpha=0.85)
        else:
            ax.bar(labels, values, color=colors, alpha=0.85)
            if len(labels) > 6:
                plt.xticks(rotation=45, ha="right")


def _plot_line(ax, data, options):
    """Line chart."""
    if "series" in data:
        for series in data["series"]:
            ax.plot(
                series.get("x", range(len(series.get("y", [])))),
                series.get("y", []),
                marker=options.get("marker", "o"),
                label=series.get("name", ""),
                linewidth=options.get("linewidth", 2),
            )
        ax.legend()
    else:
        x = data.get("x", range(len(data.get("y", []))))
        ax.plot(
            x,
            data.get("y", []),
            marker=options.get("marker", "o"),
            color=options.get("color", "#2196F3"),
            linewidth=options.get("linewidth", 2),
        )


def _plot_histogram(ax, data, options):
    """Histogram."""
    values = data.get("values", data.get("x", []))
    bins = options.get("bins", "auto")
    ax.hist(values, bins=bins, color=options.get("color", "#2196F3"),
            alpha=0.75, edgecolor="white")
    if options.get("density"):
        ax.set_ylabel("Density")


def _plot_pie(ax, data, options):
    """Pie chart."""
    labels = data.get("labels", [])
    values = data.get("values", data.get("y", []))
    explode = options.get("explode", None)
    ax.pie(values, labels=labels, autopct="%1.1f%%", startangle=90,
           explode=explode)
    ax.axis("equal")


def _plot_box(ax, data, options):
    """Box plot."""
    if "series" in data:
        plot_data = [s.get("values", s.get("y", [])) for s in data["series"]]
        labels = [s.get("name", f"Series {i+1}") for i, s in enumerate(data["series"])]
        ax.boxplot(plot_data, labels=labels)
    else:
        values = data.get("values", data.get("y", []))
        ax.boxplot(values)
