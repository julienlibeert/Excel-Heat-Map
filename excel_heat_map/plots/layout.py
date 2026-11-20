"""Matplotlib layout utilities for multi-axis plots.

Provides functions to calculate appropriate spacing and positioning for
multiple y-axes in Matplotlib figures, ensuring labels and tick values don't
overlap.
"""

import matplotlib
import numpy as np
from matplotlib.backends.backend_agg import FigureCanvasAgg
from matplotlib.transforms import offset_copy


def layout_y_axes(figure, axes, axis_ranks, fixed_size):
    """Position stacked y-axes so labels clear all adjacent tick labels.

    Measures the width of each axis's numeric labels (using its formatter)
    and axis label, then positions each extra axis just beyond the previous
    axis's label. All axes at the same rank stay aligned across stacked
    charts. If a fixed figure size is too narrow, spacing is scaled down.

    Args:
        figure (Figure): The Matplotlib figure.
        axes (dict): Mapping of axis_key -> Matplotlib Axes for each axis.
        axis_ranks (dict): Mapping of axis_key -> rank (0=primary, 1+=secondary).
        fixed_size (tuple or None): (width_inches, height_inches) if size is fixed.

    Returns:
        tuple: (left_margin_inches, right_margin_inches) for subplot adjustment.

    Note:
        This function modifies the axes' spine positions and label coordinates
        in place.
    """
    renderer = FigureCanvasAgg(figure).get_renderer()
    pixels_to_points = 72 / figure.dpi
    tick_space = (
        matplotlib.rcParams["ytick.major.size"]
        + matplotlib.rcParams["ytick.major.pad"]
    )
    label_pad = 4.0  # Points between tick label and axis label
    axis_gap = 8.0  # Points between axes
    edge_pad = 6.0  # Points before figure edge
    empty_margin = 0.2 * 72  # Points for side with no axes

    # Measure tick label and axis label widths for each (side, rank) slot
    slot_sizes = {}
    for axis_key, axis in axes.items():
        lower, upper = sorted(axis.get_ylim())
        locations = [
            location
            for location in axis.yaxis.get_majorticklocs()
            if lower <= location <= upper
        ]
        texts = axis.yaxis.get_major_formatter().format_ticks(locations)
        font = axis.yaxis.get_major_ticks()[0].label1.get_fontproperties()
        tick_width = max(
            (
                renderer.get_text_width_height_descent(text, font, ismath=False)[0]
                for text in texts
                if text
            ),
            default=0.0,
        ) * pixels_to_points
        label_width = (
            axis.yaxis.label.get_window_extent(renderer).width * pixels_to_points
        )
        slot = (axis_key[1], axis_ranks[axis_key])
        previous = slot_sizes.get(slot, (0.0, 0.0))
        slot_sizes[slot] = (
            max(previous[0], tick_width),
            max(previous[1], label_width),
        )

    # Calculate spine and label offsets for each slot
    spine_offsets = {}
    label_offsets = {}
    margins = {}
    for side in ("Left", "Right"):
        ranks = sorted(rank for slot_side, rank in slot_sizes if slot_side == side)
        offset = 0.0
        outer_edge = 0.0
        for rank in ranks:
            tick_width, label_width = slot_sizes[(side, rank)]
            spine_offsets[(side, rank)] = offset
            label_offsets[(side, rank)] = offset + tick_space + tick_width + label_pad
            outer_edge = label_offsets[(side, rank)] + label_width
            offset = outer_edge + axis_gap
        margins[side] = outer_edge + edge_pad if ranks else empty_margin

    # If fixed size, scale down if margins don't fit
    scale = 1.0
    if fixed_size is not None:
        minimum_plot_width = min(1.0, fixed_size[0] * 0.1)
        maximum_margin = (fixed_size[0] - minimum_plot_width) * 72
        natural_margin = margins["Left"] + margins["Right"]
        if natural_margin > maximum_margin:
            scale = maximum_margin / natural_margin

    # Apply positions to axes
    for axis_key, axis in axes.items():
        slot = (axis_key[1], axis_ranks[axis_key])
        side = axis_key[1].lower()
        direction = -1 if side == "left" else 1
        axis.spines[side].set_position(
            ("outward", spine_offsets[slot] * scale)
        )
        axis.yaxis.set_label_coords(
            0 if side == "left" else 1,
            0.5,
            transform=offset_copy(
                axis.transAxes,
                fig=figure,
                x=direction * label_offsets[slot] * scale,
                y=0,
                units="points",
            ),
        )
    return margins["Left"] * scale / 72, margins["Right"] * scale / 72
