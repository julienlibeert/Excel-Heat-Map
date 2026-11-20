"""2D and 3D scatterplot utilities.

This module provides functions to create 2D and 3D scatterplots with optional
color mapping, supporting both continuous and categorical color dimensions.
"""

import numpy as np
import pandas as pd
import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib import ticker
from matplotlib.cm import ScalarMappable
from matplotlib.colors import BoundaryNorm, Colormap, Normalize
from numpy.typing import ArrayLike
from typing import TypedDict


class _ScatterStyle(TypedDict, total=False):
    s: float
    marker: str
    alpha: float
    c: ArrayLike
    cmap: str | Colormap
    norm: Normalize


def _is_nothing(element):
    """Check whether element is NaN or None."""
    try:
        return np.isnan(element)
    except (TypeError, ValueError):
        return element is None


def _get_axis_formatter(format_type, decimal_places):
    """Get matplotlib axis formatter based on format type.
    
    Args:
        format_type (str): One of "FLOAT", "PERCENTAGE", "PPM", "SCIENTIFIC"
        decimal_places (int): Number of decimal places to display
    
    Returns:
        ticker.Formatter: Matplotlib formatter for axis labels
    """
    format_type = str(format_type).upper() if format_type else "FLOAT"
    decimal_places = int(decimal_places) if decimal_places else 0
    
    match format_type:
        case "FLOAT":
            return ticker.StrMethodFormatter("{x:,." + str(decimal_places) + "f}")
        case "PERCENTAGE":
            return ticker.StrMethodFormatter("{x:." + str(decimal_places) + "%}")
        case "PPM":
            return ticker.FuncFormatter(
                lambda x, pos: f"{x*1e6:,.{decimal_places}f} ppm"
            )
        case "SCIENTIFIC":
            return ticker.StrMethodFormatter("{x:,." + str(decimal_places) + "e}")
        case _:
            return ticker.StrMethodFormatter("{x:,.0f}")


def _draw_scatter(ax, df, axes, color_series, colormap_name, marker_type,
                  marker_size, color_label, color_min, color_max, marker_alpha):
    frame = df.copy()
    for column in axes:
        frame[column] = pd.to_numeric(frame[column], errors="coerce")
    valid = np.isfinite(frame[list(axes)].to_numpy(dtype=float)).all(axis=1)
    options: _ScatterStyle = {"s": marker_size, "marker": marker_type, "alpha": marker_alpha}
    categories = None
    if color_series is not None:
        numeric = pd.to_numeric(df[color_series], errors="coerce")
        if numeric.notna().any():
            valid &= np.isfinite(numeric.to_numpy(dtype=float))
            options.update(c=numeric[valid], cmap=colormap_name,
                           norm=Normalize(vmin=color_min, vmax=color_max))
        else:
            categorical = pd.Categorical(df[color_series])
            categories = categorical.categories
            valid &= categorical.codes >= 0
            count = len(categories)
            if count:
                options.update(c=categorical.codes[valid],
                               cmap=mpl.colormaps[colormap_name].resampled(count),
                               norm=BoundaryNorm(
                                   np.arange(count + 1) - 0.5, count))
    if not valid.any():
        plt.close(ax.figure)
        raise ValueError("The selected columns contain no usable scatterplot observations.")
    points = ax.scatter(*(frame.loc[valid, column] for column in axes), **options)
    if color_series is not None:
        # Share the color scale, not the markers' opacity.
        bar = ax.figure.colorbar(
            ScalarMappable(norm=points.norm, cmap=points.cmap),
            ax=ax, label=color_label, alpha=1,
        )
        if categories is not None:
            bar.set_ticks(range(len(categories)), labels=list(categories))


def create_2d_scatter(
    df,
    x_series,
    y_series,
    color_series=None,
    title="",
    colormap_name="gist_rainbow",
    marker_type="o",
    marker_size=None,
    figure_size=None,
    x_label=None,
    x_min=None,
    x_max=None,
    x_format_type=None,
    x_decimal_places=0,
    y_label=None,
    y_min=None,
    y_max=None,
    y_format_type=None,
    y_decimal_places=0,
    color_label=None,
    color_min=None,
    color_max=None,
    dpi=None,
    marker_alpha: float = 0.5,
):
    """Create a 2D scatterplot with optional color mapping.
    
    Args:
        df (pd.DataFrame): Data frame containing the data
        x_series (str): Column name for X axis
        y_series (str): Column name for Y axis
        color_series (str, optional): Column name for color mapping
        title (str): Graph title
        colormap_name (str): Matplotlib colormap name
        marker_type (str): Marker style ('o', 's', '^', etc.)
        marker_size (float, optional): Marker size (default 15)
        figure_size (tuple, optional): (width, height) in inches
        x_label (str, optional): X axis label (defaults to x_series)
        x_min (float, optional): X axis minimum
        x_max (float, optional): X axis maximum
        x_format_type (str, optional): Number format type
        x_decimal_places (int): X axis decimal places
        y_label (str, optional): Y axis label (defaults to y_series)
        y_min (float, optional): Y axis minimum
        y_max (float, optional): Y axis maximum
        y_format_type (str, optional): Number format type
        y_decimal_places (int): Y axis decimal places
        color_label (str, optional): Color bar label
        color_min (float, optional): Color minimum value
        color_max (float, optional): Color maximum value
        dpi (int, optional): Figure DPI
        marker_alpha (float): Marker opacity from 0 (invisible) to 1 (opaque)
    
    Returns:
        tuple: (figure, axes) matplotlib objects
    """
    # Set default labels
    x_label = x_label or x_series
    y_label = y_label or y_series
    color_label = color_label or color_series
    marker_size = marker_size or 15
    
    fig, ax = plt.subplots(1, 1)
    
    # Set figure size and DPI if specified
    if figure_size:
        fig.set_size_inches(*figure_size)
    if dpi:
        fig.set_dpi(dpi)
    
    _draw_scatter(ax, df, (x_series, y_series), color_series, colormap_name,
                  marker_type, marker_size, color_label, color_min, color_max, marker_alpha)
    
    # Format and configure axes
    ax.set_xlabel(x_label)
    ax.set_ylabel(y_label)
    
    ax.xaxis.set_major_formatter(_get_axis_formatter(x_format_type, x_decimal_places))
    ax.yaxis.set_major_formatter(_get_axis_formatter(y_format_type, y_decimal_places))
    
    if not _is_nothing(x_min):
        ax.set_xbound(lower=x_min)
    if not _is_nothing(x_max):
        ax.set_xbound(upper=x_max)
    if not _is_nothing(y_min):
        ax.set_ybound(lower=y_min)
    if not _is_nothing(y_max):
        ax.set_ybound(upper=y_max)
    
    ax.set_title(title)
    
    return fig, ax


def create_3d_scatter(
    df,
    x_series,
    y_series,
    z_series,
    color_series=None,
    title="",
    colormap_name="gist_rainbow",
    marker_type="o",
    marker_size=None,
    figure_size=None,
    x_label=None,
    x_min=None,
    x_max=None,
    x_format_type=None,
    x_decimal_places=0,
    y_label=None,
    y_min=None,
    y_max=None,
    y_format_type=None,
    y_decimal_places=0,
    z_label=None,
    z_min=None,
    z_max=None,
    z_format_type=None,
    z_decimal_places=0,
    color_label=None,
    color_min=None,
    color_max=None,
    elevation: float = 30,
    azimuth: float = 45,
    focal_length="orthographic",
    x_scale=1,
    y_scale=1,
    z_scale=1,
    dpi=None,
    marker_alpha: float = 0.5,
):
    """Create a 3D scatterplot with optional color mapping.
    
    Args:
        df (pd.DataFrame): Data frame containing the data
        x_series (str): Column name for X axis
        y_series (str): Column name for Y axis
        z_series (str): Column name for Z axis
        color_series (str, optional): Column name for color mapping
        title (str): Graph title
        colormap_name (str): Matplotlib colormap name
        marker_type (str): Marker style
        marker_size (float, optional): Marker size
        figure_size (tuple, optional): (width, height) in inches
        x_label, x_min, x_max, x_format_type, x_decimal_places: X axis config
        y_label, y_min, y_max, y_format_type, y_decimal_places: Y axis config
        z_label, z_min, z_max, z_format_type, z_decimal_places: Z axis config
        color_label (str, optional): Color bar label
        color_min, color_max: Color scale limits
        elevation (float): Camera elevation angle
        azimuth (float): Camera azimuth angle
        focal_length (str or float): "orthographic" or focal length value
        x_scale, y_scale, z_scale (float): Axis scaling factors
        dpi (int, optional): Figure DPI
        marker_alpha (float): Marker opacity from 0 (invisible) to 1 (opaque)
    
    Returns:
        tuple: (figure, axes) matplotlib objects
    """
    # Set default labels
    x_label = x_label or x_series
    y_label = y_label or y_series
    z_label = z_label or z_series
    color_label = color_label or color_series
    marker_size = marker_size or 15
    
    fig = plt.figure()
    ax = fig.add_subplot(111, projection="3d")
    
    # Set figure size and DPI if specified
    if figure_size:
        fig.set_size_inches(*figure_size)
    if dpi:
        fig.set_dpi(dpi)
    
    _draw_scatter(ax, df, (x_series, y_series, z_series), color_series, colormap_name,
                  marker_type, marker_size, color_label, color_min, color_max, marker_alpha)
    
    # Set labels
    ax.set_xlabel(x_label)
    ax.set_ylabel(y_label)
    ax.set_zlabel(z_label)
    ax.xaxis.set_major_formatter(_get_axis_formatter(x_format_type, x_decimal_places))
    ax.yaxis.set_major_formatter(_get_axis_formatter(y_format_type, y_decimal_places))
    ax.zaxis.set_major_formatter(_get_axis_formatter(z_format_type, z_decimal_places))
    
    # Set camera view
    ax.view_init(elev=elevation, azim=azimuth)
    
    # Set projection type
    if focal_length == "orthographic":
        ax.set_proj_type("ortho")
    else:
        try:
            ax.set_proj_type("persp", focal_length=float(focal_length))
        except (TypeError, ValueError):
            ax.set_proj_type("persp", focal_length=1.0)
    
    # Set axis limits
    if not _is_nothing(x_min):
        ax.set_xlim(left=x_min)
    if not _is_nothing(x_max):
        ax.set_xlim(right=x_max)
    if not _is_nothing(y_min):
        ax.set_ylim(bottom=y_min)
    if not _is_nothing(y_max):
        ax.set_ylim(top=y_max)
    if not _is_nothing(z_min):
        ax.set_zlim(bottom=z_min)
    if not _is_nothing(z_max):
        ax.set_zlim(top=z_max)
    
    ax.set_title(title)
    ax.set_box_aspect(aspect=(x_scale, y_scale, z_scale))
    
    return fig, ax
