"""Heat map binning and aggregation utilities.

Handles creation of bin edges for heat map axes and computation of bin statistics.
"""

import numpy as np

from ..utils.errors import ValidationError


def make_bin_edges(specification, values, lower, upper, axis_label):
    """Create bin edges for a heat map axis from a bin width specification.

    Given a bin width, generates evenly-spaced bin edges that span from the
    minimum to maximum of the data (or specified limits). Ensures at least one
    bin and handles edge cases like single-value data.

    Args:
        specification (str): Bin width as a string number (e.g., "10").
        values (array-like): Data values to bin.
        lower (float or None): Optional lower limit (overrides data minimum).
        upper (float or None): Optional upper limit (overrides data maximum).
        axis_label (str): Display name for error messages (e.g., "X-axis").

    Returns:
        tuple: (bin_edges, display_min, display_max) where:
            - bin_edges (np.ndarray): 1D array of bin edges
            - display_min (float): Minimum value used for display range
            - display_max (float): Maximum value used for display range

    Raises:
        ValidationError: If specification is invalid or empty, or if bin width
                        is non-positive.
    """
    text = specification.strip()
    if not text:
        raise ValidationError(f"Enter a bin width for {axis_label}.")
    try:
        width = float(text)
    except ValueError as exc:
        raise ValidationError(
            f"{axis_label} bin width must be a finite positive number."
        ) from exc
    if not np.isfinite(width) or width <= 0:
        raise ValidationError(
            f"{axis_label} bin width must be a finite positive number."
        )

    data_min = float(np.min(values))
    data_max = float(np.max(values))
    display_min = data_min if lower is None else lower
    display_max = data_max if upper is None else upper

    if display_min == display_max:
        # Handle single value: create a bin containing it
        display_min -= width / 2
        display_max += width / 2

    # Create bin edges starting from floor of display_min
    start = np.floor(display_min / width) * width
    stop = np.ceil(display_max / width) * width + width / 2
    bin_edges = np.arange(start, stop, width)
    
    return bin_edges, display_min, display_max
