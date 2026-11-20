"""Input validation functions for the Excel Heat Map application."""

import numpy as np
from .errors import ValidationError


def parse_optional_limit(text, label):
    """Parse an optional numeric limit from text input.

    Args:
        text (str): The text to parse. If empty/whitespace, returns None.
        label (str): Display name for error messages (e.g., "X minimum").

    Returns:
        float or None: Parsed numeric limit, or None if input is empty.

    Raises:
        ValidationError: If the input cannot be parsed as a finite number.
    """
    text = text.strip()
    if not text:
        return None
    try:
        value = float(text)
    except ValueError as exc:
        raise ValidationError(f"{label} must be a finite number or left blank.") from exc
    if not np.isfinite(value):
        raise ValidationError(f"{label} must be a finite number or left blank.")
    return value


def validate_limits(lower, upper, label):
    """Ensure lower < upper for axis limit pairs.

    Args:
        lower (float or None): The lower limit.
        upper (float or None): The upper limit.
        label (str): Display name for error messages (e.g., "X-axis").

    Raises:
        ValidationError: If both limits are set and lower >= upper.
    """
    if lower is not None and upper is not None and lower >= upper:
        raise ValidationError(f"{label} minimum must be less than maximum.")
