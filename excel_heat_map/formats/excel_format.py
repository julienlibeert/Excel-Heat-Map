"""Excel number format parsing and application.

This module handles conversion of Excel number format strings to Python format
strings and applies them to numeric values for display on chart axes.
Supports sections for positive/negative/zero values, decimals, grouping,
percentages, scientific notation, and literal text.
"""

import re
from decimal import Decimal, ROUND_HALF_UP

import numpy as np
import pandas as pd

DEFAULT_NUMBER_FORMAT = "#,##0.00"
TIME_DISPLAY_FORMAT = "%Y-%m-%d %H:%M:%S"

# Regex to match Excel number format pattern (e.g., "#,##0.00")
_NUMBER_PATTERN = re.compile(
    r"[#0?][#0?,]*(?:\.[#0?]+)?(?:[Ee][+-]?[0#?]+)?"
)


def _split_format_sections(number_format: str) -> list[str]:
    """Split Excel format string into sections.

    Excel format strings can have up to 4 sections separated by semicolons:
    positive; negative; zero; text. Quoted sections are not split.

    Args:
        number_format (str): Excel format string.

    Returns:
        list[str]: List of format sections.
    """
    sections = []
    current = []
    in_quotes = False
    for character in number_format:
        if character == '"':
            in_quotes = not in_quotes
        if character == ";" and not in_quotes:
            sections.append("".join(current))
            current.clear()
        else:
            current.append(character)
    sections.append("".join(current))
    return sections


def _format_literal(text: str) -> str:
    """Extract literal text from an Excel format string section.

    Handles quoted text, escaped characters, and special Excel codes like
    underscores (space) and asterisks (fill).

    Args:
        text (str): Format section containing literals and codes.

    Returns:
        str: Cleaned literal text.
    """
    output = []
    in_quotes = False
    index = 0
    while index < len(text):
        character = text[index]
        if character == '"':
            in_quotes = not in_quotes
        elif character == "\\" and index + 1 < len(text):
            index += 1
            output.append(text[index])
        elif character in "_*" and index + 1 < len(text):
            index += 1
            if character == "_":
                output.append(" ")
        elif character == "[" and not in_quotes:
            close = text.find("]", index + 1)
            if close != -1:
                index = close
        elif in_quotes or not character.isalnum():
            output.append(character)
        index += 1
    return "".join(output)


def _has_unquoted_percent(text: str) -> bool:
    """Check if a format section contains an unquoted percent sign.

    Args:
        text (str): Format section.

    Returns:
        bool: True if % appears outside quotes.
    """
    in_quotes = False
    escaped = False
    for character in text:
        if escaped:
            escaped = False
        elif character == "\\":
            escaped = True
        elif character == '"':
            in_quotes = not in_quotes
        elif character == "%" and not in_quotes:
            return True
    return False


def format_excel_number(value: float, number_format: str) -> str:
    """Format a numeric value using an Excel format string.

    Supports section-based positive/negative/zero formats, decimals, grouping,
    percentages, scientific notation, quoted literals, and unit suffixes.

    Args:
        value (float): The number to format.
        number_format (str): Excel format string (e.g., "#,##0.00").

    Returns:
        str: Formatted number string. Returns "" for non-finite values.

    Examples:
        >>> format_excel_number(1234.5, "#,##0.00")
        '1,234.50'
        >>> format_excel_number(-500, "0.0;-0.0")
        '-500.0'
    """
    if not np.isfinite(value):
        return ""
    sections = _split_format_sections(number_format.strip())
    if value < 0 and len(sections) > 1:
        section = sections[1]
        number = abs(value)
        automatic_negative_sign = False
    elif value == 0 and len(sections) > 2:
        section = sections[2]
        number = value
        automatic_negative_sign = False
    else:
        section = sections[0]
        automatic_negative_sign = value < 0
        number = abs(value) if automatic_negative_sign else value

    match = _NUMBER_PATTERN.search(section)
    if match is None:
        return _format_literal(section)

    pattern = match.group()
    prefix = _format_literal(section[:match.start()])
    suffix = _format_literal(section[match.end():])
    percent = _has_unquoted_percent(section)
    number_value = Decimal(str(number))
    if percent:
        number_value *= 100

    mantissa = re.split("[Ee]", pattern, maxsplit=1)[0]
    decimal_pattern = mantissa.partition(".")[2]
    minimum_decimals = decimal_pattern.count("0")
    maximum_decimals = len(decimal_pattern)
    has_grouping = "," in mantissa.partition(".")[0]
    exponent = "E" in pattern.upper()
    if exponent:
        formatted = format(number_value, f".{maximum_decimals}E")
        exponent_digits = re.search(r"[Ee][+-]?([0#?]+)$", pattern)
        if exponent_digits is not None:
            mantissa_text, exponent_text = formatted.split("E")
            exponent_value = int(exponent_text)
            sign = "+" if exponent_value >= 0 else "-"
            formatted = (
                f"{mantissa_text}E{sign}"
                f"{abs(exponent_value):0{len(exponent_digits.group(1))}d}"
            )
    else:
        quantum = Decimal(1).scaleb(-maximum_decimals)
        rounded = number_value.quantize(quantum, rounding=ROUND_HALF_UP)
        format_spec = (
            f",.{maximum_decimals}f"
            if has_grouping
            else f".{maximum_decimals}f"
        )
        formatted = format(rounded, format_spec)
    if maximum_decimals > minimum_decimals:
        mantissa_text, exponent_text = (
            formatted.split("E", maxsplit=1) if exponent else (formatted, "")
        )
        whole, separator, fraction = mantissa_text.partition(".")
        fraction = fraction.rstrip("0")
        if len(fraction) < minimum_decimals:
            fraction += "0" * (minimum_decimals - len(fraction))
        formatted = whole + (separator + fraction if fraction else "")
        if exponent:
            formatted += "E" + exponent_text

    if automatic_negative_sign:
        prefix = "-" + prefix
    return prefix + formatted + suffix


def parse_time(value: str, label: str) -> pd.Timestamp:
    """Parse a date/time string to a pandas Timestamp.

    Args:
        value (str): Date/time string to parse.
        label (str): Display name for error messages.

    Returns:
        pd.Timestamp: Parsed timestamp (timezone-naive).

    Raises:
        ValueError: If the value cannot be parsed or is NaT.
    """
    try:
        timestamp = pd.Timestamp(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{label} must be a valid date/time.") from exc
    if pd.isna(timestamp):
        raise ValueError(f"{label} must be a valid date/time.")
    if timestamp.tz is not None:
        timestamp = timestamp.tz_convert(None)
    return timestamp


def to_naive_datetimes(values: pd.Series) -> pd.Series:
    """Convert a pandas Series to timezone-naive datetimes.

    Handles multiple input formats:
    - Excel serial dates (numbers between 20000 and 100000)
    - Pandas Timestamps
    - Python datetime objects
    - ISO format strings

    Args:
        values (pd.Series): Series containing date/time values.

    Returns:
        pd.Series: Series of datetime64[ns] values (timezone-naive).

    Raises:
        ValueError: If the series cannot be parsed as valid dates.
    """
    try:
        source = pd.Series(values, copy=False).reset_index(drop=True)
        is_excel_serial = source.map(
            lambda value: isinstance(value, (int, float, np.integer, np.floating))
            and not isinstance(value, (bool, np.bool_))
        )
        result = pd.Series(pd.NaT, index=source.index, dtype="datetime64[ns, UTC]")
        if is_excel_serial.any():
            result.loc[is_excel_serial] = pd.to_datetime(
                pd.to_numeric(source.loc[is_excel_serial], errors="coerce"),
                unit="D",
                origin="1899-12-30",
                errors="coerce",
                utc=True,
            )
        if (~is_excel_serial).any():
            result.loc[~is_excel_serial] = pd.to_datetime(
                source.loc[~is_excel_serial], errors="coerce", utc=True
            )
        return result.dt.tz_convert(None)
    except (TypeError, ValueError) as exc:
        raise ValueError("The selected time column cannot be parsed as dates.") from exc
