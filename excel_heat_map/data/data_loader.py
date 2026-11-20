"""Data loading and preparation for plotting.

Loads data from Excel and formats it for use in heat maps and trend graphs.
Handles column type detection, data validation, and frame normalization.
"""

import numpy as np
import pandas as pd
from datetime import date, datetime

from ..utils.errors import DataError


def normalize_dataframe_columns(df):
    """Normalize DataFrame column names to strings.

    Args:
        df (pd.DataFrame): The DataFrame to normalize.

    Returns:
        pd.DataFrame: DataFrame with string column names.
    """
    df = df.copy()
    df.columns = [str(col) for col in df.columns]
    return df


def check_duplicate_columns(df):
    """Check for duplicate column names in a DataFrame.

    Args:
        df (pd.DataFrame): The DataFrame to check.

    Raises:
        DataError: If duplicate column names are found.
    """
    column_list = list(df.columns)
    if len(set(column_list)) != len(column_list):
        raise DataError("The selected source has duplicate column names.")


def is_datetime_like(series):
    """Check if a pandas Series contains datetime-like data.

    Returns True for:
    - datetime64 dtype
    - Object dtype with Timestamp/datetime/date values
    - Excel serial date numbers (detected by range check)

    Args:
        series (pd.Series): The series to check.

    Returns:
        bool: True if the series appears to contain dates.
    """
    if pd.api.types.is_datetime64_any_dtype(series):
        return True
    if series.dtype == object:
        sample = series.dropna().head(10)
        if any(isinstance(value, (pd.Timestamp, np.datetime64, datetime, date))
               for value in sample):
            return True
        # Check for Excel serial dates
        try:
            values = pd.to_numeric(series, errors="coerce").to_numpy(dtype=float)
            finite = values[np.isfinite(values)]
            if finite.size > 0 and np.all((finite >= 20000) & (finite <= 100000)):
                return True
        except (TypeError, ValueError):
            pass
    return False


def get_numeric_columns(df):
    """Get list of numeric (non-datetime) columns from a DataFrame.

    Args:
        df (pd.DataFrame): The DataFrame.

    Returns:
        list[str]: Column names of numeric data.
    """
    numeric = []
    for column in df.columns:
        series = df[column]
        # Skip datetime columns
        if is_datetime_like(series):
            continue
        # Try to convert to numeric
        values = pd.to_numeric(series, errors="coerce")
        if values.notna().any():
            numeric.append(column)
    return numeric


def get_datetime_columns(df):
    """Get list of datetime-like columns from a DataFrame.

    Args:
        df (pd.DataFrame): The DataFrame.

    Returns:
        list[str]: Column names of datetime data.
    """
    datetime_cols = []
    for column in df.columns:
        if is_datetime_like(df[column]):
            datetime_cols.append(column)
    return datetime_cols


def validate_data_not_empty(df):
    """Check that a DataFrame contains data rows.

    Args:
        df (pd.DataFrame): The DataFrame to check.

    Raises:
        DataError: If DataFrame is None or empty.
    """
    if df is None or df.empty:
        raise DataError("The selected source contains no data rows.")
