"""Data package initialization."""

from .excel_reader import (
    get_open_workbooks,
    select_workbook,
    get_sheet_by_name,
    get_tables,
    read_range_as_dataframe,
    read_table_as_dataframe,
)
from .data_loader import (
    normalize_dataframe_columns,
    check_duplicate_columns,
    is_datetime_like,
    get_numeric_columns,
    get_datetime_columns,
    validate_data_not_empty,
)

__all__ = [
    "get_open_workbooks",
    "select_workbook",
    "get_sheet_by_name",
    "get_tables",
    "read_range_as_dataframe",
    "read_table_as_dataframe",
    "normalize_dataframe_columns",
    "check_duplicate_columns",
    "is_datetime_like",
    "get_numeric_columns",
    "get_datetime_columns",
    "validate_data_not_empty",
]
