"""Excel data access layer.

Handles opening Excel workbooks and reading data from sheets, tables, and ranges.
Provides abstraction over xlwings for workbook/sheet/table operations.
"""

try:
    import xlwings as xw
    HAS_XLWINGS = True
except ImportError:
    xw = None
    HAS_XLWINGS = False

from ..utils.errors import DataError


def get_open_workbooks():
    """Get all currently open Excel workbooks.

    Returns:
        dict: Mapping of workbook display labels to (app_pid, book_name) tuples.
              Label format: "workbook.xlsx (Excel PID 12345)".
    
    Raises:
        ImportError: If xlwings is not installed.
    """
    if not HAS_XLWINGS:
        raise ImportError("Excel integration requires xlwings to be installed.")
    
    entries = {}
    for app in xw.apps:
        for book in app.books:
            label = f"{book.name} (Excel PID {app.pid})"
            entries[label] = (app.pid, book.name)
    return entries


def select_workbook(label, entries):
    """Select and return an Excel workbook by label.

    Args:
        label (str): Workbook label (e.g., "workbook.xlsx (Excel PID 12345)").
        entries (dict): Mapping from get_open_workbooks().

    Returns:
        xlwings.Book: The selected workbook.

    Raises:
        DataError: If the label is not in the entries dict.
        ImportError: If xlwings is not installed.
    """
    if not HAS_XLWINGS:
        raise ImportError("Excel integration requires xlwings to be installed.")
    
    if label not in entries:
        raise DataError(
            "Select an open workbook, then refresh the workbook list if needed."
        )
    pid, name = entries[label]
    return xw.apps[pid].books[name]


def get_sheet_by_name(book, sheet_name):
    """Get a sheet from a workbook by name.

    Args:
        book (xlwings.Book): The workbook.
        sheet_name (str): Name of the sheet.

    Returns:
        xlwings.Sheet: The sheet, or None if not found.
    """
    for sheet in book.sheets:
        if sheet.name == sheet_name:
            return sheet
    return None


def get_tables(book):
    """Get all table names in the first sheet of a workbook.

    Args:
        book (xlwings.Book): The workbook.

    Returns:
        list[str]: List of table names.
    """
    if not book.sheets:
        return []
    sheet = book.sheets[0]
    try:
        tables = sheet.tables
        return [table.name for table in tables] if tables else []
    except Exception:
        return []


def read_range_as_dataframe(sheet, range_address):
    """Read a named range or address from a sheet as a pandas DataFrame.

    Args:
        sheet (xlwings.Sheet): The sheet.
        range_address (str): Range address (e.g., "A1:C100") or named range.

    Returns:
        pd.DataFrame: Data from the range.

    Raises:
        DataError: If the range cannot be read.
    """
    try:
        data = sheet.range(range_address).options(ndim=2, expand='table').value
        if data is None:
            raise DataError(f"The range '{range_address}' is empty.")
        import pandas as pd
        df = pd.DataFrame(data[1:], columns=data[0]) if isinstance(data, list) else pd.DataFrame(data)
        return df
    except Exception as exc:
        raise DataError(f"Cannot read range '{range_address}': {str(exc)}") from exc


def read_table_as_dataframe(sheet, table_name):
    """Read an Excel table as a pandas DataFrame.

    Args:
        sheet (xlwings.Sheet): The sheet containing the table.
        table_name (str): Name of the table.

    Returns:
        pd.DataFrame: Data from the table.

    Raises:
        DataError: If the table cannot be read.
    """
    try:
        table = sheet.tables[table_name]
        data = table.range.options(ndim=2, expand='table').value
        if data is None:
            raise DataError(f"The table '{table_name}' is empty.")
        import pandas as pd
        return pd.DataFrame(data[1:], columns=data[0]) if isinstance(data, list) else pd.DataFrame(data)
    except Exception as exc:
        raise DataError(f"Cannot read table '{table_name}': {str(exc)}") from exc
