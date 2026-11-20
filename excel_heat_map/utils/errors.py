"""Custom exceptions for the Excel Heat Map application."""


class ExcelHeatMapError(Exception):
    """Base exception for all Excel Heat Map errors."""

    pass


class ValidationError(ExcelHeatMapError):
    """Raised when input validation fails."""

    pass


class DataError(ExcelHeatMapError):
    """Raised when data loading or processing fails."""

    pass


class LayoutError(ExcelHeatMapError):
    """Raised when plot layout computation fails."""

    pass
