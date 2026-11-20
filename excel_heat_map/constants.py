"""Shared constants for the Excel Heat Map application.

This module contains all application-wide constants used by both the heat map
and trend navigator features, including configuration values, field names,
and display settings.
"""

# Version and application metadata
VERSION_ID = "1.0a (last updated 20/11/2026)"
WINDOW_TITLE = "Julien's Excel MatplotLib Plotter"
DEFAULT_WINDOW_GEOMETRY = "1100x900"

# Graph export settings
GRAPH_EXPORT_DPI = 300

# Excel sheet and field names for settings persistence
SETTINGS_SHEET_NAME = "HeatMapSettings"
TREND_SETTINGS_SHEET_NAME = "TrendNavigatorSettings"
SCATTERPLOT_SETTINGS_SHEET_NAME = "ScatterplotSettings"

# Heat map settings field names (order matters for Excel persistence)
HEATMAP_SETTINGS_FIELDS = [
    "Settings name",
    "Source sheet",
    "Excel table",
    "Source range",
    "X column",
    "Y column",
    "Weight column",
    "X bin width",
    "Y bin height",
    "Aggregation",
    "Graph title",
    "X-axis title",
    "Y-axis title",
    "Weights colormap",
    "Reverse colormap",
    "X-axis minimum",
    "X-axis maximum",
    "Y-axis minimum",
    "Y-axis maximum",
    "Legend minimum",
    "Legend maximum",
    "Below-range color",
    "Above-range color",
]

# Trend navigator settings field names (order matters for Excel persistence)
TREND_SETTINGS_FIELDS = [
    "Settings name",
    "Source sheet",
    "Excel table",
    "Source range",
    "Trend settings JSON",
]

# Scatterplot settings field names (order matters for Excel persistence)
SCATTERPLOT_SETTINGS_FIELDS = [
    "Settings name",
    "Source sheet",
    "Excel table",
    "Source range",
    "Scatterplot settings JSON",
]

# Alias for backward compatibility with existing code
TREND_HEATMAP_SETTINGS_FIELDS = TREND_SETTINGS_FIELDS

# Trend navigator limits
MAX_TRENDS = 16

# Color map defaults
DEFAULT_COLORMAP = "viridis"
DEFAULT_UNDER_COLOR = "#440154"
DEFAULT_OVER_COLOR = "#fde725"

# Heatmap aggregation options
AGGREGATION_OPTIONS = [
    "Bin count",
    "Sum",
    "Mean",
    "Median",
    "Min",
    "Max",
    "Std Dev",
]

# Default bin values
DEFAULT_X_BIN = "1"
DEFAULT_Y_BIN = "1"
