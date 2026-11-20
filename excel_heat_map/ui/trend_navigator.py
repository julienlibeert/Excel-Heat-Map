"""Time-based multi-series graph controls and plotting for the Excel heat-map app.

This module provides the TrendNavigator UI component, which allows users to select
data columns, configure trend display properties (colors, axes, ranges), and generate
interactive time-series plots with support for multiple y-axes and hairline overlays.
"""

from __future__ import annotations

import re
import tkinter as tk
from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import date, datetime
from tkinter import colorchooser, ttk

import matplotlib
import matplotlib.dates as mdates
import numpy as np
import pandas as pd
from matplotlib.backends.backend_agg import FigureCanvasAgg
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.colors import is_color_like, to_hex
from matplotlib.figure import Figure
from matplotlib.offsetbox import AnnotationBbox, TextArea, VPacker
from matplotlib.ticker import FuncFormatter
from matplotlib.transforms import offset_copy

# Import from reorganized package modules
from ..constants import MAX_TRENDS
from ..formats import format_excel_number, parse_time, to_naive_datetimes
from ..formats import DEFAULT_NUMBER_FORMAT, TIME_DISPLAY_FORMAT
from ..plots import layout_y_axes
from .plot_controls import colormap_preview, protect_combobox_wheel


@dataclass
class TrendRow:
    row_id: int
    widgets: list[tk.Widget]
    visible: tk.BooleanVar
    column: tk.StringVar
    label: tk.StringVar
    chart_index: tk.StringVar
    side: tk.StringVar
    minimum: tk.StringVar
    maximum: tk.StringVar
    color: tk.StringVar
    unit: tk.StringVar
    number_format: tk.StringVar
    default_label: str
    column_combo: ttk.Combobox
    color_button: tk.Button
    color_override: bool = False


@dataclass
class Hairline:
    hairline_id: int
    frame: ttk.Frame
    time: tk.StringVar
    color: str
    artists: list = field(default_factory=list)
    readout_artists: list = field(default_factory=list)


class TrendNavigator:
    """Interactive trend graph UI embedded in one graph-type tab."""

    def __init__(
        self,
        parent: ttk.Frame,
        data_provider: Callable[[], pd.DataFrame],
        save_figure: Callable[[Figure | None, str], None],
        copy_figure: Callable[[Figure | None], None],
        set_status: Callable[..., None],
        figure_size_inches: Callable[[], tuple[float, float] | None] | None = None,
        save_settings: Callable[[str, dict], None] | None = None,
        refresh_settings: Callable[[], object] | None = None,
        load_settings: Callable[[str], None] | None = None,
    ) -> None:
        self.parent = parent
        self.data_provider = data_provider
        self.save_figure_callback = save_figure
        self.copy_figure_callback = copy_figure
        self.set_status = set_status
        self.figure_size_inches = figure_size_inches or (lambda: None)
        self.save_settings_callback = save_settings
        self.refresh_settings_callback = refresh_settings
        self.load_settings_callback = load_settings
        self.columns = []
        self._source_frame = None
        self.trend_rows: list[TrendRow] = []
        self.next_trend_id = 1
        self.hairlines: list[Hairline] = []
        self.next_hairline_id = 1
        self.figure = None
        self.canvas = None
        self.figure_fixed_size = False
        self.base_axes = None
        self.chart_axes = {}
        self.axes_by_group = {}
        self.plotted_series = []
        self.data_time_range = None
        self._pan_start = None
        self._drag_hairline = None

        self.time_column = tk.StringVar()
        self.title = tk.StringVar(value="Trend navigator")
        self.start_time = tk.StringVar()
        self.end_time = tk.StringVar()
        self.color_mode = tk.StringVar(value="Individual colors")
        self.colormap = tk.StringVar(value="viridis")
        self.concise_hairline_info = tk.BooleanVar(value=False)
        self.new_hairline_time = tk.StringVar()
        self.settings_name = tk.StringVar()
        self.saved_settings_name = tk.StringVar()

        self._build_controls()
        self.add_trend()
        self._applying_settings = False
        self.color_mode.trace_add("write", self._update_colormap_colors)
        self.colormap.trace_add("write", self._update_colormap_colors)

    def _build_controls(self):
        saved_settings = self._section("Saved trend settings for this workbook")
        saved_settings.pack(fill="x", pady=(0, 4))
        settings_row = ttk.Frame(saved_settings)
        settings_row.pack(fill="x", pady=2)
        ttk.Label(settings_row, text="Settings name:", width=16).pack(side="left")
        ttk.Entry(settings_row, textvariable=self.settings_name, width=24).pack(
            side="left", padx=(0, 12)
        )
        ttk.Label(settings_row, text="Saved settings:", width=14).pack(side="left")
        self.saved_settings_combo = ttk.Combobox(
            settings_row,
            textvariable=self.saved_settings_name,
            state="readonly",
            width=28,
        )
        self.saved_settings_combo.pack(side="left", padx=(0, 8))
        self._protect_combobox_wheel(self.saved_settings_combo)
        settings_buttons = ttk.Frame(saved_settings)
        settings_buttons.pack(anchor="w", pady=(2, 0))
        ttk.Button(
            settings_buttons, text="Save settings", command=self.save_settings
        ).pack(side="left", padx=(0, 5))
        ttk.Button(
            settings_buttons, text="Refresh saved settings", command=self.refresh_settings
        ).pack(side="left", padx=(0, 5))
        ttk.Button(
            settings_buttons, text="Load settings", command=self.load_settings
        ).pack(side="left")

        settings_columns = ttk.Frame(self.parent)
        settings_columns.pack(fill="x")
        settings_columns.columnconfigure((0, 1), weight=1, uniform="trend-settings")

        graph_settings = self._section("Graph and time range", settings_columns)
        graph_settings.grid(row=0, column=0, sticky="nsew", padx=(0, 4), pady=3)
        self._entry_row("Graph title:", self.title, parent=graph_settings, label_width=15)
        self._combo_row(
            "Time column:", self.time_column, parent=graph_settings, label_width=15
        )
        range_row = ttk.Frame(graph_settings)
        range_row.pack(fill="x", pady=2)
        ttk.Label(range_row, text="Start:", width=7).pack(side="left")
        ttk.Entry(range_row, textvariable=self.start_time, width=19).pack(
            side="left", padx=(0, 6)
        )
        ttk.Label(range_row, text="End:", width=5).pack(side="left")
        ttk.Entry(range_row, textvariable=self.end_time, width=19).pack(
            side="left", padx=(0, 4)
        )
        ttk.Label(range_row, text="YYYY-MM-DD HH:MM:SS").pack(side="left")
        range_buttons = ttk.Frame(graph_settings)
        range_buttons.pack(fill="x", pady=(2, 0))
        ttk.Button(
            range_buttons, text="Apply range", command=self.apply_time_range
        ).pack(side="left", padx=(0, 5))
        ttk.Button(
            range_buttons, text="Auto range", command=self.reset_time_range
        ).pack(side="left")
        ttk.Label(
            graph_settings,
            text="Scroll down to zoom at the pointer; drag to pan. Double-click to add a hairline; right-click one to delete it.",
            foreground="#666",
            wraplength=480,
        ).pack(anchor="w", pady=(4, 0))

        color_settings = self._section("Trend colors", settings_columns)
        color_settings.grid(row=0, column=1, sticky="nsew", padx=(4, 0), pady=3)
        color_row = ttk.Frame(color_settings)
        color_row.pack(fill="x", pady=2)
        ttk.Label(color_row, text="Color selection:", width=18).pack(side="left")
        color_mode_combo = ttk.Combobox(
            color_row,
            textvariable=self.color_mode,
            values=["Individual colors", "Colormap"],
            state="readonly",
            width=22,
        )
        color_mode_combo.pack(side="left")
        self._protect_combobox_wheel(color_mode_combo)
        ttk.Label(color_row, text="Colormap:", padding=(8, 0)).pack(side="left")
        colormap_combo = ttk.Combobox(
            color_row,
            textvariable=self.colormap,
            values=sorted(matplotlib.colormaps),
            state="readonly",
            width=20,
        )
        colormap_combo.pack(side="left")
        self._protect_combobox_wheel(colormap_combo)
        self.cmap_preview = colormap_preview(color_settings, self.colormap)

        trends_settings = self._section("Trends (up to 16)")
        trends_settings.pack(fill="x", pady=4)
        ttk.Label(
            trends_settings,
            text="Different chart numbers create vertically stacked plots; use the same number to overlay trends. Y axes group only when trends share a chart, side, and explicit minimum and maximum.",
            foreground="#666",
            wraplength=1000,
        ).pack(anchor="w", pady=(0, 4))
        self.trend_table = ttk.Frame(trends_settings)
        self.trend_table.pack(fill="x")
        column_specs = (
            ("Show", 5),
            ("Trend label", 16),
            ("Y column", 16),
            ("Chart #", 8),
            ("Side", 7),
            ("Min", 9),
            ("Max", 9),
            ("Color", 7),
            ("Units", 10),
            ("Excel number format", 18),
            ("", 7),
        )
        for column, (text, width) in enumerate(column_specs):
            self.trend_table.columnconfigure(
                column, minsize=width * 7, weight=1
            )
            ttk.Label(self.trend_table, text=text, width=width, anchor="w").grid(
                row=0, column=column, sticky="ew", padx=1
            )
        trend_buttons = ttk.Frame(trends_settings)
        trend_buttons.pack(anchor="w", pady=4)
        ttk.Button(
            trend_buttons, text="Add trend", command=self.add_trend
        ).pack(side="left", padx=(0, 6))

        hairlines_settings = self._section("Hairlines")
        hairlines_settings.pack(fill="x", pady=4)
        hairline_controls = ttk.Frame(hairlines_settings)
        hairline_controls.pack(fill="x", pady=2)
        ttk.Label(hairline_controls, text="Time:", width=8).pack(side="left")
        ttk.Entry(
            hairline_controls, textvariable=self.new_hairline_time, width=23
        ).pack(side="left")
        ttk.Button(
            hairline_controls,
            text="Add hairline",
            command=self.add_hairline,
        ).pack(side="left", padx=6)
        ttk.Label(
            hairlines_settings,
            text="Leave time blank to place a hairline at the visible-range midpoint; values between samples are interpolated.",
            foreground="#666",
        ).pack(anchor="w", pady=(0, 3))
        ttk.Checkbutton(
            hairlines_settings,
            text="Concise hairline info (show values only)",
            variable=self.concise_hairline_info,
            command=self._redraw_hairlines,
        ).pack(anchor="w", pady=(0, 3))
        self.hairline_rows_frame = ttk.Frame(hairlines_settings)
        self.hairline_rows_frame.pack(fill="x")
        self.hairline_info_toggle = ttk.Button(
            hairlines_settings,
            text="▼ Hairline info",
            command=self.toggle_hairline_info,
        )
        self.hairline_info_toggle.pack(anchor="w", pady=(4, 0))
        self.hairline_info_frame = ttk.Frame(hairlines_settings)
        self.hairline_info_frame.pack(fill="both", expand=True)
        self.hairline_values = ttk.Treeview(
            self.hairline_info_frame,
            columns=("hairline", "time", "trend", "value"),
            show="headings",
            height=8,
        )
        for column, heading, width in (
            ("hairline", "Hairline", 90),
            ("time", "Time", 180),
            ("trend", "Trend", 180),
            ("value", "Value", 150),
        ):
            self.hairline_values.heading(column, text=heading)
            self.hairline_values.column(column, width=width, anchor="w")
        readout_scroll = ttk.Scrollbar(
            self.hairline_info_frame,
            orient="vertical",
            command=self.hairline_values.yview,
        )
        self.hairline_values.configure(yscrollcommand=readout_scroll.set)
        self.hairline_values.pack(side="left", fill="both", expand=True, pady=(4, 8))
        readout_scroll.pack(side="left", fill="y", pady=(4, 8))

        ttk.Separator(self.parent).pack(fill="x", pady=(8, 6))
        ttk.Label(
            self.parent, text="Generated trend graph", font=("", 12, "bold")
        ).pack(anchor="w", pady=(0, 5))
        graph_buttons = ttk.Frame(self.parent)
        graph_buttons.pack(anchor="w", pady=(0, 5))
        ttk.Button(
            graph_buttons, text="Plot trends", command=self.plot_trends
        ).pack(side="left", padx=(0, 6))
        ttk.Button(
            graph_buttons,
            text="Save graph",
            command=lambda: self.save_figure_callback(
                self.figure, self.title.get().strip() or "trend graph"
            ),
        ).pack(side="left", padx=(0, 6))
        ttk.Button(
            graph_buttons,
            text="Copy graph to clipboard",
            command=lambda: self.copy_figure_callback(self.figure),
        ).pack(side="left")
        self.graph_frame = ttk.Frame(self.parent)
        self.graph_frame.pack(fill="both", expand=True)

    def _section(self, text, parent=None):
        return ttk.LabelFrame(parent or self.parent, text=text, padding=8)

    def toggle_hairline_info(self):
        if self.hairline_info_frame.winfo_manager():
            self.hairline_info_frame.pack_forget()
            self.hairline_info_toggle.configure(text="► Hairline info")
        else:
            self.hairline_info_frame.pack(
                fill="both", expand=True, after=self.hairline_info_toggle
            )
            self.hairline_info_toggle.configure(text="▼ Hairline info")

    def _entry_row(self, label, variable, parent=None, label_width=22):
        row = ttk.Frame(parent or self.parent)
        row.pack(fill="x", pady=2)
        ttk.Label(row, text=label, width=label_width).pack(side="left")
        ttk.Entry(row, textvariable=variable).pack(side="left", fill="x", expand=True)

    def _combo_row(self, label, variable, parent=None, label_width=22):
        row = ttk.Frame(parent or self.parent)
        row.pack(fill="x", pady=2)
        ttk.Label(row, text=label, width=label_width).pack(side="left")
        combo = ttk.Combobox(row, textvariable=variable, state="readonly")
        combo.pack(side="left", fill="x", expand=True)
        self._protect_combobox_wheel(combo)
        if label == "Time column:":
            self.time_combo = combo

    _protect_combobox_wheel = staticmethod(protect_combobox_wheel)

    def _update_colormap_colors(self, *_args):
        if getattr(self, "_applying_settings", False) or self.color_mode.get() != "Colormap":
            return
        cmap = matplotlib.colormaps[self.colormap.get()]
        for index, row in enumerate(self.trend_rows):
            if not row.color_override:
                color = to_hex(cmap(index / max(1, len(self.trend_rows) - 1)))
                row.color.set(color)
                row.color_button.configure(background=color)

    def add_trend(self):
        if len(self.trend_rows) >= MAX_TRENDS:
            self.set_status(f"A maximum of {MAX_TRENDS} trends can be plotted.", error=True)
            return

        row_id = self.next_trend_id
        self.next_trend_id += 1
        grid_row = len(self.trend_rows) + 1
        visible = tk.BooleanVar(value=True)
        column = tk.StringVar()
        label = tk.StringVar()
        chart_index = tk.StringVar(value="1")
        side = tk.StringVar(value="Left")
        minimum = tk.StringVar()
        maximum = tk.StringVar()
        color = tk.StringVar(value=self._default_color(row_id))
        unit = tk.StringVar()
        number_format = tk.StringVar(value=DEFAULT_NUMBER_FORMAT)
        default_label = ""

        checkbutton = ttk.Checkbutton(self.trend_table, variable=visible, width=5)
        checkbutton.grid(row=grid_row, column=0, sticky="w", padx=1, pady=1)
        label_entry = ttk.Entry(self.trend_table, textvariable=label, width=18)
        label_entry.grid(row=grid_row, column=1, sticky="ew", padx=1, pady=1)
        column_combo = ttk.Combobox(
            self.trend_table,
            textvariable=column,
            values=self.columns,
            state="readonly",
            width=18,
        )
        column_combo.grid(row=grid_row, column=2, sticky="ew", padx=1, pady=1)
        self._protect_combobox_wheel(column_combo)
        chart_index_entry = ttk.Entry(
            self.trend_table, textvariable=chart_index, width=8
        )
        chart_index_entry.grid(row=grid_row, column=3, sticky="ew", padx=1, pady=1)
        side_combo = ttk.Combobox(
            self.trend_table,
            textvariable=side,
            values=["Left", "Right"],
            state="readonly",
            width=8,
        )
        side_combo.grid(row=grid_row, column=4, sticky="ew", padx=1, pady=1)
        self._protect_combobox_wheel(side_combo)
        minimum_entry = ttk.Entry(self.trend_table, textvariable=minimum, width=10)
        minimum_entry.grid(row=grid_row, column=5, sticky="ew", padx=1, pady=1)
        maximum_entry = ttk.Entry(self.trend_table, textvariable=maximum, width=10)
        maximum_entry.grid(row=grid_row, column=6, sticky="ew", padx=1, pady=1)
        color_button = tk.Button(
            self.trend_table,
            text=" ",
            width=3,
            background=color.get(),
            command=lambda: self.choose_color(color, color_button),
        )
        color_button.grid(row=grid_row, column=7, sticky="ew", padx=1, pady=1)
        unit_entry = ttk.Entry(self.trend_table, textvariable=unit, width=12)
        unit_entry.grid(row=grid_row, column=8, sticky="ew", padx=1, pady=1)
        format_entry = ttk.Entry(
            self.trend_table, textvariable=number_format, width=22
        )
        format_entry.grid(row=grid_row, column=9, sticky="ew", padx=1, pady=1)
        remove_button = ttk.Button(self.trend_table, text="Remove")
        remove_button.grid(row=grid_row, column=10, sticky="ew", padx=1, pady=1)
        row = TrendRow(
            row_id,
            [
                checkbutton,
                label_entry,
                column_combo,
                chart_index_entry,
                side_combo,
                minimum_entry,
                maximum_entry,
                color_button,
                unit_entry,
                format_entry,
                remove_button,
            ],
            visible,
            column,
            label,
            chart_index,
            side,
            minimum,
            maximum,
            color,
            unit,
            number_format,
            default_label,
            column_combo,
            color_button,
        )
        remove_button.configure(command=lambda trend=row: self.remove_trend(trend))
        self.trend_rows.append(row)
        self._update_colormap_colors()
        self._choose_default_trend_column(row)
        self._update_default_label(row)
        column_combo.bind(
            "<<ComboboxSelected>>", lambda _event: self._update_default_label(row)
        )

    @staticmethod
    def _update_default_label(row):
        if not row.label.get().strip() or row.label.get() == row.default_label:
            row.label.set(row.column.get())
        row.default_label = row.column.get()

    def set_saved_settings_names(self, names):
        names = list(names)
        self.saved_settings_combo.configure(values=names)
        if self.saved_settings_name.get() not in names:
            self.saved_settings_name.set(names[0] if names else "")

    def export_settings(self):
        return {
            "version": 3,
            "time_column": self.time_column.get(),
            "title": self.title.get(),
            "start_time": self.start_time.get(),
            "end_time": self.end_time.get(),
            "color_mode": self.color_mode.get(),
            "colormap": self.colormap.get(),
            "trends": [
                {
                    "visible": row.visible.get(),
                    "column": row.column.get(),
                    "label": row.label.get(),
                    "chart_index": row.chart_index.get(),
                    "side": row.side.get(),
                    "minimum": row.minimum.get(),
                    "maximum": row.maximum.get(),
                    "color": row.color.get(),
                    "color_override": row.color_override,
                    "unit": row.unit.get(),
                    "number_format": row.number_format.get(),
                }
                for row in self.trend_rows
            ],
        }

    def validate_settings(self, settings, columns=None, source_frame=None):
        if (
            not isinstance(settings, dict)
            or isinstance(settings.get("version"), bool)
            or settings.get("version") not in (1, 2, 3)
        ):
            raise ValueError("The saved trend settings have an unsupported format.")
        columns = self.columns if columns is None else list(columns)
        source_frame = self._source_frame if source_frame is None else source_frame
        if source_frame is None or not columns:
            raise ValueError("Load the source columns before importing trend settings.")

        time_column = settings.get("time_column")
        if time_column not in columns:
            raise ValueError(
                f"The saved time column '{time_column}' is not in the selected source."
            )
        if not to_naive_datetimes(source_frame[time_column]).notna().any():
            raise ValueError(
                f"The saved time column '{time_column}' contains no valid date/time values."
            )
        color_mode = settings.get("color_mode")
        if color_mode not in ("Individual colors", "Colormap"):
            raise ValueError("The saved trend color mode is invalid.")
        colormap = settings.get("colormap")
        if colormap not in matplotlib.colormaps:
            raise ValueError(f"The saved trend colormap '{colormap}' is unavailable.")
        trends = settings.get("trends")
        if not isinstance(trends, list) or not 1 <= len(trends) <= MAX_TRENDS:
            raise ValueError(
                f"Saved trend settings must contain between 1 and {MAX_TRENDS} trends."
            )
        numeric_columns = self._numeric_columns(source_frame, columns)
        validated_trends = []
        for index, trend in enumerate(trends, start=1):
            if not isinstance(trend, dict):
                raise ValueError(f"Saved trend {index} is not a valid settings record.")
            column = trend.get("column")
            if column not in numeric_columns:
                raise ValueError(
                    f"Saved Y column '{column}' is not a numeric source column."
                )
            if not isinstance(trend.get("visible"), bool):
                raise ValueError(f"Saved trend {index} has an invalid visibility setting.")
            chart_index_text = str(trend.get("chart_index", 1)).strip()
            if not re.fullmatch(r"[1-9]\d*", chart_index_text):
                raise ValueError(
                    f"Saved trend {index} chart number must be a positive integer."
                )
            if trend["visible"]:
                values = pd.to_numeric(
                    source_frame[column], errors="coerce"
                ).to_numpy(dtype=float)
                if not np.isfinite(values).any():
                    raise ValueError(
                        f"Saved trend '{column}' contains no numeric data."
                    )
            side = trend.get("side")
            if side not in ("Left", "Right"):
                raise ValueError(f"Saved trend {index} has an invalid axis side.")
            color = trend.get("color")
            if not isinstance(color, str) or not is_color_like(color):
                raise ValueError(f"Saved trend {index} has an invalid color.")
            color_override = trend.get("color_override", False)
            if not isinstance(color_override, bool):
                raise ValueError(f"Saved trend {index} has an invalid color override.")
            minimum_text = str(trend.get("minimum", ""))
            maximum_text = str(trend.get("maximum", ""))
            minimum = self._parse_axis_limit(minimum_text, f"Trend {index} minimum")
            maximum = self._parse_axis_limit(maximum_text, f"Trend {index} maximum")
            if minimum is not None and maximum is not None and minimum >= maximum:
                raise ValueError(
                    f"Saved trend {index} minimum must be less than its maximum."
                )
            validated_trends.append(
                {
                    "visible": trend["visible"],
                    "column": column,
                    "label": str(trend.get("label", "")).strip() or column,
                    "chart_index": int(chart_index_text),
                    "side": side,
                    "minimum": minimum_text,
                    "maximum": maximum_text,
                    "color": color,
                    "color_override": color_override,
                    "unit": str(trend.get("unit", "")),
                    "number_format": str(
                        trend.get("number_format", DEFAULT_NUMBER_FORMAT)
                    ),
                }
            )

        start_time = str(settings.get("start_time", ""))
        end_time = str(settings.get("end_time", ""))
        start = parse_time(start_time, "Start time") if start_time.strip() else None
        end = parse_time(end_time, "End time") if end_time.strip() else None
        if start is not None and end is not None and start >= end:
            raise ValueError("Saved start time must be earlier than its end time.")
        return {
            "time_column": time_column,
            "title": str(settings.get("title", "Trend navigator")),
            "start_time": start_time,
            "end_time": end_time,
            "color_mode": color_mode,
            "colormap": colormap,
            "trends": validated_trends,
        }

    def apply_settings(self, settings):
        validated = self.validate_settings(settings)
        validated_trends = validated["trends"]
        while len(self.trend_rows) < len(validated_trends):
            self.add_trend()
        while len(self.trend_rows) > len(validated_trends):
            self.remove_trend(self.trend_rows[-1])

        for row, settings_row in zip(self.trend_rows, validated_trends):
            row.visible.set(settings_row["visible"])
            row.column.set(settings_row["column"])
            row.label.set(settings_row["label"])
            row.chart_index.set(str(settings_row["chart_index"]))
            row.default_label = settings_row["column"]
            row.side.set(settings_row["side"])
            row.minimum.set(settings_row["minimum"])
            row.maximum.set(settings_row["maximum"])
            row.color.set(settings_row["color"])
            row.color_override = settings_row["color_override"]
            row.color_button.configure(background=settings_row["color"])
            row.unit.set(settings_row["unit"])
            row.number_format.set(settings_row["number_format"])

        self.title.set(validated["title"])
        self.time_column.set(validated["time_column"])
        self.start_time.set(validated["start_time"])
        self.end_time.set(validated["end_time"])
        self._applying_settings = True
        try:
            self.color_mode.set(validated["color_mode"])
            self.colormap.set(validated["colormap"])
        finally:
            self._applying_settings = False
        self._update_colormap_colors()
        for hairline in list(self.hairlines):
            self.delete_hairline(hairline.hairline_id)

    def save_settings(self):
        name = self.settings_name.get().strip()
        if not name:
            self.set_status("Enter a name for this trend settings set.", error=True)
            return
        if self.save_settings_callback is None:
            self.set_status("Trend settings cannot be saved from this window.", error=True)
            return
        try:
            self.save_settings_callback(name, self.export_settings())
        except (ValueError, TypeError, KeyError, IndexError, tk.TclError) as exc:
            self.set_status(str(exc), error=True)

    def refresh_settings(self):
        if self.refresh_settings_callback is None:
            self.set_status("Trend settings cannot be refreshed from this window.", error=True)
            return
        self.refresh_settings_callback()

    def load_settings(self):
        name = self.saved_settings_name.get()
        if not name:
            self.set_status("Select a saved trend settings set to load.", error=True)
            return
        if self.load_settings_callback is None:
            self.set_status("Trend settings cannot be loaded from this window.", error=True)
            return
        self.load_settings_callback(name)

    @staticmethod
    def _default_color(index):
        return to_hex(matplotlib.colormaps["tab10"]((index - 1) % 10 / 10))

    def _choose_default_trend_column(self, row):
        available = self._numeric_columns()
        if available and row.column.get() not in available:
            index = (row.row_id - 1) % len(available)
            row.column.set(available[index])

    def _numeric_columns(self, source_frame=None, columns=None):
        source_frame = self._source_frame if source_frame is None else source_frame
        columns = self.columns if columns is None else columns
        if source_frame is None:
            return columns
        numeric = []
        for column in columns:
            series = source_frame[column]
            if pd.api.types.is_datetime64_any_dtype(series):
                continue
            if series.dtype == object and any(
                isinstance(value, (pd.Timestamp, np.datetime64, datetime, date))
                for value in series.dropna().head(10)
            ):
                continue
            if self._looks_like_excel_serial(series):
                continue
            values = pd.to_numeric(series, errors="coerce")
            if values.notna().any():
                numeric.append(column)
        return numeric

    @staticmethod
    def _looks_like_excel_serial(series):
        try:
            values = pd.to_numeric(series, errors="coerce").to_numpy(dtype=float)
        except (TypeError, ValueError):
            return False
        finite = values[np.isfinite(values)]
        return finite.size > 0 and np.all((finite >= 20000) & (finite <= 100000))

    def remove_trend(self, trend):
        if trend in self.trend_rows:
            for widget in trend.widgets:
                widget.destroy()
            self.trend_rows.remove(trend)
            for index, row in enumerate(self.trend_rows, start=1):
                for widget in row.widgets:
                    widget.grid_configure(row=index)
            self._update_colormap_colors()

    def set_columns(self, columns: list[str], frame: pd.DataFrame | None = None) -> None:
        self.columns = list(columns)
        self._source_frame = None if frame is None else frame.copy()
        self.time_combo.configure(values=self.columns)
        if not self.columns:
            self.time_column.set("")
            for row in self.trend_rows:
                row.column_combo.configure(values=[])
                row.column.set("")
            return

        if self._source_frame is not None:
            self._source_frame.columns = [
                str(column) for column in self._source_frame.columns
            ]
        time_columns = []
        if self._source_frame is not None:
            for column in self.columns:
                series = self._source_frame[column]
                if pd.api.types.is_datetime64_any_dtype(series):
                    time_columns.append(column)
                elif series.dtype == object:
                    sample = series.dropna().head(10)
                    if any(
                        isinstance(value, (pd.Timestamp, np.datetime64, datetime, date))
                        for value in sample
                    ):
                        time_columns.append(column)
                if self._looks_like_excel_serial(series):
                    time_columns.append(column)
        if self.time_column.get() not in self.columns:
            self.time_column.set((time_columns or self.columns)[0])
        numeric = self._numeric_columns()
        for row in self.trend_rows:
            row.column_combo.configure(values=numeric)
            self._choose_default_trend_column(row)
            self._update_default_label(row)

    def choose_color(self, variable, button):
        selected = colorchooser.askcolor(
            color=variable.get(), parent=self.parent.winfo_toplevel()
        )[1]
        if selected:
            variable.set(selected)
            button.configure(background=selected)
            for row in self.trend_rows:
                if row.color is variable:
                    row.color_override = True
                    break

    def _selected_trends(self):
        selected = [row for row in self.trend_rows if row.visible.get()]
        if not selected:
            raise ValueError("Enable at least one trend to plot.")
        return selected

    @staticmethod
    def _parse_axis_limit(text, label):
        text = text.strip()
        if not text:
            return None
        try:
            value = float(text)
        except ValueError as exc:
            raise ValueError(f"{label} must be a finite number or left blank.") from exc
        if not np.isfinite(value):
            raise ValueError(f"{label} must be a finite number or left blank.")
        return value

    def _prepare_trends(self, frame, selected, columns):
        time_name = self.time_column.get()
        if time_name not in columns:
            raise ValueError("Select a valid time column.")
        times = to_naive_datetimes(frame[time_name])
        if not times.notna().any():
            raise ValueError("The selected time column contains no valid date/time values.")

        colors = {row.row_id: row.color.get() for row in self.trend_rows}

        configs = {}
        prepared = []
        for row in selected:
            column_name = row.column.get()
            if column_name not in columns:
                raise ValueError(f"Select a valid Y column for trend {row.row_id}.")
            values = pd.to_numeric(frame[column_name], errors="coerce").to_numpy(
                dtype=float
            )
            valid_times = times.notna().to_numpy()
            if not (valid_times & np.isfinite(values)).any():
                raise ValueError(f"Trend '{column_name}' has no usable numeric data.")

            minimum = self._parse_axis_limit(
                row.minimum.get(), f"Trend '{column_name}' minimum"
            )
            maximum = self._parse_axis_limit(
                row.maximum.get(), f"Trend '{column_name}' maximum"
            )
            chart_index_text = row.chart_index.get().strip()
            if not re.fullmatch(r"[1-9]\d*", chart_index_text):
                raise ValueError(
                    f"Trend '{column_name}' chart number must be a positive integer."
                )
            chart_index = int(chart_index_text)
            if minimum is not None and maximum is not None and minimum >= maximum:
                raise ValueError(
                    f"Trend '{column_name}' minimum must be less than its maximum."
                )
            number_format = row.number_format.get().strip() or DEFAULT_NUMBER_FORMAT
            shared_limits = minimum is not None and maximum is not None
            axis_key = (
                chart_index,
                row.side.get(),
                minimum,
                maximum,
                None if shared_limits else row.row_id,
            )
            configs.setdefault(axis_key, number_format)

            trend_times = pd.DatetimeIndex(times[valid_times]).to_pydatetime()
            time_numbers = mdates.date2num(trend_times)
            trend_values = values[valid_times]
            order = np.argsort(time_numbers)
            sorted_times = time_numbers[order]
            sorted_values = trend_values[order]
            finite_values = np.isfinite(sorted_values)
            prepared.append(
                {
                    "row": row,
                    "column": column_name,
                    "label": row.label.get().strip() or column_name,
                    "chart_index": chart_index,
                    "unit": row.unit.get().strip(),
                    "axis": axis_key,
                    "minimum": minimum,
                    "maximum": maximum,
                    "number_format": number_format,
                    "times": sorted_times[finite_values],
                    "values": sorted_values[finite_values],
                    "plot_times": sorted_times,
                    "plot_values": sorted_values,
                    "color": colors[row.row_id],
                }
            )
        return prepared, configs

    def _parse_display_range(self, prepared):
        all_times = np.concatenate([series["times"] for series in prepared])
        data_min = float(np.min(all_times))
        data_max = float(np.max(all_times))
        lower = (
            mdates.date2num(parse_time(self.start_time.get(), "Start time").to_pydatetime())
            if self.start_time.get().strip()
            else data_min
        )
        upper = (
            mdates.date2num(parse_time(self.end_time.get(), "End time").to_pydatetime())
            if self.end_time.get().strip()
            else data_max
        )
        if lower >= upper:
            if data_min == data_max and not self.start_time.get().strip() and not self.end_time.get().strip():
                lower -= 1 / 86400
                upper += 1 / 86400
            else:
                raise ValueError("Start time must be earlier than end time.")
        return data_min, data_max, lower, upper

    def plot_trends(self):
        try:
            frame = self.data_provider()
            if frame is None or frame.empty:
                raise ValueError("The selected source contains no data rows.")
            frame = frame.copy()
            frame.columns = [str(column) for column in frame.columns]
            if len(set(frame.columns)) != len(frame.columns):
                raise ValueError("The selected source has duplicate column names.")
            selected = self._selected_trends()
            prepared, configs = self._prepare_trends(
                frame, selected, list(frame.columns)
            )
            data_min, data_max, display_min, display_max = self._parse_display_range(
                prepared
            )
            axis_keys = sorted(
                configs,
                key=lambda key: (
                    key[0],
                    key[1] != "Left",
                    key[1],
                    key[2] is not None,
                    key[2] if key[2] is not None else 0,
                    key[3] is not None,
                    key[3] if key[3] is not None else 0,
                    key[4] if key[4] is not None else 0,
                ),
            )
            chart_indices = sorted({series["chart_index"] for series in prepared})
            fixed_size = self.figure_size_inches()
            if fixed_size is None:
                figure_size = (10.0, max(5.5, 2.7 * len(chart_indices) + 0.8))
            else:
                figure_size = fixed_size
            figure = Figure(figsize=figure_size)
            axes = {}
            chart_axes = {}
            axis_ranks = {}
            previous_base = None
            for subplot_index, chart_index in enumerate(chart_indices, start=1):
                base_axes = figure.add_subplot(
                    len(chart_indices),
                    1,
                    subplot_index,
                    sharex=previous_base,
                )
                previous_base = base_axes
                chart_axes[chart_index] = base_axes
                chart_keys = [key for key in axis_keys if key[0] == chart_index]
                primary_key = chart_keys[0]
                primary_side = primary_key[1].lower()
                base_axes.spines[
                    "right" if primary_side == "left" else "left"
                ].set_visible(False)
                axes[primary_key] = base_axes
                side_counts = {"Left": 0, "Right": 0}
                side_counts[primary_key[1]] += 1
                axis_ranks[primary_key] = 0
                for axis_key in chart_keys[1:]:
                    axis = base_axes.twinx()
                    side = axis_key[1].lower()
                    side_counts[axis_key[1]] += 1
                    axis_rank = side_counts[axis_key[1]] - 1
                    axis.spines[
                        "left" if side == "right" else "right"
                    ].set_visible(False)
                    axis.patch.set_visible(False)
                    axes[axis_key] = axis
                    axis_ranks[axis_key] = axis_rank

            handles = []
            labels = []
            plotted_series = []
            for series in prepared:
                axis = axes[series["axis"]]
                line, = axis.plot(
                    mdates.num2date(series["plot_times"]),
                    series["plot_values"],
                    color=series["color"],
                    label=series["label"],
                    linewidth=1.4,
                )
                handles.append(line)
                labels.append(series["label"])
                plotted_series.append(series)

            for axis_key, axis in axes.items():
                axis_trends = [
                    series for series in prepared if series["axis"] == axis_key
                ]
                minimum, maximum = axis_key[2:4]
                number_format = configs[axis_key]
                shared = len(axis_trends) > 1
                color = "black" if shared else axis_trends[0]["color"]
                label = ", ".join(
                    f"{series['label']} ({series['unit']})"
                    if series["unit"]
                    else series["label"]
                    for series in axis_trends
                )
                side = axis_key[1].lower()
                axis.spines["left" if side == "left" else "right"].set_color(color)
                if side == "left":
                    axis.yaxis.tick_left()
                    axis.yaxis.set_label_position("left")
                else:
                    axis.yaxis.tick_right()
                    axis.yaxis.set_label_position("right")
                axis.set_ylabel(label, color=color)
                axis.tick_params(axis="y", colors=color)
                axis.yaxis.set_major_formatter(
                    FuncFormatter(
                        lambda value, _position, fmt=number_format: format_excel_number(
                            value, fmt
                        )
                    )
                )
                if minimum is not None or maximum is not None:
                    axis.set_ylim(bottom=minimum, top=maximum)

            base_axes = chart_axes[chart_indices[0]]
            base_axes.set_xlim(display_min, display_max)
            locator = mdates.AutoDateLocator()
            base_axes.xaxis.set_major_locator(locator)
            base_axes.xaxis.set_major_formatter(mdates.ConciseDateFormatter(locator))
            for chart_base in chart_axes.values():
                chart_base.grid(True, axis="both", alpha=0.25)
            legend_columns = min(6, max(1, len(labels)))
            legend_rows = int(np.ceil(len(labels) / legend_columns))
            figure.legend(
                handles,
                labels,
                loc="upper center",
                bbox_to_anchor=(0.5, 0.97),
                ncol=legend_columns,
            )
            figure.suptitle(self.title.get().strip() or "Trend navigator", y=0.99)
            figure.subplots_adjust(
                left=0.1,
                right=0.9,
                top=max(0.82, 0.91 - 0.025 * (legend_rows - 1)),
                bottom=0.1,
                hspace=0.2,
            )
            left_margin_inches, right_margin_inches = layout_y_axes(
                figure, axes, axis_ranks, fixed_size
            )
            if fixed_size is None:
                figure.set_size_inches(
                    max(10, left_margin_inches + right_margin_inches + 6.5),
                    figure_size[1],
                )
            figure_width = figure.get_size_inches()[0]
            figure.subplots_adjust(
                left=left_margin_inches / figure_width,
                right=1 - right_margin_inches / figure_width,
            )

            if self.canvas is not None:
                self.canvas.get_tk_widget().destroy()
            self.figure = figure
            self.figure_fixed_size = fixed_size is not None
            self.base_axes = base_axes
            self.chart_axes = chart_axes
            self.axes_by_group = axes
            self.plotted_series = plotted_series
            self.data_time_range = (data_min, data_max)
            self.canvas = FigureCanvasTkAgg(figure, master=self.graph_frame)
            self.canvas.draw()
            if self.figure_fixed_size:
                self.canvas.get_tk_widget().pack(anchor="nw")
            else:
                self.canvas.get_tk_widget().pack(fill="both", expand=True)
            self.canvas.mpl_connect("button_press_event", self._on_plot_press)
            self.canvas.mpl_connect("motion_notify_event", self._on_plot_motion)
            self.canvas.mpl_connect("button_release_event", self._on_plot_release)
            self.canvas.mpl_connect("scroll_event", self._on_plot_scroll)
            self._redraw_hairlines()
            self.set_status(
                f"Plotted {len(plotted_series)} trends over "
                f"{len(np.unique(np.concatenate([s['times'] for s in plotted_series])))} "
                f"timestamps in {len(chart_indices)} chart(s)."
            )
            return True
        except (ValueError, TypeError, KeyError, IndexError, OverflowError) as exc:
            self.set_status(str(exc), error=True)
            return False

    @staticmethod
    def _time_text(value):
        return mdates.num2date(value).replace(tzinfo=None).strftime(
            TIME_DISPLAY_FORMAT
        )

    def _set_visible_range(self, lower, upper):
        if self.base_axes is None:
            return
        self.base_axes.set_xlim(lower, upper)
        self.start_time.set(self._time_text(lower))
        self.end_time.set(self._time_text(upper))
        self._redraw_hairlines()

    def _on_plot_press(self, event):
        if (
            self.base_axes is None
            or event.inaxes not in self.axes_by_group.values()
            or event.xdata is None
        ):
            return
        if event.button == 3:
            for hairline in reversed(self.hairlines):
                for artist in hairline.artists:
                    contains, _details = artist.contains(event)
                    if contains:
                        self.delete_hairline(hairline.hairline_id)
                        return
            return
        if event.button != 1:
            return
        if event.dblclick:
            timestamp = pd.Timestamp(
                mdates.num2date(event.xdata).replace(tzinfo=None)
            )
            try:
                self._create_hairline(timestamp)
            except ValueError as exc:
                self.set_status(str(exc), error=True)
            return

        for hairline in reversed(self.hairlines):
            for artist in hairline.artists:
                contains, _details = artist.contains(event)
                if contains:
                    self._drag_hairline = hairline
                    return
        axes_width = self.base_axes.bbox.width
        if event.x is not None and axes_width > 0:
            self._pan_start = (event.x, *self.base_axes.get_xlim(), axes_width)

    def _on_plot_motion(self, event):
        if self._drag_hairline is not None:
            if self.base_axes is None or event.xdata is None:
                return
            lower, upper = self.base_axes.get_xlim()
            position = min(max(event.xdata, lower), upper)
            timestamp = pd.Timestamp(
                mdates.num2date(position).replace(tzinfo=None)
            )
            self._drag_hairline.time.set(timestamp.strftime(TIME_DISPLAY_FORMAT))
            self._redraw_hairlines()
        elif self._pan_start is not None:
            if event.x is None:
                return
            start_pixel, lower, upper, axes_width = self._pan_start
            offset = (start_pixel - event.x) * (upper - lower) / axes_width
            self._set_visible_range(lower + offset, upper + offset)

    def _on_plot_release(self, _event):
        self._drag_hairline = None
        self._pan_start = None

    def _on_plot_scroll(self, event):
        if (
            self.base_axes is None
            or event.inaxes not in self.axes_by_group.values()
            or event.xdata is None
        ):
            return
        lower, upper = self.base_axes.get_xlim()
        zoom_factor = 0.8 if event.button == "down" else 1.25
        self._set_visible_range(
            event.xdata - (event.xdata - lower) * zoom_factor,
            event.xdata + (upper - event.xdata) * zoom_factor,
        )

    def apply_time_range(self):
        if self.base_axes is None or self.data_time_range is None:
            self.set_status("Plot trends before applying a time range.", error=True)
            return
        try:
            lower = (
                mdates.date2num(
                    parse_time(self.start_time.get(), "Start time").to_pydatetime()
                )
                if self.start_time.get().strip()
                else self.data_time_range[0]
            )
            upper = (
                mdates.date2num(
                    parse_time(self.end_time.get(), "End time").to_pydatetime()
                )
                if self.end_time.get().strip()
                else self.data_time_range[1]
            )
            if lower >= upper:
                raise ValueError("Start time must be earlier than end time.")
            self._set_visible_range(lower, upper)
            self.set_status("Applied the visible time range.")
        except ValueError as exc:
            self.set_status(str(exc), error=True)

    def reset_time_range(self):
        self.start_time.set("")
        self.end_time.set("")
        if self.base_axes is not None and self.data_time_range is not None:
            lower, upper = self.data_time_range
            if lower == upper:
                lower -= 1 / 86400
                upper += 1 / 86400
            self.base_axes.set_xlim(lower, upper)
            self._redraw_hairlines()

    def add_hairline(self):
        if self.base_axes is None or self.data_time_range is None:
            self.set_status("Plot trends before adding a hairline.", error=True)
            return
        try:
            if self.new_hairline_time.get().strip():
                timestamp = parse_time(self.new_hairline_time.get(), "Hairline time")
            else:
                lower, upper = self.base_axes.get_xlim()
                timestamp = pd.Timestamp(mdates.num2date((lower + upper) / 2)).tz_localize(None)
            self._create_hairline(timestamp)
        except ValueError as exc:
            self.set_status(str(exc), error=True)

    def _create_hairline(self, timestamp):
        base_axes = self.base_axes
        if base_axes is None:
            raise ValueError("Plot trends before adding a hairline.")
        position = float(mdates.date2num(timestamp.to_pydatetime()))
        lower, upper = base_axes.get_xlim()
        if not lower <= position <= upper:
            raise ValueError("Hairline time must fall within the visible time range.")
        hairline_id = self.next_hairline_id
        self.next_hairline_id += 1
        color = to_hex(matplotlib.colormaps["tab10"]((hairline_id - 1) % 10 / 10))
        row = ttk.Frame(self.hairline_rows_frame)
        row.pack(fill="x", pady=1)
        time_var = tk.StringVar(value=timestamp.strftime(TIME_DISPLAY_FORMAT))
        ttk.Label(row, text=f"Hairline {hairline_id}", width=14).pack(side="left")
        ttk.Entry(row, textvariable=time_var, width=23).pack(side="left")
        color_button = tk.Button(row, text=" ", width=3, background=color)
        color_button.configure(
            command=lambda: self._choose_hairline_color(hairline_id, color_button)
        )
        color_button.pack(side="left", padx=3)
        ttk.Button(
            row,
            text="Update",
            command=lambda: self.update_hairline(hairline_id),
        ).pack(side="left", padx=2)
        ttk.Button(
            row,
            text="Delete",
            command=lambda: self.delete_hairline(hairline_id),
        ).pack(side="left", padx=2)
        self.hairlines.append(Hairline(hairline_id, row, time_var, color))
        self.new_hairline_time.set("")
        self._redraw_hairlines()

    def update_hairline(self, hairline_id):
        hairline = self._find_hairline(hairline_id)
        if hairline is None or self.base_axes is None:
            return
        try:
            timestamp = parse_time(hairline.time.get(), "Hairline time")
            position = mdates.date2num(timestamp.to_pydatetime())
            lower, upper = self.base_axes.get_xlim()
            if not lower <= position <= upper:
                raise ValueError("Hairline time must fall within the visible time range.")
            hairline.time.set(timestamp.strftime(TIME_DISPLAY_FORMAT))
            self._redraw_hairlines()
        except ValueError as exc:
            self.set_status(str(exc), error=True)

    def delete_hairline(self, hairline_id):
        hairline = self._find_hairline(hairline_id)
        if hairline is not None:
            if self._drag_hairline is hairline:
                self._drag_hairline = None
            for artist in hairline.artists + hairline.readout_artists:
                artist.remove()
            hairline.artists.clear()
            hairline.readout_artists.clear()
            hairline.frame.destroy()
            self.hairlines.remove(hairline)
            self._redraw_hairlines()

    def _find_hairline(self, hairline_id):
        return next(
            (item for item in self.hairlines if item.hairline_id == hairline_id), None
        )

    def _choose_hairline_color(self, hairline_id, button):
        hairline = self._find_hairline(hairline_id)
        if hairline is None:
            return
        selected = colorchooser.askcolor(
            color=hairline.color, parent=self.parent.winfo_toplevel()
        )[1]
        if selected:
            hairline.color = selected
            button.configure(background=selected)
            self._redraw_hairlines()

    def _redraw_hairlines(self):
        base_axes = self.base_axes
        if base_axes is None:
            return
        for hairline in self.hairlines:
            for artist in hairline.artists + hairline.readout_artists:
                artist.remove()
            hairline.artists.clear()
            hairline.readout_artists.clear()
            timestamp = parse_time(hairline.time.get(), "Hairline time")
            position = float(mdates.date2num(timestamp.to_pydatetime()))
            lower, upper = base_axes.get_xlim()
            for chart_index, chart_axes in self.chart_axes.items():
                hairline.artists.append(
                    chart_axes.axvline(
                        timestamp,
                        color=hairline.color,
                        linestyle="--",
                        linewidth=1.3,
                        label=f"Hairline {hairline.hairline_id}",
                        picker=6,
                    )
                )
                if not lower <= position <= upper:
                    continue
                chart_series = [
                    series
                    for series in self.plotted_series
                    if series["chart_index"] == chart_index
                ]
                readout_rows = []
                if not self.concise_hairline_info.get():
                    readout_rows.append(
                        TextArea(
                            f"Hairline {hairline.hairline_id}  "
                            f"{timestamp.strftime(TIME_DISPLAY_FORMAT)}",
                            textprops={
                                "color": hairline.color,
                                "size": 8,
                                "weight": "bold",
                            },
                        )
                    )
                for series in chart_series:
                    value = self._interpolated_value(series, position)
                    value_text = (
                        "—"
                        if value is None
                        else format_excel_number(value, series["number_format"])
                    )
                    readout_text = (
                        value_text
                        if self.concise_hairline_info.get()
                        else f"{series['label']}: {value_text}"
                    )
                    readout_rows.append(
                        TextArea(
                            readout_text,
                            textprops={"color": series["color"], "size": 8},
                        )
                    )
                readout = VPacker(
                    children=readout_rows, align="left", pad=2, sep=1
                )
                place_left = position > lower + (upper - lower) * 0.65
                box = AnnotationBbox(
                    readout,
                    (position, 1),
                    xycoords=chart_axes.get_xaxis_transform(),
                    xybox=(-6 if place_left else 6, -6),
                    boxcoords="offset points",
                    box_alignment=(1 if place_left else 0, 1),
                    bboxprops={
                        "boxstyle": "round,pad=0.3",
                        "facecolor": "white",
                        "edgecolor": hairline.color,
                        "alpha": 0.94,
                    },
                    annotation_clip=False,
                    zorder=10,
                )
                chart_axes.add_artist(box)
                hairline.readout_artists.append(box)
        if self.canvas is not None:
            self.canvas.draw_idle()
        self._update_hairline_values()

    @staticmethod
    def _interpolated_value(series, position):
        times = series["times"]
        if position < times[0] or position > times[-1]:
            return None
        return float(np.interp(position, times, series["values"]))

    def _update_hairline_values(self):
        self.hairline_values.delete(*self.hairline_values.get_children())
        for hairline in self.hairlines:
            timestamp = parse_time(hairline.time.get(), "Hairline time")
            position = mdates.date2num(timestamp.to_pydatetime())
            for series in self.plotted_series:
                value = self._interpolated_value(series, position)
                if value is None:
                    value_text = "—"
                else:
                    value_text = format_excel_number(value, series["number_format"])
                self.hairline_values.insert(
                    "",
                    "end",
                    values=(
                        f"Hairline {hairline.hairline_id}",
                        timestamp.strftime(TIME_DISPLAY_FORMAT),
                        series["label"],
                        value_text,
                    ),
                )
