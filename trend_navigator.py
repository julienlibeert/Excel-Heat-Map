"""Time-based multi-series graph controls and plotting for the Excel heat-map app."""

from __future__ import annotations

import re
import tkinter as tk
from collections.abc import Callable
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
from tkinter import colorchooser, ttk

import matplotlib
import matplotlib.dates as mdates
import numpy as np
import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.colors import to_hex
from matplotlib.figure import Figure
from matplotlib.widgets import SpanSelector
from matplotlib.ticker import FuncFormatter


MAX_TRENDS = 16
DEFAULT_NUMBER_FORMAT = "#,##0.00"
TIME_DISPLAY_FORMAT = "%Y-%m-%d %H:%M:%S"
_NUMBER_PATTERN = re.compile(
    r"[#0?][#0?,]*(?:\.[#0?]+)?(?:[Ee][+-]?[0#?]+)?"
)


def _split_format_sections(number_format: str) -> list[str]:
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
    """Format common Excel numeric formats for trend-axis ticks.

    Supports section-based positive/negative/zero formats, decimals, grouping,
    percentages, scientific notation, quoted literals, and unit suffixes.
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
    try:
        timestamp = pd.Timestamp(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{label} must be a valid date/time.") from exc
    if pd.isna(timestamp):
        raise ValueError(f"{label} must be a valid date/time.")
    if timestamp.tz is not None:
        timestamp = timestamp.tz_convert(None)
    return timestamp


def _to_naive_datetimes(values: pd.Series) -> pd.Series:
    try:
        return pd.to_datetime(values, errors="coerce", utc=True).dt.tz_convert(None)
    except (TypeError, ValueError) as exc:
        raise ValueError("The selected time column cannot be parsed as dates.") from exc


@dataclass
class TrendRow:
    row_id: int
    frame: ttk.Frame
    visible: tk.BooleanVar
    column: tk.StringVar
    axis: tk.StringVar
    minimum: tk.StringVar
    maximum: tk.StringVar
    color: tk.StringVar
    unit: tk.StringVar
    number_format: tk.StringVar
    column_combo: ttk.Combobox
    color_button: tk.Button


@dataclass
class Hairline:
    hairline_id: int
    frame: ttk.Frame
    time: tk.StringVar
    color: str
    artist: object | None = None


class TrendNavigator:
    """Interactive trend graph UI embedded in one graph-type tab."""

    def __init__(
        self,
        parent: ttk.Frame,
        data_provider: Callable[[], pd.DataFrame],
        save_figure: Callable[[Figure | None, str], None],
        copy_figure: Callable[[Figure | None], None],
        set_status: Callable[..., None],
    ) -> None:
        self.parent = parent
        self.data_provider = data_provider
        self.save_figure_callback = save_figure
        self.copy_figure_callback = copy_figure
        self.set_status = set_status
        self.columns = []
        self._source_frame = None
        self.trend_rows: list[TrendRow] = []
        self.next_trend_id = 1
        self.hairlines: list[Hairline] = []
        self.next_hairline_id = 1
        self.figure = None
        self.canvas = None
        self.span_selector = None
        self.base_axes = None
        self.axes_by_number = {}
        self.plotted_series = []
        self.data_time_range = None

        self.time_column = tk.StringVar()
        self.title = tk.StringVar(value="Trend navigator")
        self.start_time = tk.StringVar()
        self.end_time = tk.StringVar()
        self.color_mode = tk.StringVar(value="Individual colors")
        self.colormap = tk.StringVar(value="viridis")
        self.new_hairline_time = tk.StringVar()

        self._build_controls()
        self.add_trend()

    def _build_controls(self):
        self._entry_row("Graph title:", self.title)
        self._combo_row("Time column:", self.time_column)

        self._section("Visible time range")
        range_row = ttk.Frame(self.parent)
        range_row.pack(fill="x", pady=2)
        ttk.Label(range_row, text="Start:", width=8).pack(side="left")
        ttk.Entry(range_row, textvariable=self.start_time, width=22).pack(
            side="left", padx=(0, 8)
        )
        ttk.Label(range_row, text="End:", width=6).pack(side="left")
        ttk.Entry(range_row, textvariable=self.end_time, width=22).pack(side="left")
        ttk.Label(range_row, text="YYYY-MM-DD HH:MM:SS").pack(
            side="left", padx=(8, 0)
        )
        ttk.Label(
            self.parent,
            text="You can also drag across the chart to select a time range.",
            foreground="#666",
        ).pack(anchor="w", padx=(22, 0), pady=(0, 3))
        ttk.Button(
            range_row, text="Apply", command=self.apply_time_range
        ).pack(side="left", padx=(8, 4))
        ttk.Button(
            range_row, text="Auto range", command=self.reset_time_range
        ).pack(side="left")

        self._section("Trend colors")
        color_row = ttk.Frame(self.parent)
        color_row.pack(fill="x", pady=2)
        ttk.Label(color_row, text="Color selection:", width=22).pack(side="left")
        ttk.Combobox(
            color_row,
            textvariable=self.color_mode,
            values=["Individual colors", "Colormap"],
            state="readonly",
            width=22,
        ).pack(side="left")
        ttk.Label(color_row, text="Colormap:", padding=(12, 0)).pack(side="left")
        ttk.Combobox(
            color_row,
            textvariable=self.colormap,
            values=sorted(matplotlib.colormaps),
            state="readonly",
            width=20,
        ).pack(side="left")

        self._section("Trends (up to 16)")
        header = ttk.Frame(self.parent)
        header.pack(fill="x")
        for text, width in (
            ("Show", 5),
            ("Y column", 18),
            ("Y axis", 7),
            ("Min", 10),
            ("Max", 10),
            ("Color", 7),
            ("Units", 12),
            ("Excel number format", 22),
        ):
            ttk.Label(header, text=text, width=width).pack(side="left", padx=1)
        ttk.Label(
            self.parent,
            text="Each trend starts on its own axis. Trends sharing an axis must use the same range, units, and format.",
            foreground="#666",
        ).pack(anchor="w", pady=(1, 3))
        self.trend_rows_frame = ttk.Frame(self.parent)
        self.trend_rows_frame.pack(fill="x")
        trend_buttons = ttk.Frame(self.parent)
        trend_buttons.pack(anchor="w", pady=4)
        ttk.Button(
            trend_buttons, text="Add trend", command=self.add_trend
        ).pack(side="left", padx=(0, 6))

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

        self._section("Hairlines")
        hairline_controls = ttk.Frame(self.parent)
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
            self.parent,
            text="Leave time blank to place a hairline at the visible-range midpoint; values between samples are interpolated.",
            foreground="#666",
        ).pack(anchor="w", padx=(22, 0), pady=(0, 3))
        self.hairline_rows_frame = ttk.Frame(self.parent)
        self.hairline_rows_frame.pack(fill="x")
        self.hairline_values = ttk.Treeview(
            self.parent,
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
            self.parent, orient="vertical", command=self.hairline_values.yview
        )
        self.hairline_values.configure(yscrollcommand=readout_scroll.set)
        self.hairline_values.pack(side="left", fill="x", expand=True, pady=(4, 8))
        readout_scroll.pack(side="left", fill="y", pady=(4, 8))

    def _section(self, text):
        ttk.Label(self.parent, text=text, font=("", 10, "bold")).pack(
            anchor="w", pady=(8, 3)
        )

    def _entry_row(self, label, variable):
        row = ttk.Frame(self.parent)
        row.pack(fill="x", pady=2)
        ttk.Label(row, text=label, width=22).pack(side="left")
        ttk.Entry(row, textvariable=variable).pack(side="left", fill="x", expand=True)

    def _combo_row(self, label, variable):
        row = ttk.Frame(self.parent)
        row.pack(fill="x", pady=2)
        ttk.Label(row, text=label, width=22).pack(side="left")
        combo = ttk.Combobox(row, textvariable=variable, state="readonly")
        combo.pack(side="left", fill="x", expand=True)
        if label == "Time column:":
            self.time_combo = combo

    def add_trend(self):
        if len(self.trend_rows) >= MAX_TRENDS:
            self.set_status(f"A maximum of {MAX_TRENDS} trends can be plotted.", error=True)
            return

        row_id = self.next_trend_id
        self.next_trend_id += 1
        used_axes = {int(row.axis.get()) for row in self.trend_rows}
        axis_number = next(
            (number for number in range(1, MAX_TRENDS + 1) if number not in used_axes),
            row_id,
        )
        frame = ttk.Frame(self.trend_rows_frame)
        frame.pack(fill="x", pady=1)
        visible = tk.BooleanVar(value=True)
        column = tk.StringVar()
        axis = tk.StringVar(value=str(axis_number))
        minimum = tk.StringVar()
        maximum = tk.StringVar()
        color = tk.StringVar(value=self._default_color(row_id))
        unit = tk.StringVar()
        number_format = tk.StringVar(value=DEFAULT_NUMBER_FORMAT)

        ttk.Checkbutton(frame, variable=visible, width=5).pack(side="left", padx=1)
        column_combo = ttk.Combobox(
            frame,
            textvariable=column,
            values=self.columns,
            state="readonly",
            width=18,
        )
        column_combo.pack(side="left", padx=1)
        ttk.Combobox(
            frame,
            textvariable=axis,
            values=[str(number) for number in range(1, MAX_TRENDS + 1)],
            state="readonly",
            width=5,
        ).pack(side="left", padx=1)
        ttk.Entry(frame, textvariable=minimum, width=10).pack(side="left", padx=1)
        ttk.Entry(frame, textvariable=maximum, width=10).pack(side="left", padx=1)
        color_button = tk.Button(
            frame,
            text=" ",
            width=3,
            background=color.get(),
            command=lambda: self.choose_color(color, color_button),
        )
        color_button.pack(side="left", padx=3)
        ttk.Entry(frame, textvariable=unit, width=12).pack(side="left", padx=1)
        ttk.Entry(frame, textvariable=number_format, width=22).pack(
            side="left", padx=1
        )
        row = TrendRow(
            row_id,
            frame,
            visible,
            column,
            axis,
            minimum,
            maximum,
            color,
            unit,
            number_format,
            column_combo,
            color_button,
        )
        ttk.Button(
            frame, text="Remove", command=lambda trend=row: self.remove_trend(trend)
        ).pack(side="left", padx=(4, 0))
        self.trend_rows.append(row)
        self._choose_default_trend_column(row)

    @staticmethod
    def _default_color(index):
        return to_hex(matplotlib.colormaps["tab10"]((index - 1) % 10 / 10))

    def _choose_default_trend_column(self, row):
        available = self._numeric_columns()
        if available and row.column.get() not in available:
            index = (row.row_id - 1) % len(available)
            row.column.set(available[index])

    def _numeric_columns(self):
        if self._source_frame is None:
            return self.columns
        numeric = []
        for column in self.columns:
            series = self._source_frame[column]
            if pd.api.types.is_datetime64_any_dtype(series):
                continue
            if series.dtype == object and any(
                isinstance(value, (pd.Timestamp, np.datetime64, datetime, date))
                for value in series.dropna().head(10)
            ):
                continue
            values = pd.to_numeric(series, errors="coerce")
            if values.notna().any():
                numeric.append(column)
        return numeric

    def remove_trend(self, trend):
        if trend in self.trend_rows:
            trend.frame.destroy()
            self.trend_rows.remove(trend)

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
        if self.time_column.get() not in self.columns:
            self.time_column.set((time_columns or self.columns)[0])
        numeric = self._numeric_columns()
        for row in self.trend_rows:
            row.column_combo.configure(values=numeric)
            self._choose_default_trend_column(row)

    def choose_color(self, variable, button):
        selected = colorchooser.askcolor(
            color=variable.get(), parent=self.parent.winfo_toplevel()
        )[1]
        if selected:
            variable.set(selected)
            button.configure(background=selected)

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
        times = _to_naive_datetimes(frame[time_name])
        if not times.notna().any():
            raise ValueError("The selected time column contains no valid date/time values.")

        if self.color_mode.get() == "Colormap":
            cmap = matplotlib.colormaps[self.colormap.get()]
            colors = {
                row.row_id: to_hex(cmap(index / max(1, len(self.trend_rows) - 1)))
                for index, row in enumerate(self.trend_rows)
            }
        else:
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

            axis_number = int(row.axis.get())
            minimum = self._parse_axis_limit(
                row.minimum.get(), f"Trend '{column_name}' minimum"
            )
            maximum = self._parse_axis_limit(
                row.maximum.get(), f"Trend '{column_name}' maximum"
            )
            if minimum is not None and maximum is not None and minimum >= maximum:
                raise ValueError(
                    f"Trend '{column_name}' minimum must be less than its maximum."
                )
            number_format = row.number_format.get().strip() or DEFAULT_NUMBER_FORMAT
            config = (minimum, maximum, row.unit.get().strip(), number_format)
            if axis_number in configs and configs[axis_number] != config:
                raise ValueError(
                    f"Trends assigned to Y-axis {axis_number} must share the same "
                    "range, units, and number format."
                )
            configs[axis_number] = config

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
                    "axis": axis_number,
                    "minimum": minimum,
                    "maximum": maximum,
                    "unit": config[2],
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
            axis_numbers = sorted(configs)
            primary_number = axis_numbers[0]

            figure = Figure(figsize=(max(11, 10 + 0.6 * (len(axis_numbers) - 1)), 7))
            base_axes = figure.add_subplot(111)
            base_axes.spines["right"].set_visible(False)
            axes = {primary_number: base_axes}
            for index, axis_number in enumerate(axis_numbers[1:], start=1):
                axis = base_axes.twinx()
                axis.spines["right"].set_position(("outward", 48 * index))
                axis.spines["left"].set_visible(False)
                axis.patch.set_visible(False)
                axes[axis_number] = axis

            handles = []
            labels = []
            plotted_series = []
            for series in prepared:
                axis = axes[series["axis"]]
                line, = axis.plot(
                    mdates.num2date(series["plot_times"]),
                    series["plot_values"],
                    color=series["color"],
                    label=series["column"],
                    linewidth=1.4,
                )
                handles.append(line)
                labels.append(series["column"])
                plotted_series.append(series)

            for axis_number, axis in axes.items():
                axis_trends = [
                    series for series in prepared if series["axis"] == axis_number
                ]
                minimum, maximum, unit, number_format = configs[axis_number]
                color = axis_trends[0]["color"]
                label = ", ".join(series["column"] for series in axis_trends)
                if unit:
                    label += f" ({unit})"
                if axis is base_axes:
                    axis.spines["left"].set_color(color)
                    axis.yaxis.tick_left()
                    axis.yaxis.set_label_position("left")
                else:
                    axis.spines["right"].set_color(color)
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

            base_axes.set_xlim(display_min, display_max)
            base_axes.xaxis.set_major_locator(mdates.AutoDateLocator())
            base_axes.xaxis.set_major_formatter(mdates.ConciseDateFormatter(base_axes.xaxis.get_major_locator()))
            base_axes.grid(True, axis="both", alpha=0.25)
            base_axes.set_title(self.title.get().strip() or "Trend navigator")
            figure.legend(
                handles,
                labels,
                loc="upper center",
                bbox_to_anchor=(0.5, 0.98),
                ncol=min(4, max(1, len(labels))),
            )
            right_margin = 1.6 + 0.55 * (len(axis_numbers) - 1)
            figure_width = figure.get_size_inches()[0]
            figure.subplots_adjust(
                left=0.1,
                right=max(0.45, min(0.82, 1 - right_margin / figure_width)),
                top=0.83,
                bottom=0.16,
            )

            if self.canvas is not None:
                self.canvas.get_tk_widget().destroy()
            self.figure = figure
            self.base_axes = base_axes
            self.axes_by_number = axes
            self.plotted_series = plotted_series
            self.data_time_range = (data_min, data_max)
            self.canvas = FigureCanvasTkAgg(figure, master=self.graph_frame)
            self.canvas.draw()
            self.canvas.get_tk_widget().pack(fill="both", expand=True)
            self.span_selector = SpanSelector(
                base_axes,
                self._on_span_select,
                "horizontal",
                useblit=True,
                props={"alpha": 0.18, "facecolor": "tab:blue"},
                interactive=True,
                drag_from_anywhere=True,
            )
            self._redraw_hairlines()
            self.set_status(
                f"Plotted {len(plotted_series)} trends over "
                f"{len(np.unique(np.concatenate([s['times'] for s in plotted_series])))} "
                "timestamps."
            )
        except (ValueError, TypeError, KeyError, IndexError, OverflowError) as exc:
            self.set_status(str(exc), error=True)

    def _on_span_select(self, start, end):
        if start is None or end is None or start >= end:
            return
        self.start_time.set(
            mdates.num2date(start).replace(tzinfo=None).strftime(TIME_DISPLAY_FORMAT)
        )
        self.end_time.set(
            mdates.num2date(end).replace(tzinfo=None).strftime(TIME_DISPLAY_FORMAT)
        )
        self.base_axes.set_xlim(start, end)
        self.canvas.draw_idle()
        self._update_hairline_values()

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
            self.base_axes.set_xlim(lower, upper)
            self.canvas.draw_idle()
            self._update_hairline_values()
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
            self.canvas.draw_idle()
            self._update_hairline_values()

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
        position = mdates.date2num(timestamp.to_pydatetime())
        lower, upper = self.base_axes.get_xlim()
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
        if hairline is None:
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
        if self.base_axes is None:
            return
        for hairline in self.hairlines:
            if hairline.artist is not None:
                hairline.artist.remove()
            timestamp = parse_time(hairline.time.get(), "Hairline time")
            hairline.artist = self.base_axes.axvline(
                timestamp,
                color=hairline.color,
                linestyle="--",
                linewidth=1.3,
                label=f"Hairline {hairline.hairline_id}",
            )
        if self.canvas is not None:
            self.canvas.draw_idle()
        self._update_hairline_values()

    def _update_hairline_values(self):
        self.hairline_values.delete(*self.hairline_values.get_children())
        for hairline in self.hairlines:
            timestamp = parse_time(hairline.time.get(), "Hairline time")
            position = mdates.date2num(timestamp.to_pydatetime())
            for series in self.plotted_series:
                times = series["times"]
                if position < times[0] or position > times[-1]:
                    value_text = "—"
                else:
                    value = float(np.interp(position, times, series["values"]))
                    value_text = format_excel_number(value, series["number_format"])
                self.hairline_values.insert(
                    "",
                    "end",
                    values=(
                        f"Hairline {hairline.hairline_id}",
                        timestamp.strftime(TIME_DISPLAY_FORMAT),
                        series["column"],
                        value_text,
                    ),
                )
