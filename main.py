import copy
import io
import json
import tkinter as tk
from tkinter import colorchooser, filedialog, messagebox, ttk

import matplotlib

matplotlib.use("TkAgg")

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import pywintypes
import xlwings as xw
import win32clipboard
import win32con
from PIL import Image
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.colors import is_color_like
from matplotlib.figure import Figure
from trend_navigator import TrendNavigator


GRAPH_EXPORT_DPI = 300
SETTINGS_SHEET_NAME = "HeatMapSettings"
TREND_SETTINGS_SHEET_NAME = "TrendNavigatorSettings"
TREND_SETTINGS_FIELDS = [
    "Settings name",
    "Source sheet",
    "Excel table",
    "Source range",
    "Trend settings JSON",
]
SETTINGS_FIELDS = [
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

version_id = "1.0a (last updated 19/11/2026)"


def figure_to_dib(figure, figure_size=None):
    png = io.BytesIO()
    previous_size = figure.get_size_inches().copy()
    try:
        if figure_size is not None:
            figure.set_size_inches(*figure_size, forward=False)
        figure.savefig(
            png,
            format="png",
            dpi=GRAPH_EXPORT_DPI,
            bbox_inches=None if figure_size is not None else "tight",
        )
    finally:
        if figure_size is not None:
            figure.set_size_inches(*previous_size, forward=False)
    png.seek(0)
    with Image.open(png) as image:
        bitmap = io.BytesIO()
        image.convert("RGB").save(bitmap, format="BMP")
    return bitmap.getvalue()[14:]


def make_bin_edges(specification, values, lower, upper, axis_label):
    text = specification.strip()
    if not text:
        raise ValueError(f"Enter a bin width for {axis_label}.")
    try:
        width = float(text)
    except ValueError as exc:
        raise ValueError(f"{axis_label} bin width must be a finite positive number.") from exc
    if not np.isfinite(width) or width <= 0:
        raise ValueError(f"{axis_label} bin width must be a finite positive number.")

    data_min = float(np.min(values))
    data_max = float(np.max(values))
    display_min = data_min if lower is None else lower
    display_max = data_max if upper is None else upper
    if display_min == display_max:
        if lower is None and upper is None:
            display_min -= width / 2
            display_max += width / 2
        elif lower is not None and upper is None and lower == data_max:
            display_max += width
        elif upper is not None and lower is None and upper == data_min:
            display_min -= width
        else:
            raise ValueError(f"{axis_label} displayed range must have positive width.")
    if display_min > display_max:
        raise ValueError(f"{axis_label} displayed minimum must be less than its maximum.")

    bin_count = max(
        1, int(np.ceil(np.nextafter((display_max - display_min) / width, -np.inf)))
    )
    edges = display_min + np.arange(bin_count + 1, dtype=float) * width
    edges[-1] = max(edges[-1], display_max)
    return edges, display_min, display_max


def parse_optional_limit(text, label):
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


def validate_limits(lower, upper, label):
    if lower is not None and upper is not None and lower >= upper:
        raise ValueError(f"{label} minimum must be less than its maximum.")


class HeatMapApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Julien's Excel MatplotLib Plotter " + version_id)
        self.root.geometry("1100x900")
        self.workbook_entries = {}
        self.saved_settings = {}
        self.saved_trend_settings = {}
        self.figure = None
        self.figure_canvas = None
        self.figure_fixed_size = False

        self._build_window()
        self.refresh_workbooks()

    def _build_window(self):
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.scroll_canvas = tk.Canvas(self.root, highlightthickness=0)
        scrollbar = ttk.Scrollbar(
            self.root, orient="vertical", command=self.scroll_canvas.yview
        )
        horizontal_scrollbar = ttk.Scrollbar(
            self.root, orient="horizontal", command=self.scroll_canvas.xview
        )
        self.scroll_canvas.configure(
            xscrollcommand=horizontal_scrollbar.set, yscrollcommand=scrollbar.set
        )
        self.scroll_canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        horizontal_scrollbar.grid(row=1, column=0, sticky="ew")
        self.root.rowconfigure(1, weight=0)

        self.content = ttk.Frame(self.scroll_canvas, padding=12)
        self.content_window = self.scroll_canvas.create_window(
            (0, 0), window=self.content, anchor="nw"
        )
        self.content.bind(
            "<Configure>",
            lambda _event: self.scroll_canvas.configure(
                scrollregion=self.scroll_canvas.bbox("all")
            ),
        )
        self.scroll_canvas.bind(
            "<Configure>",
            self._resize_content,
        )
        self.scroll_canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.scroll_canvas.bind_all("<Button-4>", self._on_mousewheel)
        self.scroll_canvas.bind_all("<Button-5>", self._on_mousewheel)

        self.workbook_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.table_var = tk.StringVar()
        self.source_range_var = tk.StringVar()
        self.x_var = tk.StringVar()
        self.y_var = tk.StringVar()
        self.weight_var = tk.StringVar()
        self.x_bins_var = tk.StringVar(value="1")
        self.y_bins_var = tk.StringVar(value="1")
        self.aggregate_var = tk.StringVar(value="Bin count")
        self.plot_title_var = tk.StringVar(value="Bin count heat map")
        self.x_title_var = tk.StringVar()
        self.y_title_var = tk.StringVar()
        self.cmap_var = tk.StringVar(value="viridis")
        self.reverse_cmap_var = tk.BooleanVar(value=False)
        self.x_min_var = tk.StringVar()
        self.x_max_var = tk.StringVar()
        self.y_min_var = tk.StringVar()
        self.y_max_var = tk.StringVar()
        self.legend_min_var = tk.StringVar()
        self.legend_max_var = tk.StringVar()
        self.under_color_var = tk.StringVar(value="#440154")
        self.over_color_var = tk.StringVar(value="#fde725")
        self.status_var = tk.StringVar()
        self.settings_name_var = tk.StringVar()
        self.saved_settings_var = tk.StringVar()
        self.figure_width_cm_var = tk.StringVar()
        self.figure_height_cm_var = tk.StringVar()

        self._build_controls()

    def _resize_content(self, event):
        width = max(event.width, self.content.winfo_reqwidth())
        self.scroll_canvas.itemconfigure(self.content_window, width=width)

    def _on_mousewheel(self, event):
        widget = event.widget
        widget_path = str(widget)
        if ".popdown" in widget_path or widget.winfo_class() == "Listbox":
            return "break"
        if widget.winfo_class() in ("TCombobox", "Combobox"):
            return "break"

        if self.trend_navigator.canvas is not None:
            graph_widget = str(self.trend_navigator.canvas.get_tk_widget())
            if widget_path == graph_widget or widget_path.startswith(graph_widget + "."):
                return "break"

        if event.num == 4:
            direction = 1
        elif event.num == 5:
            direction = -1
        else:
            direction = -1 if event.delta > 0 else 1
        self.scroll_canvas.yview_scroll(direction, "units")
        return "break"

    @staticmethod
    def _protect_combobox_wheel(combo):
        for sequence in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            combo.bind(sequence, lambda _event: "break")

    def _section(self, text, parent=None):
        if parent is None:
            parent = self.content
        ttk.Label(parent, text=text, font=("", 10, "bold")).pack(
            anchor="w", pady=(10, 4)
        )

    def _entry_row(self, label, variable, placeholder=None, width=36, parent=None):
        if parent is None:
            parent = self.content
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=2)
        ttk.Label(row, text=label, width=22).pack(side="left")
        entry = ttk.Entry(row, textvariable=variable, width=width)
        entry.pack(side="left", fill="x", expand=True)
        if placeholder:
            ttk.Label(row, text=placeholder, foreground="#666").pack(
                side="left", padx=(8, 0)
            )
        return entry

    def _combo_row(self, label, variable, values=(), width=48, parent=None):
        if parent is None:
            parent = self.content
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=2)
        ttk.Label(row, text=label, width=22).pack(side="left")
        combo = ttk.Combobox(
            row, textvariable=variable, values=list(values), width=width, state="readonly"
        )
        combo.pack(side="left", fill="x", expand=True)
        self._protect_combobox_wheel(combo)
        return combo

    def _build_controls(self):
        self._section("Shared Excel data source")
        row = ttk.Frame(self.content)
        row.pack(fill="x", pady=2)
        ttk.Label(row, text="Workbook:", width=22).pack(side="left")
        self.workbook_combo = ttk.Combobox(
            row, textvariable=self.workbook_var, state="readonly"
        )
        self.workbook_combo.pack(side="left", fill="x", expand=True)
        self._protect_combobox_wheel(self.workbook_combo)
        self.workbook_combo.bind(
            "<<ComboboxSelected>>", lambda *_args: self.on_workbook_change()
        )
        ttk.Button(row, text="Refresh workbooks", command=self.refresh_workbooks).pack(
            side="left", padx=(8, 0)
        )
        self.sheet_combo = self._combo_row("Sheet:", self.sheet_var)
        self.sheet_combo.bind(
            "<<ComboboxSelected>>", lambda *_args: self.on_sheet_change()
        )
        self.table_combo = self._combo_row("Excel table:", self.table_var)
        self.table_combo.bind(
            "<<ComboboxSelected>>", lambda *_args: self.on_table_change()
        )
        self._entry_row("Source range:", self.source_range_var, "e.g. A1:C500")
        ttk.Button(self.content, text="Load columns", command=self.load_columns).pack(
            anchor="w", pady=(4, 8)
        )

        self.status_label = ttk.Label(
            self.content, textvariable=self.status_var, wraplength=950
        )
        self.status_label.pack(anchor="w", pady=(0, 6))

        dimensions = ttk.Frame(self.content)
        dimensions.pack(fill="x", pady=(0, 6))
        ttk.Label(
            dimensions, text="Graph size (cm; leave both blank for automatic):"
        ).pack(side="left")
        ttk.Label(dimensions, text="Width").pack(side="left", padx=(12, 3))
        ttk.Entry(dimensions, textvariable=self.figure_width_cm_var, width=8).pack(
            side="left"
        )
        ttk.Label(dimensions, text="Height").pack(side="left", padx=(10, 3))
        ttk.Entry(dimensions, textvariable=self.figure_height_cm_var, width=8).pack(
            side="left"
        )

        self.graph_tabs = ttk.Notebook(self.content)
        self.graph_tabs.pack(fill="both", expand=True)
        self.heatmap_tab = ttk.Frame(self.graph_tabs, padding=8)
        self.trend_tab = ttk.Frame(self.graph_tabs, padding=8)
        self.graph_tabs.add(self.heatmap_tab, text="Heat map")
        self.graph_tabs.add(self.trend_tab, text="Trend navigator")
        self._build_heatmap_controls(self.heatmap_tab)
        self.trend_navigator = TrendNavigator(
            self.trend_tab,
            data_provider=self.get_source_frame,
            save_figure=lambda figure, title: self.save_graph(
                figure, title, use_heatmap=False
            ),
            copy_figure=lambda figure: self.copy_graph_to_clipboard(
                figure, use_heatmap=False
            ),
            set_status=self.set_status,
            figure_size_inches=self.get_figure_size_inches,
            save_settings=self.save_trend_settings,
            refresh_settings=self.refresh_saved_trend_settings,
            load_settings=self.load_trend_settings,
        )

    def get_figure_size_inches(self):
        width_text = self.figure_width_cm_var.get().strip()
        height_text = self.figure_height_cm_var.get().strip()
        if not width_text and not height_text:
            return None
        if not width_text or not height_text:
            raise ValueError("Enter both graph width and height in centimeters.")
        try:
            width_cm = float(width_text)
            height_cm = float(height_text)
        except ValueError as exc:
            raise ValueError("Graph width and height must be positive numbers.") from exc
        if not np.isfinite(width_cm) or not np.isfinite(height_cm):
            raise ValueError("Graph width and height must be positive numbers.")
        if width_cm <= 0 or height_cm <= 0:
            raise ValueError("Graph width and height must be positive numbers.")
        return width_cm / 2.54, height_cm / 2.54

    def _build_heatmap_controls(self, parent):
        settings_columns = ttk.Frame(parent)
        settings_columns.pack(fill="x")
        settings_columns.columnconfigure((0, 1), weight=1, uniform="heatmap")
        data_settings = ttk.LabelFrame(
            settings_columns, text="Data and labels", padding=8
        )
        data_settings.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        color_settings = ttk.LabelFrame(
            settings_columns, text="Weights colormap", padding=8
        )
        color_settings.grid(row=0, column=1, sticky="nsew", padx=(5, 0))

        self.x_combo = self._combo_row("X column:", self.x_var, parent=data_settings)
        self.y_combo = self._combo_row("Y column:", self.y_var, parent=data_settings)
        self.weight_combo = self._combo_row(
            "Weight:", self.weight_var, parent=data_settings
        )
        self._entry_row(
            "X bin width:", self.x_bins_var, "X-axis units", parent=data_settings
        )
        self._entry_row(
            "Y bin height:", self.y_bins_var, "Y-axis units", parent=data_settings
        )
        self._combo_row(
            "Aggregate:",
            self.aggregate_var,
            ["Bin count", "Sum of weights", "Average of weights"],
            parent=data_settings,
        )
        self._entry_row("Graph title:", self.plot_title_var, parent=data_settings)
        self._entry_row(
            "X-axis title:", self.x_title_var, "blank = column name", parent=data_settings
        )
        self._entry_row(
            "Y-axis title:", self.y_title_var, "blank = column name", parent=data_settings
        )

        row = ttk.Frame(color_settings)
        row.pack(fill="x", pady=2)
        ttk.Label(row, text="Colormap:", width=22).pack(side="left")
        self.cmap_combo = ttk.Combobox(
            row,
            textvariable=self.cmap_var,
            values=sorted(plt.colormaps()),
            state="readonly",
        )
        self.cmap_combo.pack(side="left", fill="x", expand=True)
        self._protect_combobox_wheel(self.cmap_combo)
        self.cmap_combo.bind(
            "<<ComboboxSelected>>", lambda *_args: self.update_cmap_preview()
        )
        ttk.Checkbutton(
            color_settings,
            text="Reverse gradient",
            variable=self.reverse_cmap_var,
            command=self.update_cmap_preview,
        ).pack(anchor="w", pady=4)
        self.cmap_preview = tk.Canvas(
            color_settings, height=18, width=400, highlightthickness=0
        )
        self.cmap_preview.pack(fill="x", pady=3)
        self.update_cmap_preview()

        limits = ttk.Frame(parent)
        limits.pack(fill="x", pady=(6, 0))
        limits.columnconfigure((0, 1), weight=1, uniform="heatmap-limits")
        axes_group = ttk.LabelFrame(
            limits, text="Axis ranges (blank = automatic)", padding=8
        )
        axes_group.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        legend_group = ttk.LabelFrame(
            limits, text="Color scale range and clipping colors", padding=8
        )
        legend_group.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        for title, variable, companion_title, companion, group in (
            ("X minimum:", self.x_min_var, "X maximum:", self.x_max_var, axes_group),
            ("Y minimum:", self.y_min_var, "Y maximum:", self.y_max_var, axes_group),
            (
                "Color minimum:",
                self.legend_min_var,
                "Color maximum:",
                self.legend_max_var,
                legend_group,
            ),
        ):
            pair = ttk.Frame(group)
            pair.pack(fill="x", pady=2)
            ttk.Label(pair, text=title, width=14).pack(side="left")
            ttk.Entry(pair, textvariable=variable, width=12).pack(side="left")
            ttk.Label(pair, text=companion_title, width=15).pack(
                side="left", padx=(8, 0)
            )
            ttk.Entry(pair, textvariable=companion, width=12).pack(side="left")
        colors = ttk.Frame(legend_group)
        colors.pack(fill="x", pady=2)
        colors.columnconfigure((0, 1), weight=1, uniform="range-colors")
        for column, (label, variable) in enumerate(
            (
                ("Below range:", self.under_color_var),
                ("Above range:", self.over_color_var),
            )
        ):
            color_control = ttk.Frame(colors)
            color_control.grid(row=0, column=column, sticky="ew", padx=(0, 6))
            ttk.Label(color_control, text=label, width=11).pack(side="left")
            ttk.Entry(color_control, textvariable=variable, width=9).pack(
                side="left", fill="x", expand=True
            )
            ttk.Button(
                color_control,
                text="Choose",
                command=lambda selected=variable: self.choose_color(selected),
            ).pack(side="left", padx=(4, 0))

        self._section("Saved heat-map settings for this workbook", parent)
        self._entry_row("Settings name:", self.settings_name_var, parent=parent)
        self.saved_settings_combo = self._combo_row(
            "Saved settings:", self.saved_settings_var, parent=parent
        )
        settings_buttons = ttk.Frame(parent)
        settings_buttons.pack(anchor="w", pady=4)
        ttk.Button(
            settings_buttons, text="Save settings", command=self.save_settings
        ).pack(side="left", padx=(0, 6))
        ttk.Button(
            settings_buttons,
            text="Refresh saved settings",
            command=self.refresh_saved_settings,
        ).pack(side="left", padx=(0, 6))
        ttk.Button(
            settings_buttons, text="Import settings", command=self.load_settings
        ).pack(side="left")

        ttk.Button(
            parent, text="Plot heat map", command=self.plot_heatmap
        ).pack(anchor="w", pady=(10, 4))
        ttk.Separator(parent).pack(fill="x", pady=(12, 6))
        ttk.Label(parent, text="Generated heat map", font=("", 12, "bold")).pack(
            anchor="w", pady=(0, 6)
        )
        export_buttons = ttk.Frame(parent)
        export_buttons.pack(anchor="w", pady=(0, 6))
        ttk.Button(
            export_buttons, text="Save graph", command=self.save_graph
        ).pack(side="left", padx=(0, 6))
        ttk.Button(
            export_buttons,
            text="Copy graph to clipboard",
            command=self.copy_graph_to_clipboard,
        ).pack(side="left")
        self.plot_frame = ttk.Frame(parent)
        self.plot_frame.pack(fill="both", expand=True)

    def choose_color(self, variable):
        selected = colorchooser.askcolor(color=variable.get(), parent=self.root)[1]
        if selected:
            variable.set(selected)

    def set_status(self, message, error=False):
        self.status_var.set(message)
        self.status_label.configure(foreground="#b00020" if error else "#222")

    def open_workbook_entries(self):
        entries = {}
        for app in xw.apps:
            for book in app.books:
                label = f"{book.name} (Excel PID {app.pid})"
                entries[label] = (app.pid, book.name)
        return entries

    def selected_book(self):
        label = self.workbook_var.get()
        if label not in self.workbook_entries:
            raise ValueError(
                "Select an open workbook, then refresh the workbook list if needed."
            )
        pid, name = self.workbook_entries[label]
        return xw.apps[pid].books[name]

    def settings_sheet(self, book, create=False):
        for sheet in book.sheets:
            if sheet.name.casefold() == SETTINGS_SHEET_NAME.casefold():
                headers = (
                    sheet.range("A1")
                    .resize(1, len(SETTINGS_FIELDS))
                    .options(ndim=2)
                    .value
                )
                if headers is None or headers[0] != SETTINGS_FIELDS:
                    raise ValueError(
                        f"The '{SETTINGS_SHEET_NAME}' sheet exists but does not "
                        "contain a compatible heat-map settings table."
                    )
                return sheet
        if not create:
            return None
        sheet = book.sheets.add(SETTINGS_SHEET_NAME, after=book.sheets[-1])
        sheet.range("A1").value = [SETTINGS_FIELDS]
        return sheet

    def read_saved_settings(self, sheet):
        last_row = sheet.used_range.last_cell.row
        if last_row < 2:
            return {}
        rows = (
            sheet.range((2, 1), (last_row, len(SETTINGS_FIELDS)))
            .options(ndim=2)
            .value
        )
        if rows is None:
            return {}
        settings = {}
        seen_names = set()
        for row_number, row in enumerate(rows, start=2):
            if row and row[0] is not None and str(row[0]).strip():
                name = str(row[0]).strip()
                folded_name = name.casefold()
                if folded_name in seen_names:
                    raise ValueError(
                        f"Duplicate settings name '{name}' found in "
                        f"'{SETTINGS_SHEET_NAME}'."
                    )
                seen_names.add(folded_name)
                settings[name] = (row_number, dict(zip(SETTINGS_FIELDS, row)))
        return settings

    def trend_settings_sheet(self, book, create=False):
        for sheet in book.sheets:
            if sheet.name.casefold() == TREND_SETTINGS_SHEET_NAME.casefold():
                headers = (
                    sheet.range("A1")
                    .resize(1, len(TREND_SETTINGS_FIELDS))
                    .options(ndim=2)
                    .value
                )
                if headers is None or headers[0] != TREND_SETTINGS_FIELDS:
                    raise ValueError(
                        f"The '{TREND_SETTINGS_SHEET_NAME}' sheet exists but does not "
                        "contain a compatible trend settings table."
                    )
                return sheet
        if not create:
            return None
        sheet = book.sheets.add(
            TREND_SETTINGS_SHEET_NAME, after=book.sheets[-1]
        )
        sheet.range("A1").value = [TREND_SETTINGS_FIELDS]
        return sheet

    def read_saved_trend_settings(self, sheet):
        last_row = sheet.used_range.last_cell.row
        if last_row < 2:
            return {}
        rows = (
            sheet.range((2, 1), (last_row, len(TREND_SETTINGS_FIELDS)))
            .options(ndim=2)
            .value
        )
        if rows is None:
            return {}
        settings = {}
        seen_names = set()
        for row_number, row in enumerate(rows, start=2):
            if row and row[0] is not None and str(row[0]).strip():
                name = str(row[0]).strip()
                folded_name = name.casefold()
                if folded_name in seen_names:
                    raise ValueError(
                        f"Duplicate trend settings name '{name}' found in "
                        f"'{TREND_SETTINGS_SHEET_NAME}'."
                    )
                seen_names.add(folded_name)
                settings[name] = (
                    row_number,
                    dict(zip(TREND_SETTINGS_FIELDS, row)),
                )
        return settings

    def refresh_saved_trend_settings(self):
        try:
            sheet = self.trend_settings_sheet(self.selected_book())
            self.saved_trend_settings = (
                {}
                if sheet is None
                else self.read_saved_trend_settings(sheet)
            )
            names = sorted(self.saved_trend_settings, key=str.casefold)
            self.trend_navigator.set_saved_settings_names(names)
            self.set_status(
                f"Found {len(names)} saved trend settings."
                if names
                else "No saved trend settings found in this workbook."
            )
            return True
        except (
            ValueError,
            TypeError,
            KeyError,
            IndexError,
            pywintypes.com_error,
        ) as exc:
            self.saved_trend_settings = {}
            self.trend_navigator.set_saved_settings_names([])
            self.set_status(str(exc), error=True)
            return False

    def save_trend_settings(self, name, settings):
        try:
            self._write_trend_settings(name, settings)
        except (
            ValueError,
            TypeError,
            KeyError,
            IndexError,
            tk.TclError,
            pywintypes.com_error,
        ) as exc:
            self.set_status(str(exc), error=True)

    def _write_trend_settings(self, name, settings):
        if not name:
            raise ValueError("Enter a name for this trend settings set.")
        if not self.sheet_var.get():
            raise ValueError("Select the source sheet before saving trend settings.")
        address = self.source_range_var.get().strip()
        if not address:
            raise ValueError("Choose a table or enter a source range first.")

        book = self.selected_book()
        source_sheet = book.sheets[self.sheet_var.get()]
        frame = self.read_source_frame(source_sheet, address)
        columns = self.source_column_names(frame)
        frame = frame.copy()
        frame.columns = columns
        self.trend_navigator.validate_settings(settings, columns, frame)
        serialized = json.dumps(settings, ensure_ascii=True, allow_nan=False)
        if len(serialized) > 32767:
            raise ValueError("Trend settings are too large to store in one Excel cell.")

        settings_sheet = self.trend_settings_sheet(book, create=True)
        saved = self.read_saved_trend_settings(settings_sheet)
        existing = next(
            (
                (saved_name, entry)
                for saved_name, entry in saved.items()
                if saved_name.casefold() == name.casefold()
            ),
            None,
        )
        if existing is not None:
            if not messagebox.askyesno(
                "Replace trend settings?",
                f"A trend settings set named '{existing[0]}' already exists. Replace it?",
                parent=self.root,
            ):
                return
            row_number = existing[1][0]
        else:
            row_number = settings_sheet.used_range.last_cell.row + 1

        settings_sheet.range((row_number, 1)).value = [
            [
                name,
                self.sheet_var.get(),
                self.table_var.get(),
                address,
                serialized,
            ]
        ]
        if self.refresh_saved_trend_settings():
            self.trend_navigator.saved_settings_name.set(name)
            self.trend_navigator.settings_name.set(name)
            self.set_status(f"Saved trend settings as '{name}'.")

    def load_trend_settings(self, name):
        try:
            if name not in self.saved_trend_settings:
                raise ValueError("Select a saved trend settings set to load.")
            record = self.saved_trend_settings[name][1]

            def text_value(field):
                value = record.get(field)
                return "" if value is None else str(value)

            book = self.selected_book()
            saved_sheet_name = text_value("Source sheet")
            source_sheets = {sheet.name.casefold(): sheet for sheet in book.sheets}
            source_sheet = source_sheets.get(saved_sheet_name.casefold())
            if source_sheet is None:
                raise ValueError(
                    f"The saved source sheet '{saved_sheet_name}' no longer exists."
                )

            table_name = text_value("Excel table")
            if table_name:
                tables = {
                    table.name.casefold(): table.name
                    for table in source_sheet.tables
                }
                if table_name.casefold() not in tables:
                    raise ValueError(
                        f"The saved Excel table '{table_name}' no longer exists "
                        f"on sheet '{source_sheet.name}'."
                    )
                table_name = tables[table_name.casefold()]

            address = text_value("Source range")
            if not address:
                raise ValueError("The saved trend settings do not include a source range.")
            frame = self.read_source_frame(source_sheet, address, table_name)
            columns = self.source_column_names(frame)
            frame = frame.copy()
            frame.columns = columns

            try:
                settings = json.loads(text_value("Trend settings JSON"))
            except json.JSONDecodeError as exc:
                raise ValueError("The saved trend settings contain invalid JSON.") from exc
            self.trend_navigator.validate_settings(settings, columns, frame)

            self.sheet_combo.configure(
                values=[sheet.name for sheet in book.sheets]
            )
            self.sheet_var.set(source_sheet.name)
            self.table_combo.configure(
                values=[table.name for table in source_sheet.tables]
            )
            self.table_var.set(table_name)
            self.source_range_var.set(address)
            self.x_combo.configure(values=columns)
            self.y_combo.configure(values=columns)
            self.weight_combo.configure(values=["(none)"] + columns)
            if self.x_var.get() not in columns:
                self.x_var.set(columns[0])
            if self.y_var.get() not in columns:
                self.y_var.set(columns[1] if len(columns) > 1 else columns[0])
            if self.weight_var.get() not in ["(none)"] + columns:
                self.weight_var.set("(none)")
            self.trend_navigator.set_columns(columns, frame)
            self.trend_navigator.apply_settings(settings)
            self.trend_navigator.settings_name.set(name)
            self.set_status(
                f"Loaded trend settings '{name}' and restored its source and graph options."
            )
        except (
            ValueError,
            TypeError,
            KeyError,
            IndexError,
            tk.TclError,
            pywintypes.com_error,
        ) as exc:
            self.set_status(str(exc), error=True)

    def refresh_saved_settings(self):
        try:
            sheet = self.settings_sheet(self.selected_book())
            self.saved_settings = (
                {} if sheet is None else self.read_saved_settings(sheet)
            )
            names = sorted(self.saved_settings, key=str.casefold)
            self.saved_settings_combo.configure(values=names)
            if self.saved_settings_var.get() not in self.saved_settings:
                self.saved_settings_var.set(names[0] if names else "")
            if names:
                self.set_status(f"Found {len(names)} saved graph settings.")
            else:
                self.set_status(
                    f"No saved graph settings found in this workbook. "
                    f"They will be stored on the '{SETTINGS_SHEET_NAME}' sheet."
                )
            return True
        except (ValueError, TypeError, KeyError, IndexError, pywintypes.com_error) as exc:
            self.saved_settings = {}
            self.saved_settings_combo.configure(values=[])
            self.saved_settings_var.set("")
            self.set_status(str(exc), error=True)
            return False

    def refresh_workbooks(self):
        try:
            self.trend_navigator.set_columns([])
            self.workbook_entries = self.open_workbook_entries()
            labels = list(self.workbook_entries)
            self.workbook_combo.configure(values=labels)
            if labels:
                self.workbook_var.set(labels[0])
                self.refresh_sheets()
                self.set_status(
                    "Open workbooks found. Select a source and load its columns."
                )
                self.refresh_saved_settings()
                self.refresh_saved_trend_settings()
            else:
                self.workbook_var.set("")
                self.sheet_combo.configure(values=[])
                self.table_combo.configure(values=[])
                self.saved_settings_combo.configure(values=[])
                self.sheet_var.set("")
                self.table_var.set("")
                self.saved_settings_var.set("")
                self.saved_settings = {}
                self.saved_trend_settings = {}
                self.trend_navigator.set_saved_settings_names([])
                self.set_status(
                    "No open Excel workbooks found. Open a workbook in desktop Excel, then refresh.",
                    error=True,
                )
        except (ValueError, TypeError, KeyError, IndexError) as exc:
            self.set_status(str(exc), error=True)

    def refresh_sheets(self):
        book = self.selected_book()
        sheets = [sheet.name for sheet in book.sheets]
        self.sheet_combo.configure(values=sheets)
        self.sheet_var.set(sheets[0] if sheets else "")
        self.refresh_tables()

    def refresh_tables(self):
        if not self.sheet_var.get():
            self.table_combo.configure(values=[])
            self.table_var.set("")
            return
        sheet = self.selected_book().sheets[self.sheet_var.get()]
        tables = [table.name for table in sheet.tables]
        self.table_combo.configure(values=tables)
        self.table_var.set("")

    def on_workbook_change(self):
        try:
            self.refresh_sheets()
            self.refresh_saved_settings()
            self.refresh_saved_trend_settings()
            self.trend_navigator.set_columns([])
        except (ValueError, TypeError, KeyError, IndexError) as exc:
            self.set_status(str(exc), error=True)

    def on_sheet_change(self):
        try:
            self.refresh_tables()
            self.trend_navigator.set_columns([])
        except (ValueError, TypeError, KeyError, IndexError) as exc:
            self.set_status(str(exc), error=True)

    def on_table_change(self):
        try:
            table = self.selected_book().sheets[self.sheet_var.get()].tables[
                self.table_var.get()
            ]
            self.source_range_var.set(table.range.address)
            self.trend_navigator.set_columns([])
            self.set_status(
                "Table range selected. Load columns to refresh the X/Y/weight choices."
            )
        except (ValueError, TypeError, KeyError, IndexError) as exc:
            self.set_status(str(exc), error=True)

    def read_source_frame(self, sheet, address, table_name=None):
        if table_name is None:
            table_name = self.table_var.get()
        if table_name:
            table = sheet.tables[table_name]
            if address == table.range.address:
                header_range = table.header_row_range
                if header_range is None:
                    raise ValueError(
                        "The selected Excel table does not have a visible header row."
                    )
                headers = header_range.options(ndim=2).value[0]
                body_range = table.data_body_range
                rows = [] if body_range is None else body_range.options(ndim=2).value
                return pd.DataFrame(rows, columns=headers)
        return sheet.range(address).options(
            pd.DataFrame, header=1, index=False
        ).value

    def get_source_frame(self):
        if not self.sheet_var.get():
            raise ValueError("Select an open workbook and sheet first.")
        address = self.source_range_var.get().strip()
        if not address:
            raise ValueError("Choose a table or enter a source range first.")
        frame = self.read_source_frame(
            self.selected_book().sheets[self.sheet_var.get()], address
        )
        if frame is None or frame.empty:
            raise ValueError("The selected source contains no data rows.")
        return frame

    def source_column_names(self, frame):
        if frame is None or frame.empty or len(frame.columns) == 0:
            raise ValueError(
                "The selected source contains no data columns. Check the range and header row."
            )
        if not frame.columns.is_unique:
            raise ValueError(
                "The source has duplicate column headers. Rename them in Excel before loading."
            )
        columns = [str(column) for column in frame.columns]
        if len(set(columns)) != len(columns):
            raise ValueError(
                "The source has column headers that become identical when displayed. "
                "Rename them in Excel."
            )
        return columns

    def load_columns(self):
        try:
            if not self.sheet_var.get():
                raise ValueError("Select an open workbook and sheet first.")
            address = self.source_range_var.get().strip()
            if not address:
                raise ValueError("Choose a table or enter an A1 source range with a header row.")
            sheet = self.selected_book().sheets[self.sheet_var.get()]
            frame = self.read_source_frame(sheet, address)
            columns = self.source_column_names(frame)
            self.trend_navigator.set_columns(columns, frame)
            self.x_combo.configure(values=columns)
            self.y_combo.configure(values=columns)
            self.weight_combo.configure(values=["(none)"] + columns)
            self.x_var.set(columns[0] if columns else "")
            self.y_var.set(columns[1] if len(columns) >= 2 else "")
            self.weight_var.set(columns[2] if len(columns) >= 3 else "(none)")
            self.set_status(
                f"Loaded {len(frame)} data rows and {len(columns)} columns."
            )
        except (ValueError, TypeError, KeyError, IndexError) as exc:
            self.set_status(str(exc), error=True)

    def save_settings(self):
        try:
            name = self.settings_name_var.get().strip()
            if not name:
                raise ValueError("Enter a name for this settings set.")
            if not self.sheet_var.get():
                raise ValueError("Select the source sheet before saving settings.")
            address = self.source_range_var.get().strip()
            if not address:
                raise ValueError("Choose a table or enter a source range first.")
            if not self.x_var.get() or not self.y_var.get():
                raise ValueError("Load the source columns and select both X and Y columns.")

            book = self.selected_book()
            source_sheet = book.sheets[self.sheet_var.get()]
            frame = self.read_source_frame(source_sheet, address)
            columns = self.source_column_names(frame)
            weight = "" if self.weight_var.get() == "(none)" else self.weight_var.get()
            for column in (self.x_var.get(), self.y_var.get(), weight):
                if column and column not in columns:
                    raise ValueError(
                        f"Column '{column}' is not present in the selected source."
                    )

            settings_sheet = self.settings_sheet(book, create=True)
            saved = self.read_saved_settings(settings_sheet)
            existing = next(
                (
                    (saved_name, entry)
                    for saved_name, entry in saved.items()
                    if saved_name.casefold() == name.casefold()
                ),
                None,
            )
            if existing is not None:
                if not messagebox.askyesno(
                    "Replace saved settings?",
                    f"A settings set named '{existing[0]}' already exists. Replace it?",
                    parent=self.root,
                ):
                    return
                row_number = existing[1][0]
            else:
                row_number = settings_sheet.used_range.last_cell.row + 1

            values = [
                name,
                self.sheet_var.get(),
                self.table_var.get(),
                address,
                self.x_var.get(),
                self.y_var.get(),
                weight,
                self.x_bins_var.get(),
                self.y_bins_var.get(),
                self.aggregate_var.get(),
                self.plot_title_var.get(),
                self.x_title_var.get(),
                self.y_title_var.get(),
                self.cmap_var.get(),
                self.reverse_cmap_var.get(),
                self.x_min_var.get(),
                self.x_max_var.get(),
                self.y_min_var.get(),
                self.y_max_var.get(),
                self.legend_min_var.get(),
                self.legend_max_var.get(),
                self.under_color_var.get(),
                self.over_color_var.get(),
            ]
            settings_sheet.range((row_number, 1)).value = [values]
            if self.refresh_saved_settings():
                self.saved_settings_var.set(name)
                self.set_status(f"Saved graph settings as '{name}'.")
        except (
            ValueError,
            TypeError,
            KeyError,
            IndexError,
            tk.TclError,
            pywintypes.com_error,
        ) as exc:
            self.set_status(str(exc), error=True)

    def load_settings(self):
        try:
            name = self.saved_settings_var.get()
            if not name or name not in self.saved_settings:
                raise ValueError("Select a saved settings set to load.")
            record = self.saved_settings[name][1]

            def text_value(field):
                value = record.get(field)
                return "" if value is None else str(value)

            book = self.selected_book()
            saved_sheet_name = text_value("Source sheet")
            source_sheets = {
                sheet.name.casefold(): sheet for sheet in book.sheets
            }
            source_sheet = source_sheets.get(saved_sheet_name.casefold())
            if source_sheet is None:
                raise ValueError(
                    f"The saved source sheet '{saved_sheet_name}' no longer exists."
                )
            table_name = text_value("Excel table")
            if table_name:
                tables = {table.name.casefold(): table.name for table in source_sheet.tables}
                if table_name.casefold() not in tables:
                    raise ValueError(
                        f"The saved Excel table '{table_name}' no longer exists "
                        f"on sheet '{source_sheet.name}'."
                    )
                table_name = tables[table_name.casefold()]
            address = text_value("Source range")
            if not address:
                raise ValueError("The saved settings do not contain a source range.")
            frame = self.read_source_frame(source_sheet, address, table_name)
            columns = self.source_column_names(frame)
            x_column = text_value("X column")
            y_column = text_value("Y column")
            weight_column = text_value("Weight column")
            for column in (x_column, y_column, weight_column):
                if column and column not in columns:
                    raise ValueError(
                        f"Saved column '{column}' is not present in the selected source."
                    )
            if not x_column or not y_column:
                raise ValueError("The saved settings must include both X and Y columns.")

            aggregation = text_value("Aggregation")
            if aggregation not in (
                "Bin count",
                "Sum of weights",
                "Average of weights",
            ):
                raise ValueError(f"Unknown saved aggregation '{aggregation}'.")
            if aggregation != "Bin count" and not weight_column:
                raise ValueError("The saved settings need a weight column for this aggregation.")
            colormap = text_value("Weights colormap")
            if colormap not in plt.colormaps():
                raise ValueError(f"Unknown saved colormap '{colormap}'.")
            for field in ("Below-range color", "Above-range color"):
                if not is_color_like(text_value(field)):
                    raise ValueError(f"The saved {field.lower()} is not a valid color.")

            reverse_value = record.get("Reverse colormap")
            if isinstance(reverse_value, bool):
                reverse = reverse_value
            elif isinstance(reverse_value, (int, float)) and reverse_value in (0, 1):
                reverse = bool(reverse_value)
            elif isinstance(reverse_value, str) and reverse_value.strip().lower() in (
                "true",
                "false",
            ):
                reverse = reverse_value.strip().lower() == "true"
            else:
                raise ValueError("The saved reverse-colormap setting is invalid.")

            x_min = parse_optional_limit(text_value("X-axis minimum"), "X-axis minimum")
            x_max = parse_optional_limit(text_value("X-axis maximum"), "X-axis maximum")
            y_min = parse_optional_limit(text_value("Y-axis minimum"), "Y-axis minimum")
            y_max = parse_optional_limit(text_value("Y-axis maximum"), "Y-axis maximum")
            legend_min = parse_optional_limit(
                text_value("Legend minimum"), "Legend minimum"
            )
            legend_max = parse_optional_limit(
                text_value("Legend maximum"), "Legend maximum"
            )
            validate_limits(x_min, x_max, "X-axis")
            validate_limits(y_min, y_max, "Y-axis")
            validate_limits(legend_min, legend_max, "Legend")

            x_values = pd.to_numeric(frame[x_column], errors="coerce").to_numpy(
                dtype=float
            )
            y_values = pd.to_numeric(frame[y_column], errors="coerce").to_numpy(
                dtype=float
            )
            valid = np.isfinite(x_values) & np.isfinite(y_values)
            if aggregation != "Bin count":
                weights = pd.to_numeric(
                    frame[weight_column], errors="coerce"
                ).to_numpy(dtype=float)
                valid &= np.isfinite(weights)
            if not valid.any():
                raise ValueError("The saved source has no usable numeric observations.")
            make_bin_edges(
                text_value("X bin width"), x_values[valid], x_min, x_max, "X"
            )
            make_bin_edges(
                text_value("Y bin height"), y_values[valid], y_min, y_max, "Y"
            )

            self.trend_navigator.set_columns(columns, frame)
            self.sheet_combo.configure(
                values=[sheet.name for sheet in book.sheets]
            )
            self.sheet_var.set(source_sheet.name)
            self.table_combo.configure(
                values=[table.name for table in source_sheet.tables]
            )
            self.table_var.set(table_name)
            self.source_range_var.set(address)
            self.x_combo.configure(values=columns)
            self.y_combo.configure(values=columns)
            self.weight_combo.configure(values=["(none)"] + columns)
            self.x_var.set(x_column)
            self.y_var.set(y_column)
            self.weight_var.set(weight_column or "(none)")
            self.x_bins_var.set(text_value("X bin width"))
            self.y_bins_var.set(text_value("Y bin height"))
            self.aggregate_var.set(aggregation)
            self.settings_name_var.set(name)
            self.plot_title_var.set(text_value("Graph title"))
            self.x_title_var.set(text_value("X-axis title"))
            self.y_title_var.set(text_value("Y-axis title"))
            self.cmap_var.set(colormap)
            self.reverse_cmap_var.set(reverse)
            self.x_min_var.set(text_value("X-axis minimum"))
            self.x_max_var.set(text_value("X-axis maximum"))
            self.y_min_var.set(text_value("Y-axis minimum"))
            self.y_max_var.set(text_value("Y-axis maximum"))
            self.legend_min_var.set(text_value("Legend minimum"))
            self.legend_max_var.set(text_value("Legend maximum"))
            self.under_color_var.set(text_value("Below-range color"))
            self.over_color_var.set(text_value("Above-range color"))
            self.update_cmap_preview()
            self.set_status(
                f"Loaded settings '{name}' and restored its source dataset and graph options."
            )
        except (
            ValueError,
            TypeError,
            KeyError,
            IndexError,
            tk.TclError,
            pywintypes.com_error,
        ) as exc:
            self.set_status(str(exc), error=True)

    def update_cmap_preview(self):
        cmap = copy.copy(plt.get_cmap(self.cmap_var.get()))
        if self.reverse_cmap_var.get():
            cmap = cmap.reversed()
        self.cmap_preview.delete("all")
        width = 400
        for index in range(width):
            red, green, blue = cmap(index / (width - 1))[:3]
            color = f"#{round(red * 255):02x}{round(green * 255):02x}{round(blue * 255):02x}"
            self.cmap_preview.create_line(index, 0, index, 18, fill=color)

    def plot_heatmap(self):
        try:
            if not self.sheet_var.get():
                raise ValueError("Select an open workbook and sheet first.")
            address = self.source_range_var.get().strip()
            if not address:
                raise ValueError(
                    "Choose a table or enter an A1 source range, then load its columns."
                )
            if not self.x_var.get() or not self.y_var.get():
                raise ValueError("Load the source columns and select both X and Y columns.")
            sheet = self.selected_book().sheets[self.sheet_var.get()]
            frame = self.read_source_frame(sheet, address)
            if frame is None or frame.empty:
                raise ValueError("The selected source contains no data rows.")
            frame = frame.copy()
            frame.columns = [str(column) for column in frame.columns]
            x = pd.to_numeric(frame[self.x_var.get()], errors="coerce").to_numpy(
                dtype=float
            )
            y = pd.to_numeric(frame[self.y_var.get()], errors="coerce").to_numpy(
                dtype=float
            )
            aggregation = self.aggregate_var.get()
            use_weights = aggregation != "Bin count" and self.weight_var.get() != "(none)"
            weights = None
            valid = np.isfinite(x) & np.isfinite(y)
            if use_weights:
                weights = pd.to_numeric(
                    frame[self.weight_var.get()], errors="coerce"
                ).to_numpy(dtype=float)
                valid &= np.isfinite(weights)
            x, y = x[valid], y[valid]
            if weights is not None:
                weights = weights[valid]
            if x.size == 0:
                raise ValueError(
                    "No rows have usable numeric X/Y values (and numeric weights, when selected)."
                )
            x_min = parse_optional_limit(self.x_min_var.get(), "X-axis minimum")
            x_max = parse_optional_limit(self.x_max_var.get(), "X-axis maximum")
            y_min = parse_optional_limit(self.y_min_var.get(), "Y-axis minimum")
            y_max = parse_optional_limit(self.y_max_var.get(), "Y-axis maximum")
            legend_min = parse_optional_limit(
                self.legend_min_var.get(), "Legend minimum"
            )
            legend_max = parse_optional_limit(
                self.legend_max_var.get(), "Legend maximum"
            )
            validate_limits(x_min, x_max, "X-axis")
            validate_limits(y_min, y_max, "Y-axis")
            validate_limits(legend_min, legend_max, "Legend")

            x_edges, x_display_min, x_display_max = make_bin_edges(
                self.x_bins_var.get(), x, x_min, x_max, "X"
            )
            y_edges, y_display_min, y_display_max = make_bin_edges(
                self.y_bins_var.get(), y, y_min, y_max, "Y"
            )
            in_displayed_range = (
                (x >= x_display_min)
                & (x <= x_display_max)
                & (y >= y_display_min)
                & (y <= y_display_max)
            )
            x, y = x[in_displayed_range], y[in_displayed_range]
            if weights is not None:
                weights = weights[in_displayed_range]
            if x.size == 0:
                raise ValueError("No observations fall within the displayed axis ranges.")

            if aggregation == "Bin count":
                values, _, _ = np.histogram2d(x, y, bins=(x_edges, y_edges))
                color_label = "Bin count"
            else:
                if weights is None:
                    raise ValueError("Select a weight column for sum or average aggregation.")
                sums, _, _ = np.histogram2d(
                    x, y, bins=(x_edges, y_edges), weights=weights
                )
                if aggregation == "Sum of weights":
                    values = sums
                    color_label = "Sum of weights"
                else:
                    counts, _, _ = np.histogram2d(x, y, bins=(x_edges, y_edges))
                    values = np.divide(
                        sums,
                        counts,
                        out=np.full_like(sums, np.nan),
                        where=counts > 0,
                    )
                    color_label = "Average of weights"

            figure_size = self.get_figure_size_inches()
            figure = Figure(figsize=figure_size or (9, 6))
            ax = figure.add_subplot(111)
            cmap = copy.copy(plt.get_cmap(self.cmap_var.get()))
            if self.reverse_cmap_var.get():
                cmap = cmap.reversed()
            cmap.set_under(self.under_color_var.get())
            cmap.set_over(self.over_color_var.get())
            image = ax.pcolormesh(
                x_edges,
                y_edges,
                np.ma.masked_invalid(values.T),
                shading="auto",
                cmap=cmap,
                vmin=legend_min,
                vmax=legend_max,
            )
            extend = (
                "both"
                if legend_min is not None and legend_max is not None
                else "min"
                if legend_min is not None
                else "max"
                if legend_max is not None
                else "neither"
            )
            figure.colorbar(image, ax=ax, label=color_label, extend=extend)
            ax.set_xlabel(self.x_title_var.get().strip() or self.x_var.get())
            ax.set_ylabel(self.y_title_var.get().strip() or self.y_var.get())
            ax.set_title(
                self.plot_title_var.get().strip() or f"{color_label} heat map"
            )
            ax.set_xlim(left=x_display_min, right=x_display_max)
            ax.set_ylim(bottom=y_display_min, top=y_display_max)
            figure.tight_layout()
            self.figure_fixed_size = figure_size is not None
            self._display_figure(figure)
            self.set_status(
                f"Plotted {x.size} valid observations across "
                f"{len(x_edges) - 1} × {len(y_edges) - 1} bins."
            )
        except (ValueError, TypeError, KeyError, IndexError) as exc:
            self.set_status(str(exc), error=True)

    def _display_figure(self, figure):
        if self.figure_canvas is not None:
            self.figure_canvas.get_tk_widget().destroy()
        if self.figure is not None:
            self.figure.clear()
        self.figure = figure
        self.figure_canvas = FigureCanvasTkAgg(figure, master=self.plot_frame)
        self.figure_canvas.draw()
        if self.figure_fixed_size:
            self.figure_canvas.get_tk_widget().pack(anchor="nw")
        else:
            self.figure_canvas.get_tk_widget().pack(fill="both", expand=True)

    def save_graph(self, figure=None, graph_title="heat map", use_heatmap=True):
        if figure is None and use_heatmap:
            figure = self.figure
        if figure is None:
            self.set_status(f"Generate a {graph_title} before saving it.", error=True)
            return
        path = filedialog.asksaveasfilename(
            parent=self.root,
            title=f"Save {graph_title}",
            defaultextension=".png",
            filetypes=[
                ("PNG image", "*.png"),
                ("JPEG image", "*.jpg *.jpeg"),
                ("PDF document", "*.pdf"),
                ("SVG image", "*.svg"),
            ],
        )
        if not path:
            return
        try:
            figure_size = self.get_figure_size_inches()
            previous_size = figure.get_size_inches().copy()
            try:
                if figure_size is not None:
                    figure.set_size_inches(*figure_size, forward=False)
                figure.savefig(
                    path,
                    dpi=GRAPH_EXPORT_DPI,
                    bbox_inches=None if figure_size is not None else "tight",
                )
            finally:
                if figure_size is not None:
                    figure.set_size_inches(*previous_size, forward=False)
        except (OSError, RuntimeError, ValueError, tk.TclError) as exc:
            self.set_status(f"Could not save graph: {exc}", error=True)
        else:
            self.set_status(f"Graph saved to {path}")

    def copy_graph_to_clipboard(self, figure=None, use_heatmap=True):
        if figure is None and use_heatmap:
            figure = self.figure
        if figure is None:
            self.set_status("Generate a graph before copying it.", error=True)
            return
        try:
            dib = figure_to_dib(figure, self.get_figure_size_inches())
            win32clipboard.OpenClipboard()
            try:
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardData(win32con.CF_DIB, dib)
            finally:
                win32clipboard.CloseClipboard()
        except (OSError, RuntimeError, ValueError, tk.TclError, pywintypes.error) as exc:
            self.set_status(f"Could not copy graph to clipboard: {exc}", error=True)
        else:
            self.set_status(
                f"Graph copied to the clipboard at {GRAPH_EXPORT_DPI} DPI."
            )


def main():
    root = tk.Tk()
    HeatMapApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
