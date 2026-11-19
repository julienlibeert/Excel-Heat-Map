import copy
import io
import re
import tkinter as tk
from tkinter import colorchooser, filedialog, ttk

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


GRAPH_EXPORT_DPI = 300


def figure_to_dib(figure):
    png = io.BytesIO()
    figure.savefig(png, format="png", dpi=GRAPH_EXPORT_DPI, bbox_inches="tight")
    png.seek(0)
    with Image.open(png) as image:
        bitmap = io.BytesIO()
        image.convert("RGB").save(bitmap, format="BMP")
    return bitmap.getvalue()[14:]


def parse_bins(specification, values, axis_label):
    text = specification.strip()
    if not text:
        raise ValueError(f"Enter a bin count or explicit edges for {axis_label}.")
    if re.fullmatch(r"\d+", text):
        count = int(text)
        if count < 1:
            raise ValueError(f"{axis_label} bin count must be at least 1.")
        return np.histogram_bin_edges(values, bins=count)
    try:
        edges = np.asarray([float(part.strip()) for part in text.split(",")], dtype=float)
    except ValueError as exc:
        raise ValueError(
            f"{axis_label} bins must be an integer count or comma-separated numeric edges."
        ) from exc
    if len(edges) < 2 or not np.isfinite(edges).all() or not np.all(np.diff(edges) > 0):
        raise ValueError(
            f"{axis_label} bin edges must contain at least two finite, strictly increasing values."
        )
    return edges


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
        self.root.title("Excel Heat Map")
        self.root.geometry("1100x900")
        self.workbook_entries = {}
        self.figure = None
        self.figure_canvas = None

        self._build_window()
        self.refresh_workbooks()

    def _build_window(self):
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.scroll_canvas = tk.Canvas(self.root, highlightthickness=0)
        scrollbar = ttk.Scrollbar(
            self.root, orient="vertical", command=self.scroll_canvas.yview
        )
        self.scroll_canvas.configure(yscrollcommand=scrollbar.set)
        self.scroll_canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

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
            lambda event: self.scroll_canvas.itemconfigure(
                self.content_window, width=event.width
            ),
        )
        self.scroll_canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        self.workbook_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.table_var = tk.StringVar()
        self.source_range_var = tk.StringVar()
        self.x_var = tk.StringVar()
        self.y_var = tk.StringVar()
        self.weight_var = tk.StringVar()
        self.x_bins_var = tk.StringVar(value="20")
        self.y_bins_var = tk.StringVar(value="20")
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

        self._build_controls()
        ttk.Separator(self.content).pack(fill="x", pady=(14, 8))
        ttk.Label(self.content, text="Generated heat map", font=("", 12, "bold")).pack(
            anchor="w", pady=(0, 6)
        )
        export_buttons = ttk.Frame(self.content)
        export_buttons.pack(anchor="w", pady=(0, 6))
        ttk.Button(
            export_buttons, text="Save graph", command=self.save_graph
        ).pack(side="left", padx=(0, 6))
        ttk.Button(
            export_buttons,
            text="Copy graph to clipboard",
            command=self.copy_graph_to_clipboard,
        ).pack(side="left")
        self.plot_frame = ttk.Frame(self.content)
        self.plot_frame.pack(fill="both", expand=True)

    def _on_mousewheel(self, event):
        self.scroll_canvas.yview_scroll(int(-event.delta / 120), "units")

    def _section(self, text):
        ttk.Label(self.content, text=text, font=("", 10, "bold")).pack(
            anchor="w", pady=(10, 4)
        )

    def _entry_row(self, label, variable, placeholder=None, width=36):
        row = ttk.Frame(self.content)
        row.pack(fill="x", pady=2)
        ttk.Label(row, text=label, width=22).pack(side="left")
        entry = ttk.Entry(row, textvariable=variable, width=width)
        entry.pack(side="left", fill="x", expand=True)
        if placeholder:
            ttk.Label(row, text=placeholder, foreground="#666").pack(
                side="left", padx=(8, 0)
            )
        return entry

    def _combo_row(self, label, variable, values=(), width=48):
        row = ttk.Frame(self.content)
        row.pack(fill="x", pady=2)
        ttk.Label(row, text=label, width=22).pack(side="left")
        combo = ttk.Combobox(
            row, textvariable=variable, values=list(values), width=width, state="readonly"
        )
        combo.pack(side="left", fill="x", expand=True)
        return combo

    def _build_controls(self):
        self._section("1. Choose an open workbook and sheet")
        row = ttk.Frame(self.content)
        row.pack(fill="x", pady=2)
        ttk.Label(row, text="Workbook:", width=22).pack(side="left")
        self.workbook_combo = ttk.Combobox(
            row, textvariable=self.workbook_var, state="readonly"
        )
        self.workbook_combo.pack(side="left", fill="x", expand=True)
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

        self._section("2. Select a table or enter a source range (include headers)")
        self.table_combo = self._combo_row("Excel table:", self.table_var)
        self.table_combo.bind(
            "<<ComboboxSelected>>", lambda *_args: self.on_table_change()
        )
        self._entry_row("Source range:", self.source_range_var, "e.g. A1:C500")
        ttk.Button(self.content, text="Load columns", command=self.load_columns).pack(
            anchor="w", pady=4
        )

        self._section("3. Choose data columns and heat-map settings")
        self.x_combo = self._combo_row("X column:", self.x_var)
        self.y_combo = self._combo_row("Y column:", self.y_var)
        self.weight_combo = self._combo_row("Weight:", self.weight_var)
        self._entry_row("X bins:", self.x_bins_var, "count or comma-separated edges")
        self._entry_row("Y bins:", self.y_bins_var, "count or comma-separated edges")
        self._combo_row(
            "Aggregate:",
            self.aggregate_var,
            ["Bin count", "Sum of weights", "Average of weights"],
        )
        self._entry_row("Graph title:", self.plot_title_var)
        self._entry_row("X-axis title:", self.x_title_var, "blank uses column name")
        self._entry_row("Y-axis title:", self.y_title_var, "blank uses column name")

        self._section("Weights colormap")
        row = ttk.Frame(self.content)
        row.pack(fill="x", pady=2)
        ttk.Label(row, text="Weights colormap:", width=22).pack(side="left")
        self.cmap_combo = ttk.Combobox(
            row,
            textvariable=self.cmap_var,
            values=sorted(plt.colormaps()),
            state="readonly",
        )
        self.cmap_combo.pack(side="left", fill="x", expand=True)
        self.cmap_combo.bind(
            "<<ComboboxSelected>>", lambda *_args: self.update_cmap_preview()
        )
        ttk.Checkbutton(
            row,
            text="Reverse gradient",
            variable=self.reverse_cmap_var,
            command=self.update_cmap_preview,
        ).pack(side="left", padx=(8, 0))
        self.cmap_preview = tk.Canvas(
            self.content, height=18, width=400, highlightthickness=0
        )
        self.cmap_preview.pack(anchor="w", padx=(22, 0), pady=3)
        self.update_cmap_preview()

        self._section("Axis ranges (leave blank for automatic limits)")
        self._entry_row("X-axis minimum:", self.x_min_var)
        self._entry_row("X-axis maximum:", self.x_max_var)
        self._entry_row("Y-axis minimum:", self.y_min_var)
        self._entry_row("Y-axis maximum:", self.y_max_var)

        self._section("Legend range and out-of-range colors")
        self._entry_row("Legend minimum:", self.legend_min_var)
        self._entry_row("Legend maximum:", self.legend_max_var)
        self._color_row("Below range:", self.under_color_var)
        self._color_row("Above range:", self.over_color_var)

        ttk.Button(
            self.content, text="Plot heat map", command=self.plot_heatmap
        ).pack(anchor="w", pady=(10, 4))
        self.status_label = ttk.Label(
            self.content, textvariable=self.status_var, wraplength=900
        )
        self.status_label.pack(anchor="w", pady=(2, 6))

    def _color_row(self, label, variable):
        row = ttk.Frame(self.content)
        row.pack(fill="x", pady=2)
        ttk.Label(row, text=label, width=22).pack(side="left")
        ttk.Entry(row, textvariable=variable, width=18).pack(side="left")
        ttk.Button(
            row, text="Choose color", command=lambda: self.choose_color(variable)
        ).pack(side="left", padx=8)

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

    def refresh_workbooks(self):
        try:
            self.workbook_entries = self.open_workbook_entries()
            labels = list(self.workbook_entries)
            self.workbook_combo.configure(values=labels)
            if labels:
                self.workbook_var.set(labels[0])
                self.refresh_sheets()
                self.set_status(
                    "Open workbooks found. Select a source and load its columns."
                )
            else:
                self.workbook_var.set("")
                self.sheet_combo.configure(values=[])
                self.table_combo.configure(values=[])
                self.sheet_var.set("")
                self.table_var.set("")
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
        except (ValueError, TypeError, KeyError, IndexError) as exc:
            self.set_status(str(exc), error=True)

    def on_sheet_change(self):
        try:
            self.refresh_tables()
        except (ValueError, TypeError, KeyError, IndexError) as exc:
            self.set_status(str(exc), error=True)

    def on_table_change(self):
        try:
            table = self.selected_book().sheets[self.sheet_var.get()].tables[
                self.table_var.get()
            ]
            self.source_range_var.set(table.range.address)
            self.set_status(
                "Table range selected. Load columns to refresh the X/Y/weight choices."
            )
        except (ValueError, TypeError, KeyError, IndexError) as exc:
            self.set_status(str(exc), error=True)

    def read_source_frame(self, sheet, address):
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

    def load_columns(self):
        try:
            if not self.sheet_var.get():
                raise ValueError("Select an open workbook and sheet first.")
            address = self.source_range_var.get().strip()
            if not address:
                raise ValueError("Choose a table or enter an A1 source range with a header row.")
            sheet = self.selected_book().sheets[self.sheet_var.get()]
            frame = self.read_source_frame(sheet, address)
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
                    "The source has column headers that become identical when displayed. Rename them in Excel."
                )
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
            x_edges = parse_bins(self.x_bins_var.get(), x, "X")
            y_edges = parse_bins(self.y_bins_var.get(), y, "Y")
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

            figure, ax = plt.subplots(figsize=(9, 6))
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
            ax.set_xlim(left=x_min, right=x_max)
            ax.set_ylim(bottom=y_min, top=y_max)
            figure.tight_layout()
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
            plt.close(self.figure)
        self.figure = figure
        self.figure_canvas = FigureCanvasTkAgg(figure, master=self.plot_frame)
        self.figure_canvas.draw()
        self.figure_canvas.get_tk_widget().pack(fill="both", expand=True)

    def save_graph(self):
        if self.figure is None:
            self.set_status("Generate a heat map before saving it.", error=True)
            return
        path = filedialog.asksaveasfilename(
            parent=self.root,
            title="Save heat map",
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
            self.figure.savefig(path, dpi=GRAPH_EXPORT_DPI, bbox_inches="tight")
        except (OSError, RuntimeError, ValueError, tk.TclError) as exc:
            self.set_status(f"Could not save graph: {exc}", error=True)
        else:
            self.set_status(f"Graph saved to {path}")

    def copy_graph_to_clipboard(self):
        if self.figure is None:
            self.set_status("Generate a heat map before copying it.", error=True)
            return
        try:
            dib = figure_to_dib(self.figure)
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
