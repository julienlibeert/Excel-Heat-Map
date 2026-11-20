"""Scatterplot configuration and interactive Tk canvas."""

import tkinter as tk
from tkinter import ttk
from typing import TypedDict, overload

import matplotlib
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.backends._backend_tk import NavigationToolbar2Tk
from mpl_toolkits.mplot3d import Axes3D

from ..plots.scatterplot import create_2d_scatter, create_3d_scatter
from ..utils import parse_optional_limit, validate_limits
from .plot_controls import colormap_preview, protect_combobox_wheel


class ScatterArguments(TypedDict):
    x_series: str
    y_series: str
    color_series: str | None
    title: str
    colormap_name: str
    marker_type: str
    marker_size: float
    marker_alpha: float
    figure_size: tuple[float, float] | None
    x_label: str | None
    y_label: str | None
    color_label: str | None
    x_min: float | None
    x_max: float | None
    y_min: float | None
    y_max: float | None
    color_min: float | None
    color_max: float | None


class ScatterPlotNavigator:
    def __init__(
        self,
        parent,
        data_provider,
        save_figure,
        copy_figure,
        set_status,
        figure_size_inches,
        save_settings,
        refresh_settings,
        load_settings,
    ):
        self.parent = parent
        self.data_provider = data_provider
        self.save_figure_callback = save_figure
        self.copy_figure_callback = copy_figure
        self.set_status = set_status
        self.figure_size_inches = figure_size_inches
        self.save_settings_callback = save_settings
        self.refresh_settings_callback = refresh_settings
        self.load_settings_callback = load_settings
        self.columns = []
        self._source_frame = None
        self.figure = self.canvas = self.axes = self.toolbar = None
        self._rotation_start: tuple[float, float, float, float] | None = None
        self.settings_name = tk.StringVar()
        self.saved_settings_name = tk.StringVar()
        defaults = dict(
            x="",
            y="",
            z="(None)",
            color="(None)",
            title="Scatterplot",
            colormap="viridis",
            marker="o",
            marker_size="15",
            transparency="50",
            elevation="30",
            azimuth="45",
        )
        for axis in ("x", "y", "z", "color"):
            for suffix in ("label", "min", "max"):
                defaults[f"{axis}_{suffix}"] = ""
        self.options = {
            key: tk.StringVar(value=value) for key, value in defaults.items()
        }
        self._build_controls()

    @overload
    def _row(
        self, parent, label: str, variable: tk.StringVar, values: list[str]
    ) -> ttk.Combobox: ...

    @overload
    def _row(
        self, parent, label: str, variable: tk.StringVar, values: None = None
    ) -> ttk.Entry: ...

    def _row(self, parent, label, variable, values=None):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=2)
        ttk.Label(row, text=label, width=22).pack(side="left")
        if values is None:
            control = ttk.Entry(row, textvariable=variable)
        else:
            control = ttk.Combobox(
                row, textvariable=variable, values=values, state="readonly"
            )
            protect_combobox_wheel(control)
        control.pack(side="left", fill="x", expand=True)
        return control

    def _build_controls(self):
        saved = ttk.LabelFrame(
            self.parent, text="Saved scatterplot settings for this workbook", padding=8
        )
        saved.pack(fill="x")
        self._row(saved, "Settings name:", self.settings_name)
        self.saved_settings_combo = self._row(
            saved, "Saved settings:", self.saved_settings_name, []
        )
        buttons = ttk.Frame(saved)
        buttons.pack(anchor="w")
        for label, callback in (
            ("Save settings", self.save_settings),
            ("Refresh saved settings", self.refresh_settings_callback),
            (
                "Load settings",
                lambda: self.load_settings_callback(self.saved_settings_name.get()),
            ),
        ):
            ttk.Button(buttons, text=label, command=callback).pack(
                side="left", padx=(0, 6)
            )
        settings = ttk.LabelFrame(self.parent, text="Scatterplot options", padding=8)
        settings.pack(fill="x", pady=6)
        self._row(settings, "Graph title:", self.options["title"])
        self.column_combos = {}
        for axis in ("x", "y", "z", "color"):
            combo = self._row(
                settings, f"{axis.upper()} column:", self.options[axis], []
            )
            combo.master.pack_configure(pady=(10, 2))
            self.column_combos[axis] = combo
            ttk.Label(combo.master, text="Label override:").pack(
                side="left", padx=(12, 4)
            )
            ttk.Entry(
                combo.master, textvariable=self.options[f"{axis}_label"], width=24
            ).pack(side="left", fill="x", expand=True)
            limits = ttk.Frame(settings)
            limits.pack(fill="x")
            for suffix in ("min", "max"):
                ttk.Label(limits, text=f"{axis.upper()} {suffix}:").pack(side="left")
                ttk.Entry(
                    limits, textvariable=self.options[f"{axis}_{suffix}"], width=12
                ).pack(side="left", padx=6)
        self._row(
            settings,
            "Colormap:",
            self.options["colormap"],
            sorted(matplotlib.colormaps),
        )
        self.cmap_preview = colormap_preview(settings, self.options["colormap"])
        marker_combo = self._row(
            settings,
            "Marker:",
            self.options["marker"],
            ["o", ".", "s", "^", "v", "D", "+", "x", "*"],
        )
        marker_combo.master.pack_configure(pady=(12, 2))
        self._row(settings, "Marker size:", self.options["marker_size"])
        transparency_row = ttk.Frame(settings)
        transparency_row.pack(fill="x", pady=2)
        ttk.Label(transparency_row, text="Marker transparency:", width=22).pack(
            side="left"
        )
        ttk.Label(transparency_row, text="Opaque").pack(side="left")
        self.transparency_value = tk.DoubleVar(value=50)
        self.transparency_slider = ttk.Scale(
            transparency_row,
            from_=0,
            to=100,
            variable=self.transparency_value,
            command=self._update_marker_transparency,
        )
        self.transparency_slider.pack(side="left", fill="x", expand=True, padx=8)
        ttk.Label(transparency_row, text="Invisible").pack(side="left")
        self.transparency_text = tk.StringVar(value="50%")
        ttk.Label(transparency_row, textvariable=self.transparency_text, width=6).pack(
            side="left", padx=(8, 0)
        )
        for key in ("elevation", "azimuth"):
            self._row(settings, f"3D {key}:", self.options[key])
        ttk.Label(
            settings,
            text="Z = (None): 2D box zoom. With Z: drag left/right to orbit, up/down to tilt (no roll). Home resets the view.",
            wraplength=900,
        ).pack(anchor="w")
        buttons = ttk.Frame(self.parent)
        buttons.pack(anchor="w", pady=5)
        for label, callback in (
            ("Plot scatterplot", self.plot_scatter),
            (
                "Save graph",
                lambda: self.save_figure_callback(
                    self.figure, self.options["title"].get() or "scatterplot"
                ),
            ),
            ("Copy graph to clipboard", lambda: self.copy_figure_callback(self.figure)),
        ):
            ttk.Button(buttons, text=label, command=callback).pack(
                side="left", padx=(0, 6)
            )
        self.graph_frame = ttk.Frame(self.parent)
        self.graph_frame.pack(fill="both", expand=True)

    def set_columns(self, columns, source_frame=None):
        self.columns = list(columns)
        self._source_frame = None if source_frame is None else source_frame.copy()
        if self._source_frame is not None:
            self._source_frame.columns = self.columns
        for axis, combo in self.column_combos.items():
            choices = (["(None)"] if axis in ("z", "color") else []) + self.columns
            combo.configure(values=choices)
            if self.options[axis].get() not in choices:
                index = min(axis == "y", len(choices) - 1)
                self.options[axis].set(choices[index] if choices else "")

    def set_saved_settings_names(self, names: list[str]):
        names = list(names)
        self.saved_settings_combo.configure(values=names)
        if self.saved_settings_name.get() not in names:
            self.saved_settings_name.set(names[0] if names else "")

    def export_settings(self):
        return {"version": 2, **{key: var.get() for key, var in self.options.items()}}

    def _update_marker_transparency(self, value=None):
        if value is not None:
            self.options["transparency"].set(value)
        transparency = float(self.options["transparency"].get())
        self.transparency_value.set(transparency)
        self.transparency_text.set(f"{transparency:.0f}%")
        if self.axes is not None and self.canvas is not None:
            for points in self.axes.collections:
                points.set_alpha(1 - transparency / 100)
            self.canvas.draw_idle()

    def _on_plot_press(self, event):
        self._rotation_start = None
        if (
            isinstance(self.axes, Axes3D)
            and event.inaxes is self.axes
            and event.button == 1
            and self.axes.get_navigate_mode() is None
        ):
            self._rotation_start = (event.x, event.y, self.axes.elev, self.axes.azim)

    def _on_plot_motion(self, event):
        if (
            self._rotation_start is None
            or not isinstance(self.axes, Axes3D)
            or event.inaxes is not self.axes
            or self.axes.get_navigate_mode() is not None
        ):
            return
        x, y, elevation, azimuth = self._rotation_start
        elevation -= (event.y - y) / self.axes.bbox.height * 180
        azimuth -= (event.x - x) / self.axes.bbox.width * 180
        # Keep the vertical axis upright instead of rolling or flipping at the poles.
        self.axes.view_init(
            elev=float(np.clip(elevation, -89, 89)), azim=azimuth, roll=0
        )
        if self.canvas is not None:
            self.canvas.draw_idle()

    def _on_plot_release(self, _event):
        self._rotation_start = None
        if isinstance(self.axes, Axes3D):
            self.options["elevation"].set(str(self.axes.elev))
            self.options["azimuth"].set(str(self.axes.azim))

    def validate_settings(self, settings, columns=None, source_frame=None):
        if (
            not isinstance(settings, dict)
            or type(settings.get("version")) is not int
            or settings["version"] not in (1, 2)
        ):
            raise ValueError("Unsupported scatterplot settings format.")
        if settings["version"] == 1:
            # Older settings used a fixed 50% marker transparency.
            settings = {**settings, "transparency": settings.get("transparency", "50")}
        if any(not isinstance(settings.get(key), str) for key in self.options):
            raise ValueError("Scatterplot settings contain missing or invalid options.")
        columns = self.columns if columns is None else columns
        frame = self._source_frame if source_frame is None else source_frame
        axes = [settings["x"], settings["y"]]
        if settings["z"] != "(None)":
            axes.append(settings["z"])
        required = axes + ([] if settings["color"] == "(None)" else [settings["color"]])
        if any(column not in columns for column in required):
            raise ValueError("Select source columns for all scatterplot dimensions.")
        if settings["colormap"] not in matplotlib.colormaps:
            raise ValueError("Select a valid scatterplot colormap.")
        if settings["marker"] not in ("o", ".", "s", "^", "v", "D", "+", "x", "*"):
            raise ValueError("Select a valid scatterplot marker.")
        for key in ("marker_size", "elevation", "azimuth"):
            value = float(settings[key])
            if not np.isfinite(value) or (key == "marker_size" and value <= 0):
                raise ValueError(
                    f"{key} must be a finite {'positive ' if key == 'marker_size' else ''}number."
                )
        transparency = float(settings["transparency"])
        if not np.isfinite(transparency) or not 0 <= transparency <= 100:
            raise ValueError("Marker transparency must be between 0 and 100 percent.")
        for axis in ("x", "y", "z", "color"):
            lower = parse_optional_limit(settings[f"{axis}_min"], f"{axis} minimum")
            upper = parse_optional_limit(settings[f"{axis}_max"], f"{axis} maximum")
            validate_limits(lower, upper, axis)
        if frame is None:
            raise ValueError("Load source data before plotting or saving.")
        valid = np.ones(len(frame), dtype=bool)
        for column in axes:
            valid &= np.isfinite(
                pd.to_numeric(frame[column], errors="coerce").to_numpy(dtype=float)
            )
        if settings["color"] != "(None)":
            colors = frame[settings["color"]]
            numeric = pd.to_numeric(colors, errors="coerce")
            valid &= (
                np.isfinite(numeric.to_numpy(dtype=float))
                if numeric.notna().any()
                else colors.notna().to_numpy()
            )
        if not valid.any():
            raise ValueError("The selected axes have no usable numeric observations.")
        return settings

    def apply_settings(self, settings):
        settings = self.validate_settings(settings)
        for key, variable in self.options.items():
            variable.set(settings[key])
        self._update_marker_transparency()

    def save_settings(self):
        name = self.settings_name.get().strip()
        if not name:
            self.set_status(
                "Enter a name for this scatterplot settings set.", error=True
            )
            return
        self.save_settings_callback(name, self.export_settings())

    def plot_scatter(self):
        figure = None
        try:
            frame = self.data_provider()
            if frame is None or frame.empty:
                raise ValueError("The selected source contains no data rows.")
            frame = frame.copy()
            frame.columns = [str(column) for column in frame.columns]
            if len(set(frame.columns)) != len(frame.columns):
                raise ValueError("The selected source has duplicate column names.")
            settings = self.validate_settings(
                self.export_settings(), list(frame.columns), frame
            )
            kwargs: ScatterArguments = {
                "x_series": settings["x"],
                "y_series": settings["y"],
                "color_series": None
                if settings["color"] == "(None)"
                else settings["color"],
                "title": settings["title"],
                "colormap_name": settings["colormap"],
                "marker_type": settings["marker"],
                "marker_size": float(settings["marker_size"]),
                "marker_alpha": 1 - float(settings["transparency"]) / 100,
                "figure_size": self.figure_size_inches(),
                "x_label": settings["x_label"] or None,
                "y_label": settings["y_label"] or None,
                "color_label": settings["color_label"] or None,
                "x_min": parse_optional_limit(settings["x_min"], "X minimum"),
                "x_max": parse_optional_limit(settings["x_max"], "X maximum"),
                "y_min": parse_optional_limit(settings["y_min"], "Y minimum"),
                "y_max": parse_optional_limit(settings["y_max"], "Y maximum"),
                "color_min": parse_optional_limit(
                    settings["color_min"], "Color minimum"
                ),
                "color_max": parse_optional_limit(
                    settings["color_max"], "Color maximum"
                ),
            }
            is_3d = settings["z"] != "(None)"
            if is_3d:
                figure, axes = create_3d_scatter(
                    frame,
                    z_series=settings["z"],
                    elevation=float(settings["elevation"]),
                    azimuth=float(settings["azimuth"]),
                    z_label=settings["z_label"] or None,
                    z_min=parse_optional_limit(settings["z_min"], "Z minimum"),
                    z_max=parse_optional_limit(settings["z_max"], "Z maximum"),
                    **kwargs,
                )
            else:
                figure, axes = create_2d_scatter(frame, **kwargs)
            # Embed the figure without leaving a pyplot-managed extra window alive.
            plt.close(figure)
            for widget in self.graph_frame.winfo_children():
                widget.destroy()
            self.figure, self.axes = figure, axes
            self._rotation_start = None
            self.canvas = FigureCanvasTkAgg(figure, master=self.graph_frame)
            self.toolbar = NavigationToolbar2Tk(
                self.canvas, self.graph_frame, pack_toolbar=False
            )
            self.toolbar.pack(fill="x")
            if kwargs["figure_size"] is None:
                self.canvas.get_tk_widget().pack(fill="both", expand=True)
            else:
                self.canvas.get_tk_widget().pack(anchor="nw")
            self.canvas.mpl_connect("button_press_event", self._on_plot_press)
            self.canvas.mpl_connect("motion_notify_event", self._on_plot_motion)
            self.canvas.mpl_connect("button_release_event", self._on_plot_release)
            if isinstance(axes, Axes3D):
                # Matplotlib accepts button lists, but its stubs only declare int.
                axes.mouse_init(rotate_btn=[])  # type: ignore[arg-type]
            else:
                self.toolbar.zoom()
            self.canvas.draw()
            self._update_marker_transparency()
            self.set_status("Scatterplot generated.")
            return True
        except (ValueError, TypeError, KeyError, IndexError, tk.TclError) as exc:
            if figure is not None:
                plt.close(figure)
            self.set_status(str(exc), error=True)
            return False
