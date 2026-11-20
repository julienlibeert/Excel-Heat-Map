import copy
import tkinter as tk
import unittest
from types import SimpleNamespace
from unittest.mock import patch

import numpy as np
import pandas as pd
from matplotlib.backend_bases import MouseEvent
from matplotlib.collections import QuadMesh
from matplotlib.colors import to_hex
import matplotlib.pyplot as plt
import matplotlib

from excel_heat_map.plots.scatterplot import create_2d_scatter, create_3d_scatter
from excel_heat_map.ui.heatmap_ui import HeatMapApp


class MemoryRange:
    """Small in-memory Excel range for exercising the real persistence methods."""

    def __init__(self, sheet, start, end=None):
        self.sheet, self.start, self.end = sheet, start, end or start

    def resize(self, rows, columns):
        return MemoryRange(
            self.sheet,
            self.start,
            (self.start[0] + rows - 1, self.start[1] + columns - 1),
        )

    def options(self, *_args, **_kwargs):
        return self

    @property
    def value(self):
        return [
            [
                self.sheet.cells.get((row, col))
                for col in range(self.start[1], self.end[1] + 1)
            ]
            for row in range(self.start[0], self.end[0] + 1)
        ]

    @value.setter
    def value(self, rows):
        for row_offset, row in enumerate(rows):
            for col_offset, value in enumerate(row):
                self.sheet.cells[
                    self.start[0] + row_offset, self.start[1] + col_offset
                ] = value


class MemorySheets(list):
    def __getitem__(self, key):
        if isinstance(key, str):
            return next(sheet for sheet in self if sheet.name == key)
        return super().__getitem__(key)

    def add(self, name, after=None):
        sheet = MemorySheet(name)
        self.append(sheet)
        return sheet


class MemorySheet:
    def __init__(self, name):
        self.name, self.cells = name, {}
        self.tables = MemorySheets()

    @property
    def used_range(self):
        return SimpleNamespace(
            last_cell=SimpleNamespace(
                row=max((row for row, _ in self.cells), default=1)
            )
        )

    def range(self, start, end=None):
        return MemoryRange(self, (1, 1) if start == "A1" else start, end)


class PlotWorkflows(unittest.TestCase):
    def setUp(self):
        self.root = tk.Tk()
        self.root.withdraw()
        with patch.object(HeatMapApp, "refresh_workbooks"):
            self.app = HeatMapApp(self.root)
        self.frame = pd.DataFrame(
            {
                "time": pd.date_range("2026-01-01", periods=4),
                "x": [1.0, 2.0, 3.0, 4.0],
                "y": [2.0, 4.0, 6.0, 8.0],
                "z": [4.0, 3.0, 2.0, 1.0],
                "color": [5.0, 5.0, 5.0, 5.0],
            }
        )
        self.source = MemorySheet("Data")
        self.book = SimpleNamespace(sheets=MemorySheets([self.source]))
        self.app.selected_book = lambda: self.book
        self.app.sheet_var.set("Data")
        self.app.source_range_var.set("$A$1:$F$5")
        self.app.read_source_frame = lambda *_args: self.frame.copy()
        self.nav = self.app.scatterplot_navigator
        self.nav.data_provider = lambda: self.frame.copy()
        for nav in (self.nav, self.app.trend_navigator):
            nav.set_columns(list(self.frame.columns), self.frame)
        self.nav.options["x"].set("x")
        self.nav.options["y"].set("y")
        self.nav.options["color"].set("color")

    def tearDown(self):
        plt.close("all")
        self.root.update_idletasks()
        self.root.destroy()

    def mouse(self, name, x, y, button=1):
        event = MouseEvent(name, self.nav.canvas, x, y, button=button)
        self.nav.canvas.callbacks.process(name, event)

    def test_scatter_labels_and_constant_numeric_colors(self):
        for three_d in (False, True):
            with self.subTest(three_d=three_d):
                self.nav.options["z"].set("z" if three_d else "(None)")
                for axis in ("x", "y", "z", "color"):
                    self.nav.options[f"{axis}_label"].set(f"Custom {axis}")
                self.assertTrue(self.nav.plot_scatter())
                ax = self.nav.axes
                self.assertEqual(ax.get_xlabel(), "Custom x")
                self.assertEqual(ax.get_ylabel(), "Custom y")
                if three_d:
                    self.assertEqual(ax.get_zlabel(), "Custom z")
                self.assertEqual(self.nav.figure.axes[1].get_ylabel(), "Custom color")
                self.assertEqual(len(ax.collections), 1)

    def test_categorical_color_alignment_and_no_source_mutation(self):
        frame = pd.DataFrame(
            {
                "x": [1, 2, 3, 4],
                "y": [4, 3, 2, 1],
                "z": [2, 3, 4, 5],
                "c": ["a", "b", None, "a"],
            }
        )
        original = frame.copy(deep=True)
        for plot, extra in (
            (create_2d_scatter, {}),
            (create_3d_scatter, {"z_series": "z"}),
        ):
            figure, ax = plot(frame, "x", "y", color_series="c", **extra)
            points = ax.collections[0]
            np.testing.assert_array_equal(points.get_array(), [0, 1, 0])
            self.assertEqual(
                [label.get_text() for label in figure.axes[1].get_yticklabels()],
                ["a", "b"],
            )
        pd.testing.assert_frame_equal(frame, original)

    def test_mixed_numeric_color_and_invalid_coordinates(self):
        frame = pd.DataFrame(
            {"x": [1, "bad", 3, 4], "y": [2, 3, 4, 5], "c": ["10", "20", "bad", "30"]}
        )
        _, ax = create_2d_scatter(frame, "x", "y", "c")
        np.testing.assert_array_equal(ax.collections[0].get_offsets(), [[1, 2], [4, 5]])
        np.testing.assert_array_equal(ax.collections[0].get_array(), [10, 30])

    def test_box_zoom_and_3d_mouse_rotation(self):
        self.assertTrue(self.nav.plot_scatter())
        axes = self.nav.axes
        before = axes.get_xlim(), axes.get_ylim()
        left, bottom, width, height = axes.bbox.bounds
        self.mouse("button_press_event", left + width * 0.2, bottom + height * 0.2)
        self.mouse("motion_notify_event", left + width * 0.7, bottom + height * 0.7)
        self.mouse("button_release_event", left + width * 0.7, bottom + height * 0.7)
        self.assertLess(np.ptp(axes.get_xlim()), np.ptp(before[0]))
        self.assertLess(np.ptp(axes.get_ylim()), np.ptp(before[1]))
        self.nav.options["z"].set("z")
        self.assertTrue(self.nav.plot_scatter())
        axes = self.nav.axes
        before = axes.elev, axes.azim
        left, bottom, width, height = axes.bbox.bounds
        self.mouse("button_press_event", left + width * 0.5, bottom + height * 0.5)
        self.mouse("motion_notify_event", left + width * 0.65, bottom + height * 0.65)
        self.mouse("button_release_event", left + width * 0.65, bottom + height * 0.65)
        self.assertNotAlmostEqual(axes.elev, before[0])
        self.assertNotAlmostEqual(axes.azim, before[1])
        saved = self.nav.export_settings()
        self.assertAlmostEqual(float(saved["elevation"]), axes.elev)
        self.assertAlmostEqual(float(saved["azimuth"]), axes.azim)

    def test_turntable_rotation_keeps_axes_upright_and_separates_drag_directions(self):
        self.nav.options["z"].set("z")
        self.assertTrue(self.nav.plot_scatter())
        axes = self.nav.axes
        left, bottom, width, height = axes.bbox.bounds

        def drag(start_x, start_y, end_x, end_y):
            self.mouse(
                "button_press_event", left + width * start_x, bottom + height * start_y
            )
            self.mouse(
                "motion_notify_event", left + width * end_x, bottom + height * end_y
            )
            self.mouse(
                "button_release_event", left + width * end_x, bottom + height * end_y
            )

        elevation, azimuth = axes.elev, axes.azim
        drag(0.5, 0.5, 0.7, 0.5)
        self.assertAlmostEqual(axes.elev, elevation)
        self.assertAlmostEqual(axes.azim, azimuth - 36, delta=1)
        self.assertAlmostEqual(axes.roll, 0)
        azimuth = axes.azim
        drag(0.5, 0.5, 0.5, 0.7)
        self.assertAlmostEqual(axes.azim, azimuth)
        self.assertAlmostEqual(axes.elev, elevation - 36, delta=1)
        self.assertAlmostEqual(axes.roll, 0)

        axes.view_init(elev=30, azim=45)
        drag(0.5, 0.95, 0.5, 0.05)
        self.assertEqual(axes.elev, 89)
        axes.view_init(elev=-30, azim=45)
        drag(0.5, 0.05, 0.5, 0.95)
        self.assertEqual(axes.elev, -89)
        self.assertAlmostEqual(axes.roll, 0)
        self.nav.toolbar.home()
        self.assertAlmostEqual(axes.elev, elevation)
        self.assertAlmostEqual(axes.azim, 45)

    def test_transparency_slider_updates_both_plots_and_roundtrips(self):
        for three_d in (False, True):
            self.nav.options["z"].set("z" if three_d else "(None)")
            self.assertTrue(self.nav.plot_scatter())
            for transparency in (0, 25, 100):
                self.nav.transparency_slider.set(transparency)
                self.assertAlmostEqual(
                    self.nav.axes.collections[0].get_alpha(), 1 - transparency / 100
                )
                self.assertEqual(self.nav.transparency_text.get(), f"{transparency}%")
        self.nav.transparency_slider.set(75)
        self.nav.settings_name.set("Transparency")
        self.nav.save_settings()
        self.nav.transparency_slider.set(0)
        self.app.load_trend_settings("Transparency", "scatterplot")
        self.assertEqual(self.nav.transparency_slider.get(), 75)
        self.assertAlmostEqual(self.nav.axes.collections[0].get_alpha(), 0.25)

    def test_transparency_migration_and_validation(self):
        old_settings = self.nav.export_settings()
        old_settings["version"] = 1
        old_settings.pop("transparency")
        self.nav.transparency_slider.set(0)
        self.nav.apply_settings(old_settings)
        self.assertEqual(self.nav.transparency_slider.get(), 50)
        self.assertTrue(self.nav.plot_scatter())
        self.assertAlmostEqual(self.nav.axes.collections[0].get_alpha(), 0.5)
        self.assertNotIn("transparency", old_settings)
        for invalid in ("-1", "101", "nan", "inf", "not a number"):
            with self.subTest(transparency=invalid):
                settings = {**self.nav.export_settings(), "transparency": invalid}
                with self.assertRaises(ValueError):
                    self.nav.validate_settings(settings)
        settings = self.nav.export_settings()
        settings.pop("transparency")
        with self.assertRaises(ValueError):
            self.nav.validate_settings(settings)

    def test_colorbar_stays_opaque_when_marker_transparency_changes(self):
        for colors in ([1.0, 2.0, 3.0, 4.0], ["a", "b", "a", "b"]):
            self.frame["color"] = colors
            for three_d in (False, True):
                with self.subTest(colors=colors, three_d=three_d):
                    self.nav.options["z"].set("z" if three_d else "(None)")
                    self.nav.options["transparency"].set("75")
                    self.assertTrue(self.nav.plot_scatter())
                    solids = next(
                        artist
                        for artist in self.nav.figure.axes[1].collections
                        if isinstance(artist, QuadMesh)
                    )
                    for transparency in (75, 0, 25, 100):
                        self.nav.transparency_slider.set(transparency)
                        self.nav.canvas.draw()
                        self.assertAlmostEqual(
                            self.nav.axes.collections[0].get_alpha(),
                            1 - transparency / 100,
                        )
                        np.testing.assert_array_equal(solids.get_facecolors()[:, 3], 1)

    def test_scatter_option_layout_groups_column_and_label_with_title_first(self):
        for axis, combo in self.nav.column_combos.items():
            entries = [
                widget
                for widget in combo.master.winfo_children()
                if widget.winfo_class() == "TEntry"
            ]
            self.assertEqual(len(entries), 1)
            self.assertEqual(
                entries[0].cget("textvariable"), str(self.nav.options[f"{axis}_label"])
            )
        options = self.nav.column_combos["x"].master.master
        first_row = options.pack_slaves()[0]
        self.assertEqual(first_row.winfo_children()[0].cget("text"), "Graph title:")

    def test_scatter_settings_restore_source_and_camera(self):
        self.nav.options["z"].set("z")
        self.nav.options["elevation"].set("15")
        self.nav.options["x_label"].set("Saved X")
        self.nav.settings_name.set("Example")
        self.nav.save_settings()
        self.assertIn("Example", self.app.saved_scatterplot_settings)
        record = self.app.saved_scatterplot_settings["Example"][1]
        self.assertEqual(record["Source sheet"], "Data")
        self.assertEqual(record["Source range"], "$A$1:$F$5")
        self.app.sheet_var.set("Other")
        self.app.source_range_var.set("Z1:Z2")
        self.nav.options["z"].set("(None)")
        self.app.load_trend_settings("Example", "scatterplot")
        self.assertEqual(self.app.sheet_var.get(), "Data")
        self.assertEqual(self.app.source_range_var.get(), "$A$1:$F$5")
        self.assertEqual(self.nav.axes.elev, 15)
        self.assertEqual(self.nav.axes.get_xlabel(), "Saved X")
        self.nav.options["elevation"].set("60")
        self.assertTrue(self.nav.plot_scatter())
        self.assertEqual(self.nav.axes.elev, 60)
        self.app.load_trend_settings("Example", "scatterplot")
        self.assertEqual(self.nav.axes.elev, 15)

    def test_trend_palette_buttons_and_individual_override_roundtrip(self):
        nav = self.app.trend_navigator
        nav.add_trend()
        nav.color_mode.set("Colormap")
        nav.colormap.set("plasma")
        expected = [
            to_hex(matplotlib.colormaps["plasma"](float(index))) for index in (0, 1)
        ]
        for row, color in zip(nav.trend_rows, expected):
            self.assertEqual(row.color.get(), color)
            self.assertEqual(row.color_button.cget("background"), color)
        first, second = nav.trend_rows
        with patch(
            "excel_heat_map.ui.trend_navigator.colorchooser.askcolor",
            return_value=((255, 0, 0), "#ff0000"),
        ):
            nav.choose_color(first.color, first.color_button)
        self.assertEqual(nav.color_mode.get(), "Colormap")
        self.assertEqual(second.color.get(), expected[1])
        settings = nav.export_settings()
        nav.apply_settings(settings)
        self.assertEqual(first.color.get(), "#ff0000")
        self.assertEqual(second.color.get(), expected[1])
        prepared, _ = nav._prepare_trends(
            self.frame, nav._selected_trends(), list(self.frame.columns)
        )
        self.assertEqual(prepared[0]["color"], "#ff0000")
        nav.settings_name.set("Trend")
        nav.save_settings()
        record = self.app.saved_trend_settings["Trend"][1]
        self.assertEqual(record["Source range"], "$A$1:$F$5")
        self.app.source_range_var.set("Z1:Z2")
        self.app.load_trend_settings("Trend")
        self.assertEqual(self.app.source_range_var.get(), "$A$1:$F$5")
        self.assertEqual(nav.trend_rows[0].color.get(), "#ff0000")
        old = copy.deepcopy(settings)
        old["version"] = 2
        for trend in old["trends"]:
            trend.pop("color_override")
        nav.apply_settings(old)
        self.assertEqual(first.color.get(), expected[0])

    def test_heatmap_settings_roundtrip_preserves_source(self):
        self.app.x_var.set("x")
        self.app.y_var.set("y")
        self.app.weight_var.set("(none)")
        self.app.settings_name_var.set("Heat")
        self.app.save_settings()
        record = self.app.saved_settings["Heat"][1]
        self.assertEqual(record["Source range"], "$A$1:$F$5")
        self.app.source_range_var.set("Z1:Z2")
        self.app.load_settings()
        self.assertEqual(self.app.source_range_var.get(), "$A$1:$F$5")
        self.assertEqual(self.app.x_var.get(), "x")

    def test_saved_table_restores_current_expanded_range_for_all_graphs(self):
        self.source.tables.append(
            SimpleNamespace(
                name="Measurements", range=SimpleNamespace(address="$A$1:$F$50")
            )
        )
        for kind in ("trend", "scatterplot"):
            nav = self.app.trend_navigator if kind == "trend" else self.nav
            self.app.table_var.set("Measurements")
            self.app.save_trend_settings("Table", nav.export_settings(), kind)
            self.app.table_var.set("")
            self.app.source_range_var.set("Z1:Z2")
            self.app.load_trend_settings("Table", kind)
            self.assertEqual(self.app.table_var.get(), "Measurements")
            self.assertEqual(self.app.source_range_var.get(), "$A$1:$F$50")
        self.app.x_var.set("x")
        self.app.y_var.set("y")
        self.app.weight_var.set("(none)")
        self.app.settings_name_var.set("Table heat")
        self.app.save_settings()
        self.app.source_range_var.set("Z1:Z2")
        self.app.load_settings()
        self.assertEqual(self.app.source_range_var.get(), "$A$1:$F$50")

    def test_previews_update_and_dropdown_wheels_share_protection(self):
        trend = self.app.trend_navigator
        for canvas, variable in (
            (self.nav.cmap_preview, self.nav.options["colormap"]),
            (trend.cmap_preview, trend.colormap),
        ):
            before = canvas.itemcget(canvas.find_all()[0], "fill")
            variable.set("plasma")
            self.assertNotEqual(before, canvas.itemcget(canvas.find_all()[0], "fill"))
        for combo in (
            self.app.cmap_combo,
            trend.time_combo,
            self.nav.column_combos["z"],
        ):
            for sequence in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
                binding = combo.bind(sequence)
                self.assertTrue(binding)
                command = binding.split("[", 1)[1].split()[0]
                self.assertEqual(combo.tk.call(command, *["0"] * 19), "break")

    def test_invalid_scatter_and_missing_source_are_reported(self):
        self.nav.options["x"].set("missing")
        self.assertFalse(self.nav.plot_scatter())
        self.assertIn("source columns", self.app.status_var.get())
        self.nav.options["x"].set("x")
        self.app.save_trend_settings("Saved", self.nav.export_settings(), "scatterplot")
        self.book.sheets.remove(self.source)
        self.app.load_trend_settings("Saved", "scatterplot")
        self.assertIn("no longer exists", self.app.status_var.get())
        self.assertNotIn("Loaded", self.app.status_var.get())


if __name__ == "__main__":
    unittest.main()
