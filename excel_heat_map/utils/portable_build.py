"""Runtime dependency collection and an Excel-independent frozen-app self-test."""

import json
import os
from pathlib import Path
import shutil
import sys
import traceback


def collect_runtime(cache: Path) -> list[str]:
    import pefile

    base = Path(sys.base_prefix)
    search_paths = (base / "DLLs", base / "Library" / "bin", base)
    system = Path(os.environ["SystemRoot"]) / "System32"
    pending = list((base / "DLLs").glob("*.pyd")) + list(base.glob("python*.dll"))
    visited: set[Path] = set()
    libraries: dict[str, Path] = {}
    while pending:
        binary = pending.pop()
        if binary in visited:
            continue
        visited.add(binary)
        with pefile.PE(str(binary), fast_load=True) as image:
            image.parse_data_directories(
                directories=[
                    pefile.DIRECTORY_ENTRY["IMAGE_DIRECTORY_ENTRY_IMPORT"],
                    pefile.DIRECTORY_ENTRY["IMAGE_DIRECTORY_ENTRY_DELAY_IMPORT"],
                ]
            )
            imports = [
                entry.dll.decode("ascii")
                for directory in (
                    "DIRECTORY_ENTRY_IMPORT",
                    "DIRECTORY_ENTRY_DELAY_IMPORT",
                )
                for entry in getattr(image, directory, [])
            ]
        for name in imports:
            if name.lower().startswith(("api-ms-win-", "ext-ms-win-")):
                continue
            dependency = next(
                (
                    folder / name
                    for folder in (binary.parent, *search_paths)
                    if (folder / name).is_file()
                ),
                None,
            )
            if dependency is None:
                if (system / name).is_file():
                    continue
                raise FileNotFoundError(f"Cannot resolve {name}, required by {binary}")
            key = name.casefold()
            if key not in libraries:
                libraries[key] = dependency
                pending.append(dependency)

    cache.mkdir(parents=True, exist_ok=True)
    names = []
    for dependency in sorted(libraries.values()):
        destination = cache / dependency.name
        # Refresh the cached DLL when the underlying Python installation changes.
        if (
            not destination.exists()
            or dependency.read_bytes() != destination.read_bytes()
        ):
            shutil.copy2(dependency, destination)
        names.append(dependency.name)
    (cache / "manifest.json").write_text(
        json.dumps(
            {"python": sys.version, "base": str(base), "files": names}, indent=2
        ),
        encoding="utf-8",
    )
    return names


def run_self_test(report_path: Path) -> int:
    report: dict = {"ok": False, "frozen": bool(getattr(sys, "frozen", False))}
    root = None
    try:
        import tkinter as tk

        import matplotlib
        import numpy as np
        import pandas as pd
        import pythoncom
        import pywintypes
        import win32clipboard
        import xlwings

        from excel_heat_map.ui.heatmap_ui import HAS_WINDOWS, HeatMapApp
        from excel_heat_map.utils.clipboard import figure_to_dib

        if not HAS_WINDOWS:
            raise RuntimeError(
                "The Excel/Windows COM dependencies could not be loaded."
            )
        root = tk.Tk()
        root.withdraw()
        app = HeatMapApp(root, connect_excel=False)
        frame = pd.DataFrame(
            {
                "time": pd.date_range("2026-01-01", periods=3),
                "x": [1.0, 2.0, 3.0],
                "y": [2.0, 4.0, 6.0],
                "z": [3.0, 1.0, 2.0],
                "category": ["A", "B", "A"],
            }
        )
        scatter = app.scatterplot_navigator
        scatter.data_provider = lambda: frame.copy()
        scatter.set_columns(list(frame.columns), frame)
        scatter.options["x"].set("x")
        scatter.options["y"].set("y")
        scatter.options["color"].set("category")
        for z in ("(None)", "z"):
            scatter.options["z"].set(z)
            if not scatter.plot_scatter():
                raise RuntimeError(app.status_var.get())
        trend = app.trend_navigator
        trend.data_provider = lambda: frame.copy()
        trend.set_columns(list(frame.columns), frame)
        if not trend.plot_trends():
            raise RuntimeError(app.status_var.get())
        if not figure_to_dib(scatter.figure):
            raise RuntimeError("Pillow/Matplotlib image export produced no data.")
        root.update_idletasks()
        report.update(
            ok=True,
            executable=sys.executable,
            windows_modules=[
                module.__name__ for module in (pythoncom, pywintypes, win32clipboard)
            ],
            versions={
                "numpy": np.__version__,
                "pandas": pd.__version__,
                "matplotlib": matplotlib.__version__,
                "xlwings": xlwings.__version__,
            },
            checks=[
                "Windows COM imports",
                "Tk UI",
                "2D scatter",
                "3D scatter",
                "trend graph",
                "Pillow image export",
            ],
        )
    except Exception:
        # A windowed executable has no console; preserve the complete fatal error.
        report["error"] = traceback.format_exc()
    finally:
        if root is not None:
            root.destroy()
    report_path.write_text(json.dumps(report, indent=2), encoding="utf-8")
    return 0 if report["ok"] else 1


if __name__ == "__main__":
    files = collect_runtime(Path(sys.argv[1]))
    print(f"Cached {len(files)} Python runtime DLLs.")
