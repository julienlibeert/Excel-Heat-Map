"""Simple entry point for the Excel Heat Map application.

This module initializes the Tkinter root window and launches the HeatMapApp UI.
"""

import argparse
from pathlib import Path
import tkinter as tk
import matplotlib

matplotlib.use("TkAgg")

from excel_heat_map.ui.heatmap_ui import HeatMapApp


def main():
    """Launch the application."""
    parser = argparse.ArgumentParser(description="Excel Heat Map")
    parser.add_argument("--self-test", type=Path, help=argparse.SUPPRESS)
    arguments = parser.parse_args()
    if arguments.self_test is not None:
        from excel_heat_map.utils.portable_build import run_self_test

        raise SystemExit(run_self_test(arguments.self_test))
    root = tk.Tk()
    HeatMapApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
