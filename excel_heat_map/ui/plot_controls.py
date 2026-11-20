"""Controls shared by the graph tabs."""

import tkinter as tk

import matplotlib
from matplotlib.colors import to_hex


def protect_combobox_wheel(combo):
    for sequence in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
        combo.bind(sequence, lambda _event: "break")


def draw_colormap_preview(canvas, name, reverse=False):
    cmap = matplotlib.colormaps[name]
    if reverse:
        cmap = cmap.reversed()
    canvas.delete("all")
    width = max(2, canvas.winfo_width(), int(canvas.cget("width")))
    for index in range(width):
        canvas.create_line(index, 0, index, 18, fill=to_hex(cmap(index / (width - 1))))


def colormap_preview(parent, variable):
    canvas = tk.Canvas(parent, height=18, width=400, highlightthickness=0)
    canvas.pack(fill="x", pady=3)

    def redraw(*_args):
        draw_colormap_preview(canvas, variable.get())

    variable.trace_add("write", redraw)
    canvas.bind("<Configure>", redraw)
    redraw()
    return canvas
