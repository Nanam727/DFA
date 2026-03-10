"""Reusable placeholder alignment icons."""

from __future__ import annotations

from typing import Dict

import customtkinter as ctk
from PIL import Image, ImageDraw


def _draw_icon(kind: str, size: int) -> Image.Image:
    img = Image.new("RGBA", (size, size), (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)
    color = (58, 58, 58, 255)
    line_h = max(2, size // 14)
    gap = max(2, size // 8)
    top = max(2, size // 6)
    left_pad = max(2, size // 8)
    right_pad = max(2, size // 8)

    widths = {
        "left": [0.9, 0.7, 0.85, 0.6],
        "center": [0.75, 0.55, 0.8, 0.5],
        "right": [0.9, 0.7, 0.85, 0.6],
        "justify": [0.9, 0.9, 0.9, 0.75],
    }.get(kind, [0.85, 0.7, 0.8, 0.65])

    for idx, ratio in enumerate(widths):
        y = top + idx * gap
        usable = size - left_pad - right_pad
        seg = int(usable * ratio)
        if kind == "center":
            x1 = (size - seg) // 2
            x2 = x1 + seg
        elif kind == "right":
            x2 = size - right_pad
            x1 = x2 - seg
        else:
            x1 = left_pad
            x2 = x1 + seg
        draw.rectangle([x1, y, x2, y + line_h], fill=color)

    return img


def create_align_icons(size: int = 16) -> Dict[str, ctk.CTkImage]:
    """Create placeholder icons for align buttons."""
    return {
        "left": ctk.CTkImage(light_image=_draw_icon("left", size), size=(size, size)),
        "center": ctk.CTkImage(light_image=_draw_icon("center", size), size=(size, size)),
        "right": ctk.CTkImage(light_image=_draw_icon("right", size), size=(size, size)),
        "justify": ctk.CTkImage(light_image=_draw_icon("justify", size), size=(size, size)),
    }
