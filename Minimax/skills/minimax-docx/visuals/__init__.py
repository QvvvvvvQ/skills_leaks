"""
Visuals - Document styling and background generation.
"""

from .palettes import REGISTRY, fetch_palette, fetch_typography, fetch_ornamental
from .backdrops import render_backdrops

__all__ = [
    "REGISTRY",
    "fetch_palette",
    "fetch_typography",
    "fetch_ornamental",
    "render_backdrops",
]
