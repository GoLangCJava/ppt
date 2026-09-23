"""
Consulting PPT Generator
~~~~~~~~~~~~~~~~~~~~~~~~
A Python library for generating high-density, executive-level consulting presentations (McKinsey / BCG / Bain style) using python-pptx.
"""

__version__ = "0.1.0"

from ppt.core.canvas import SlideCanvas
from ppt.config.theme import ConsultingTheme, Color
from ppt.parser.builder import DeckBuilder

__all__ = ["SlideCanvas", "ConsultingTheme", "Color", "DeckBuilder"]
