"""
Theme and Design Tokens for Consulting Presentations.
"""

from dataclasses import dataclass, field
from typing import Dict, Any, Optional
from pptx.dml.color import RGBColor


def hex_to_rgb(hex_str: str) -> RGBColor:
    """Convert hex color string (e.g. '#009CDE' or '009CDE') to RGBColor."""
    clean = hex_str.lstrip("#")
    if len(clean) != 6:
        raise ValueError(f"Invalid hex color string: {hex_str}")
    r = int(clean[0:2], 16)
    g = int(clean[2:4], 16)
    b = int(clean[4:6], 16)
    return RGBColor(r, g, b)


@dataclass
class ColorPalette:
    """Core consulting color palette."""
    primary: RGBColor = field(default_factory=lambda: hex_to_rgb("009CDE"))
    primary_dark: RGBColor = field(default_factory=lambda: hex_to_rgb("005A9C"))
    primary_light: RGBColor = field(default_factory=lambda: hex_to_rgb("EAF4FB"))
    secondary: RGBColor = field(default_factory=lambda: hex_to_rgb("7C3AED"))
    accent_teal: RGBColor = field(default_factory=lambda: hex_to_rgb("0284C7"))
    success: RGBColor = field(default_factory=lambda: hex_to_rgb("10B981"))
    warning: RGBColor = field(default_factory=lambda: hex_to_rgb("F59E0B"))
    danger: RGBColor = field(default_factory=lambda: hex_to_rgb("EF4444"))
    
    # Neutral backgrounds & surfaces
    bg_dark: RGBColor = field(default_factory=lambda: hex_to_rgb("0F172A"))
    bg_light: RGBColor = field(default_factory=lambda: hex_to_rgb("F8FAFC"))
    bg_card: RGBColor = field(default_factory=lambda: hex_to_rgb("FFFFFF"))
    bg_tint: RGBColor = field(default_factory=lambda: hex_to_rgb("F1F5F9"))
    
    # Borders
    border_subtle: RGBColor = field(default_factory=lambda: hex_to_rgb("E2E8F0"))
    border_accent: RGBColor = field(default_factory=lambda: hex_to_rgb("009CDE"))
    
    # Text colors
    text_primary: RGBColor = field(default_factory=lambda: hex_to_rgb("0F172A"))
    text_secondary: RGBColor = field(default_factory=lambda: hex_to_rgb("475569"))
    text_muted: RGBColor = field(default_factory=lambda: hex_to_rgb("64748B"))
    text_white: RGBColor = field(default_factory=lambda: hex_to_rgb("FFFFFF"))


@dataclass
class FontTokens:
    """Typography tokens for consulting slides."""
    font_title: str = "Georgia"
    font_body: str = "Microsoft YaHei"
    font_mono: str = "Consolas"
    
    # Size scale in points (pt)
    size_hero: float = 24.0
    size_slide_title: float = 16.0
    size_action_title: float = 12.0
    size_section_title: float = 11.0
    size_card_title: float = 10.0
    size_body: float = 8.5
    size_caption: float = 7.5
    size_badge: float = 7.0


@dataclass
class DimensionTokens:
    """Slide geometry tokens."""
    slide_width_inches: float = 10.0
    slide_height_inches: float = 5.625
    margin_left_inches: float = 0.55
    margin_right_inches: float = 0.55
    margin_top_inches: float = 0.35
    margin_bottom_inches: float = 0.35


@dataclass
class ConsultingTheme:
    """Comprehensive design system theme."""
    name: str = "consulting_blue"
    colors: ColorPalette = field(default_factory=ColorPalette)
    fonts: FontTokens = field(default_factory=FontTokens)
    dimensions: DimensionTokens = field(default_factory=DimensionTokens)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "ConsultingTheme":
        """Build theme from config dict."""
        theme = cls()
        if "name" in data:
            theme.name = data["name"]
        
        colors_data = data.get("colors", {})
        for k, v in colors_data.items():
            if hasattr(theme.colors, k) and isinstance(v, str):
                setattr(theme.colors, k, hex_to_rgb(v))
                
        fonts_data = data.get("fonts", {})
        for k, v in fonts_data.items():
            if hasattr(theme.fonts, k):
                setattr(theme.fonts, k, v)
                
        return theme


# Common color helper alias
Color = hex_to_rgb
