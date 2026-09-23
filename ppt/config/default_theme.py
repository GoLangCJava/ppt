"""
Preconfigured theme presets.
"""

from ppt.config.theme import ConsultingTheme, ColorPalette, FontTokens, DimensionTokens, hex_to_rgb


def get_consulting_blue_theme() -> ConsultingTheme:
    """Default high-end consulting theme based on China RIMS solution deck."""
    palette = ColorPalette(
        primary=hex_to_rgb("009CDE"),
        primary_dark=hex_to_rgb("005A9C"),
        primary_light=hex_to_rgb("EAF4FB"),
        secondary=hex_to_rgb("7C3AED"),
        accent_teal=hex_to_rgb("0284C7"),
        success=hex_to_rgb("10B981"),
        warning=hex_to_rgb("F59E0B"),
        danger=hex_to_rgb("EF4444"),
        bg_dark=hex_to_rgb("0F172A"),
        bg_light=hex_to_rgb("F8FAFC"),
        bg_card=hex_to_rgb("FFFFFF"),
        bg_tint=hex_to_rgb("F1F5F9"),
        border_subtle=hex_to_rgb("E2E8F0"),
        border_accent=hex_to_rgb("009CDE"),
        text_primary=hex_to_rgb("0F172A"),
        text_secondary=hex_to_rgb("334155"),
        text_muted=hex_to_rgb("64748B"),
        text_white=hex_to_rgb("FFFFFF"),
    )
    return ConsultingTheme(name="consulting_blue", colors=palette)


def get_executive_dark_theme() -> ConsultingTheme:
    """Executive dark theme for keynote and strategic summits."""
    palette = ColorPalette(
        primary=hex_to_rgb("38BDF8"),
        primary_dark=hex_to_rgb("0284C7"),
        primary_light=hex_to_rgb("1E293B"),
        secondary=hex_to_rgb("A855F7"),
        accent_teal=hex_to_rgb("2DD4BF"),
        success=hex_to_rgb("34D399"),
        warning=hex_to_rgb("FBBF24"),
        danger=hex_to_rgb("F87171"),
        bg_dark=hex_to_rgb("090D16"),
        bg_light=hex_to_rgb("0F172A"),
        bg_card=hex_to_rgb("1E293B"),
        bg_tint=hex_to_rgb("172033"),
        border_subtle=hex_to_rgb("334155"),
        border_accent=hex_to_rgb("38BDF8"),
        text_primary=hex_to_rgb("F8FAFC"),
        text_secondary=hex_to_rgb("CBD5E1"),
        text_muted=hex_to_rgb("94A3B8"),
        text_white=hex_to_rgb("FFFFFF"),
    )
    return ConsultingTheme(name="executive_dark", colors=palette)


def get_strategy_navy_theme() -> ConsultingTheme:
    """Classic strategy consulting navy theme (McKinsey / BCG style)."""
    palette = ColorPalette(
        primary=hex_to_rgb("1E3A8A"),
        primary_dark=hex_to_rgb("172554"),
        primary_light=hex_to_rgb("EFF6FF"),
        secondary=hex_to_rgb("D97706"),
        accent_teal=hex_to_rgb("2563EB"),
        success=hex_to_rgb("15803D"),
        warning=hex_to_rgb("B45309"),
        danger=hex_to_rgb("B91C1C"),
        bg_dark=hex_to_rgb("0B132B"),
        bg_light=hex_to_rgb("F8FAFC"),
        bg_card=hex_to_rgb("FFFFFF"),
        bg_tint=hex_to_rgb("F1F5F9"),
        border_subtle=hex_to_rgb("E2E8F0"),
        border_accent=hex_to_rgb("1E3A8A"),
        text_primary=hex_to_rgb("0F172A"),
        text_secondary=hex_to_rgb("334155"),
        text_muted=hex_to_rgb("64748B"),
        text_white=hex_to_rgb("FFFFFF"),
    )
    return ConsultingTheme(name="strategy_navy", colors=palette)


THEMES = {
    "consulting_blue": get_consulting_blue_theme,
    "executive_dark": get_executive_dark_theme,
    "strategy_navy": get_strategy_navy_theme,
}


def load_theme(name: str = "consulting_blue") -> ConsultingTheme:
    """Load theme by name."""
    factory = THEMES.get(name, get_consulting_blue_theme)
    return factory()
