"""
Slide Templates Library.
"""

from ppt.templates.cover_template import render_cover_slide
from ppt.templates.section_template import render_section_slide
from ppt.templates.card_slide_template import render_card_slide
from ppt.templates.arch_slide_template import render_architecture_slide
from ppt.templates.compare_slide_template import render_compare_slide
from ppt.templates.roadmap_slide_template import render_roadmap_slide

__all__ = [
    "render_cover_slide",
    "render_section_slide",
    "render_card_slide",
    "render_architecture_slide",
    "render_compare_slide",
    "render_roadmap_slide",
]
