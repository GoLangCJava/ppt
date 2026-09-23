"""
Standard multi-card structured analysis slide template.
"""

from typing import List, Optional
from ppt.core.canvas import SlideCanvas
from ppt.components.card_grid import CardItem, render_card_grid


def render_card_slide(
    canvas: SlideCanvas,
    tracker: str,
    action_title: str,
    cards: List[CardItem],
    lead_note: Optional[str] = None,
    takeaway: Optional[str] = None,
):
    """
    Render a standard consulting slide with Header, 2-5 Cards Grid, and optional Takeaway.
    """
    canvas.add_header(tracker=tracker, action_title=action_title, lead_note=lead_note)

    # Compute vertical space
    if takeaway:
        canvas.add_takeaway_footer(takeaway)
        body_h = canvas.body_height - 0.45
    else:
        body_h = canvas.body_height

    render_card_grid(canvas, cards, top=canvas.body_top, height=body_h)
