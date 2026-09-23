"""
Architecture blueprint slide template.
"""

from typing import List, Optional
from ppt.core.canvas import SlideCanvas
from ppt.components.architecture_view import ArchitectureLayer, render_architecture_view


def render_architecture_slide(
    canvas: SlideCanvas,
    tracker: str,
    action_title: str,
    layers: List[ArchitectureLayer],
    lead_note: Optional[str] = None,
    takeaway: Optional[str] = None,
):
    """
    Render an enterprise architecture slide with Header, N-tier Stack, and optional Takeaway.
    """
    canvas.add_header(tracker=tracker, action_title=action_title, lead_note=lead_note)

    if takeaway:
        canvas.add_takeaway_footer(takeaway)
        body_h = canvas.body_height - 0.45
    else:
        body_h = canvas.body_height

    render_architecture_view(canvas, layers, top=canvas.body_top, height=body_h)
