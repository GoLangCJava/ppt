"""
Evaluation and comparison slide template.
"""

from typing import List, Optional
from ppt.core.canvas import SlideCanvas
from ppt.components.comparison_table import (
    ComparisonOption,
    DimensionRow,
    RecommendationTakeaway,
    render_comparison_table,
)


def render_compare_slide(
    canvas: SlideCanvas,
    tracker: str,
    action_title: str,
    options: List[ComparisonOption],
    rows: List[DimensionRow],
    lead_note: Optional[str] = None,
    takeaway: Optional[RecommendationTakeaway] = None,
):
    """
    Render a consulting comparison matrix slide.
    """
    canvas.add_header(tracker=tracker, action_title=action_title, lead_note=lead_note)
    render_comparison_table(
        canvas,
        options=options,
        rows=rows,
        takeaway=takeaway,
        top=canvas.body_top,
        height=canvas.body_height,
    )
