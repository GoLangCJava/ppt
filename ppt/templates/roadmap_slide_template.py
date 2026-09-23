"""
Roadmap and implementation timeline slide template.
"""

from typing import List, Optional
from ppt.core.canvas import SlideCanvas
from ppt.components.roadmap_timeline import (
    RoadmapPhase,
    EffortMetric,
    render_roadmap_timeline,
)


def render_roadmap_slide(
    canvas: SlideCanvas,
    tracker: str,
    action_title: str,
    phases: List[RoadmapPhase],
    traditional_effort: str = "320 人天",
    ai_effort: str = "195 人天",
    metrics: Optional[List[EffortMetric]] = None,
    summary_text: Optional[str] = None,
    lead_note: Optional[str] = None,
):
    """
    Render a roadmap slide with Header, Gantt Phases, Exit Criteria, and Effort KPIs.
    """
    canvas.add_header(tracker=tracker, action_title=action_title, lead_note=lead_note)
    render_roadmap_timeline(
        canvas,
        phases=phases,
        traditional_effort=traditional_effort,
        ai_effort=ai_effort,
        metrics=metrics,
        summary_text=summary_text,
        top=canvas.body_top,
        height=canvas.body_height,
    )
