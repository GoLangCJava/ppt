"""
Consulting Slide Components Library.
"""

from ppt.components.card_grid import CardItem, render_card_grid
from ppt.components.architecture_view import (
    ArchitectureModule,
    ArchitectureLayer,
    render_architecture_view,
)
from ppt.components.comparison_table import (
    ComparisonOption,
    DimensionRow,
    OptionCell,
    RecommendationTakeaway,
    render_comparison_table,
)
from ppt.components.roadmap_timeline import (
    RoadmapPhase,
    EffortMetric,
    render_roadmap_timeline,
)
from ppt.components.kpi_summary import KPICard, render_kpi_grid
from ppt.components.takeaway_box import render_takeaway_box

__all__ = [
    "CardItem",
    "render_card_grid",
    "ArchitectureModule",
    "ArchitectureLayer",
    "render_architecture_view",
    "ComparisonOption",
    "DimensionRow",
    "OptionCell",
    "RecommendationTakeaway",
    "render_comparison_table",
    "RoadmapPhase",
    "EffortMetric",
    "render_roadmap_timeline",
    "KPICard",
    "render_kpi_grid",
    "render_takeaway_box",
]
