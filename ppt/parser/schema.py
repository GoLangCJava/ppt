"""
Pydantic Schemas for Declarative Presentation Specifications (DSL).
"""

from typing import List, Optional, Union, Literal, Dict, Any
from pydantic import BaseModel, Field


class CardSpec(BaseModel):
    title: str
    badge: Optional[str] = None
    subtitle: Optional[str] = None
    bullets: List[str] = Field(default_factory=list)
    tags: List[str] = Field(default_factory=list)
    highlight: bool = False


class ArchModuleSpec(BaseModel):
    title: str
    description: str
    tags: List[str] = Field(default_factory=list)


class ArchLayerSpec(BaseModel):
    layer_num: str
    name: str
    sub_en: str
    modules: List[ArchModuleSpec]
    protocol_connector: Optional[str] = None


class OptionSpec(BaseModel):
    name: str
    subtitle: str
    is_recommended: bool = False


class OptionCellSpec(BaseModel):
    score_symbol: str = "●"  # '●', '▲', '○'
    verdict: str
    bullets: List[str] = Field(default_factory=list)


class DimensionRowSpec(BaseModel):
    index: str
    title: str
    description: str
    cells: List[OptionCellSpec]


class RecommendationSpec(BaseModel):
    recommended_title: str = "推荐方案"
    recommended_text: str = ""
    tradeoff_title: str = "关键取舍"
    tradeoff_text: str = ""


class RoadmapPhaseSpec(BaseModel):
    phase_id: str
    duration_weeks: str
    title: str
    tasks: List[str]
    exit_criteria: str
    is_critical: bool = False


class EffortMetricSpec(BaseModel):
    big_stat: str
    sub_label: str


class KPISpec(BaseModel):
    value: str
    label: str
    description: str
    tag: Optional[str] = None
    highlight: bool = False


# Slide Types
class CoverSlideSpec(BaseModel):
    type: Literal["cover"] = "cover"
    title: str
    subtitle: Optional[str] = None
    client_tag: Optional[str] = None
    author: Optional[str] = None
    date: Optional[str] = None
    dark_mode: bool = False


class SectionSlideSpec(BaseModel):
    type: Literal["section"] = "section"
    section_num: str
    title: str
    subtitle: Optional[str] = None
    deck_tracker: Optional[str] = None
    dark_mode: bool = True


class CardsSlideSpec(BaseModel):
    type: Literal["cards"] = "cards"
    tracker: str = "Executive Summary"
    action_title: str
    lead_note: Optional[str] = None
    cards: List[CardSpec]
    takeaway: Optional[str] = None


class ArchSlideSpec(BaseModel):
    type: Literal["architecture"] = "architecture"
    tracker: str = "Architecture Blueprint"
    action_title: str
    lead_note: Optional[str] = None
    layers: List[ArchLayerSpec]
    takeaway: Optional[str] = None


class CompareSlideSpec(BaseModel):
    type: Literal["comparison"] = "comparison"
    tracker: str = "Solution Evaluation"
    action_title: str
    lead_note: Optional[str] = None
    options: List[OptionSpec]
    rows: List[DimensionRowSpec]
    takeaway: Optional[RecommendationSpec] = None


class RoadmapSlideSpec(BaseModel):
    type: Literal["roadmap"] = "roadmap"
    tracker: str = "Delivery Roadmap"
    action_title: str
    lead_note: Optional[str] = None
    phases: List[RoadmapPhaseSpec]
    traditional_effort: str = "320 人天"
    ai_effort: str = "195 人天"
    metrics: List[EffortMetricSpec] = Field(default_factory=list)
    summary_text: Optional[str] = None


class KPISlideSpec(BaseModel):
    type: Literal["kpis"] = "kpis"
    tracker: str = "Executive Metrics"
    action_title: str
    lead_note: Optional[str] = None
    kpis: List[KPISpec]
    takeaway: Optional[str] = None


SlideSpecUnion = Union[
    CoverSlideSpec,
    SectionSlideSpec,
    CardsSlideSpec,
    ArchSlideSpec,
    CompareSlideSpec,
    RoadmapSlideSpec,
    KPISlideSpec,
]


class DeckSpec(BaseModel):
    """Full Presentation Specification."""
    title: str
    theme: str = "consulting_blue"
    author: Optional[str] = None
    date: Optional[str] = None
    aspect_ratio: str = "16:9"
    slides: List[SlideSpecUnion]
