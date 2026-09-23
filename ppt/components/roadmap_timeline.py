"""
Delivery Roadmap & Timeline component.
Faithfully recreates Slide 43 (Implementation Roadmap, Exit Criteria & Effort Comparison).
"""

from dataclasses import dataclass, field
from typing import List, Optional
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

from ppt.core.canvas import SlideCanvas
from ppt.core.shapes import add_rounded_card, add_rect_card, add_badge
from ppt.core.typography import format_run, set_text_frame_margins


@dataclass
class RoadmapPhase:
    """A single phase on the implementation roadmap."""
    phase_id: str          # e.g. 'STEP 1+2'
    duration_weeks: str    # e.g. '2 Weeks'
    title: str             # e.g. '需求分析与架构设计'
    tasks: List[str]       # bullet tasks
    exit_criteria: str     # e.g. '设计评审通过 + 接口冻结'
    is_critical: bool = False


@dataclass
class EffortMetric:
    """Quantitative effort or ROI metric."""
    big_stat: str          # e.g. '124 人天'
    sub_label: str         # e.g. '~40% Effort Saving'
    color: Optional[RGBColor] = None


def render_roadmap_timeline(
    canvas: SlideCanvas,
    phases: List[RoadmapPhase],
    traditional_effort: str = "320 人天",
    ai_effort: str = "195 人天",
    metrics: Optional[List[EffortMetric]] = None,
    summary_text: Optional[str] = None,
    top: Optional[float] = None,
    height: Optional[float] = None,
):
    """
    Render a consulting delivery roadmap with phase cards, exit criteria, and an effort comparison dashboard.
    """
    theme = canvas.theme
    start_y = top if top is not None else canvas.body_top
    avail_h = height if height is not None else canvas.body_height

    num_phases = len(phases)
    if num_phases == 0:
        return

    # Split vertical space: Upper 55% for Roadmap Gantt, Lower 45% for Effort & KPIs
    roadmap_h = avail_h * 0.52
    effort_h = avail_h - roadmap_h - 0.10

    # 1. Render Upper Roadmap Phases
    phase_gap = 0.10
    phase_w = (canvas.content_width - (phase_gap * (num_phases - 1))) / num_phases

    curr_x = canvas.margin_left
    for p_idx, phase in enumerate(phases):
        is_crit = phase.is_critical
        bg_col = theme.colors.primary_light if is_crit else theme.colors.bg_card
        border_col = theme.colors.primary if is_crit else theme.colors.border_subtle

        # Phase Card Box
        card_h = roadmap_h - 0.45
        p_card = add_rounded_card(
            canvas.slide,
            curr_x,
            start_y,
            phase_w,
            card_h,
            fill_color=bg_col,
            border_color=border_col,
            border_width_pt=1.0,
            name=f"RoadmapPhase_{p_idx+1}",
        )
        tf_p = p_card.text_frame
        set_text_frame_margins(tf_p, top=0.06, bottom=0.04, left=0.08, right=0.08)

        # Duration Badge / Title
        p_dur = tf_p.paragraphs[0]
        r_dur = p_dur.add_run()
        format_run(
            r_dur,
            f"{phase.duration_weeks} ｜ {phase.phase_id}\n",
            theme.fonts.font_title,
            theme.fonts.size_body,
            theme.colors.primary_dark,
            bold=True,
        )

        p_title = tf_p.add_paragraph()
        p_title.space_after = Pt(2.0)
        r_title = p_title.add_run()
        format_run(
            r_title,
            phase.title,
            theme.fonts.font_title,
            theme.fonts.size_card_title,
            theme.colors.text_primary,
            bold=True,
        )

        for task in phase.tasks:
            p_task = tf_p.add_paragraph()
            p_task.space_before = Pt(1.0)
            r_t = p_task.add_run()
            format_run(
                r_t,
                f"• {task}",
                theme.fonts.font_body,
                6.5,
                theme.colors.text_secondary,
                bold=False,
            )

        # Exit Criteria Box underneath phase card
        crit_top = start_y + card_h + 0.05
        crit_box = add_rounded_card(
            canvas.slide,
            curr_x,
            crit_top,
            phase_w,
            0.36,
            fill_color=theme.colors.bg_tint,
            border_color=theme.colors.border_subtle,
            border_width_pt=0.5,
            name=f"ExitCriteria_{p_idx+1}",
        )
        tf_c = crit_box.text_frame
        set_text_frame_margins(tf_c, top=0.03, bottom=0.03, left=0.06, right=0.06)
        p_c = tf_c.paragraphs[0]
        r_c1 = p_c.add_run()
        format_run(
            r_c1,
            "退出标准: ",
            theme.fonts.font_body,
            6.5,
            theme.colors.primary_dark,
            bold=True,
        )
        r_c2 = p_c.add_run()
        format_run(
            r_c2,
            phase.exit_criteria,
            theme.fonts.font_body,
            6.5,
            theme.colors.text_secondary,
            bold=False,
        )

        curr_x += phase_w + phase_gap

    # 2. Render Lower Effort Comparison & KPI Dashboard
    lower_y = start_y + roadmap_h + 0.10
    left_w = canvas.content_width * 0.58
    right_w = canvas.content_width - left_w - 0.15
    right_x = canvas.margin_left + left_w + 0.15

    # Left: Comparison Bars & Summary
    effort_box = add_rounded_card(
        canvas.slide,
        canvas.margin_left,
        lower_y,
        left_w,
        effort_h,
        fill_color=theme.colors.bg_card,
        border_color=theme.colors.border_subtle,
        border_width_pt=0.75,
        name="EffortComparisonCard",
    )
    tf_e = effort_box.text_frame
    set_text_frame_margins(tf_e, top=0.08, bottom=0.06, left=0.12, right=0.12)

    p_et = tf_e.paragraphs[0]
    r_et = p_et.add_run()
    format_run(
        r_et,
        "实施投入对比与交付效率 (Effort Comparison)\n",
        theme.fonts.font_title,
        theme.fonts.size_card_title,
        theme.colors.primary_dark,
        bold=True,
    )

    p_bars = tf_e.add_paragraph()
    p_bars.space_before = Pt(3.0)
    r_b1 = p_bars.add_run()
    format_run(
        r_b1,
        f"传统 SLC 模式投入： {traditional_effort}\n",
        theme.fonts.font_body,
        theme.fonts.size_body,
        theme.colors.text_secondary,
        bold=False,
    )
    r_b2 = p_bars.add_run()
    format_run(
        r_b2,
        f"SLC + AI 模式投入： {ai_effort}  (工期提效 ~40%)\n",
        theme.fonts.font_body,
        theme.fonts.size_body,
        theme.colors.success,
        bold=True,
    )

    if summary_text:
        p_sum = tf_e.add_paragraph()
        p_sum.space_before = Pt(3.0)
        r_sum = p_sum.add_run()
        format_run(
            r_sum,
            f"💡 {summary_text}",
            theme.fonts.font_body,
            theme.fonts.size_caption,
            theme.colors.text_muted,
            bold=False,
        )

    # Right: KPI Highlight Badges
    if metrics:
        num_m = len(metrics)
        m_h = (effort_h - (0.06 * (num_m - 1))) / num_m
        curr_m_y = lower_y

        for m_idx, m in enumerate(metrics):
            m_col = m.color or theme.colors.primary
            m_card = add_rounded_card(
                canvas.slide,
                right_x,
                curr_m_y,
                right_w,
                m_h,
                fill_color=theme.colors.primary_light if m_idx == 0 else theme.colors.bg_tint,
                border_color=m_col,
                border_width_pt=1.0 if m_idx == 0 else 0.5,
                name=f"KPIMetric_{m_idx+1}",
            )
            tf_m = m_card.text_frame
            set_text_frame_margins(tf_m, top=0.04, bottom=0.04, left=0.10, right=0.10)
            p_m = tf_m.paragraphs[0]
            p_m.alignment = PP_ALIGN.CENTER

            r_big = p_m.add_run()
            format_run(
                r_big,
                f"{m.big_stat}  ",
                theme.fonts.font_title,
                theme.fonts.size_action_title,
                m_col,
                bold=True,
            )

            r_sub = p_m.add_run()
            format_run(
                r_sub,
                m.sub_label,
                theme.fonts.font_body,
                theme.fonts.size_caption,
                theme.colors.text_secondary,
                bold=False,
            )

            curr_m_y += m_h + 0.06
