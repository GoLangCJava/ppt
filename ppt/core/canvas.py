"""
Slide canvas management and grid layout engine.
"""

from typing import List, Tuple, Optional, Dict, Any
from pptx.slide import Slide
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.shapes.autoshape import Shape
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN

from ppt.config.theme import ConsultingTheme, DimensionTokens
from ppt.core.shapes import add_rounded_card, add_rect_card, add_badge, add_divider_line
from ppt.core.typography import format_run, add_formatted_paragraph, set_text_frame_margins


class SlideCanvas:
    """
    High-level canvas wrapper for drawing consulting slides with strict grid alignments.
    """

    def __init__(self, slide: Slide, theme: ConsultingTheme):
        self.slide = slide
        self.theme = theme
        self.dim = theme.dimensions

        # Standard consulting layout vertical zones
        self.margin_left = self.dim.margin_left_inches
        self.margin_right = self.dim.margin_right_inches
        self.margin_top = self.dim.margin_top_inches
        self.margin_bottom = self.dim.margin_bottom_inches

        self.content_width = self.dim.slide_width_inches - self.margin_left - self.margin_right
        self.slide_height = self.dim.slide_height_inches
        self.slide_width = self.dim.slide_width_inches

        # Header ends at top: ~1.30 inches, leaving body from 1.35 to 4.90
        self.header_height = 0.95
        self.body_top = self.margin_top + self.header_height + 0.10
        self.body_bottom = self.slide_height - self.margin_bottom - 0.50
        self.body_height = self.body_bottom - self.body_top

    def set_background(self, color: RGBColor) -> Shape:
        """Set full slide background color."""
        bg = self.slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(0),
            Inches(0),
            Inches(self.slide_width),
            Inches(self.slide_height),
        )
        bg.fill.solid()
        bg.fill.fore_color.rgb = color
        bg.line.fill.background()
        bg.name = "SlideBackground"
        # Move to back if needed
        return bg

    def add_header(
        self,
        tracker: str,
        action_title: str,
        lead_note: Optional[str] = None,
        tracker_badge: Optional[str] = None,
    ) -> Dict[str, Shape]:
        """
        Add a standard consulting header:
        - Tracker line (e.g. 'China RIMS Solution Proposal | BTS China')
        - Action Title (bold takeaway title)
        - Lead note (highlighted summary sentence explaining the slide)
        """
        shapes = {}

        # 1. Tracker / Kicker at the very top
        tracker_box = self.slide.shapes.add_textbox(
            Inches(self.margin_left),
            Inches(self.margin_top),
            Inches(self.content_width),
            Inches(0.24),
        )
        tracker_box.name = "HeaderTracker"
        tf_tracker = tracker_box.text_frame
        set_text_frame_margins(tf_tracker, top=0, bottom=0, left=0, right=0)
        p_track = tf_tracker.paragraphs[0]
        r_track = p_track.add_run()
        format_run(
            r_track,
            tracker.upper(),
            self.theme.fonts.font_title,
            self.theme.fonts.size_caption,
            self.theme.colors.text_muted,
            bold=False,
        )
        shapes["tracker"] = tracker_box

        # 2. Main Action Title
        title_box = self.slide.shapes.add_textbox(
            Inches(self.margin_left),
            Inches(self.margin_top + 0.22),
            Inches(self.content_width),
            Inches(0.40),
        )
        title_box.name = "HeaderActionTitle"
        tf_title = title_box.text_frame
        set_text_frame_margins(tf_title, top=0, bottom=0, left=0, right=0)
        p_title = tf_title.paragraphs[0]
        r_title = p_title.add_run()
        format_run(
            r_title,
            action_title,
            self.theme.fonts.font_title,
            self.theme.fonts.size_slide_title,
            self.theme.colors.text_primary,
            bold=True,
        )
        shapes["title"] = title_box

        # 3. Lead note / Subtitle banner
        if lead_note:
            lead_box = self.slide.shapes.add_textbox(
                Inches(self.margin_left),
                Inches(self.margin_top + 0.65),
                Inches(self.content_width),
                Inches(0.32),
            )
            lead_box.name = "HeaderLeadNote"
            tf_lead = lead_box.text_frame
            set_text_frame_margins(tf_lead, top=0, bottom=0, left=0, right=0)
            p_lead = tf_lead.paragraphs[0]
            r_lead = p_lead.add_run()
            format_run(
                r_lead,
                lead_note,
                self.theme.fonts.font_body,
                self.theme.fonts.size_action_title,
                self.theme.colors.primary_dark,
                bold=False,
            )
            shapes["lead"] = lead_box

        # 4. Subtle separator line
        div = add_divider_line(
            self.slide,
            self.margin_left,
            self.margin_top + self.header_height + 0.05,
            self.content_width,
            self.theme.colors.border_subtle,
            width_pt=0.5,
        )
        shapes["divider"] = div
        return shapes

    def add_takeaway_footer(
        self,
        takeaway_text: str,
        label: str = "核心结论",
        bg_color: Optional[RGBColor] = None,
        border_color: Optional[RGBColor] = None,
    ) -> Shape:
        """Add an executive takeaway card at the bottom of the slide."""
        foot_top = self.slide_height - self.margin_bottom - 0.42
        foot_height = 0.40
        foot_width = self.content_width

        bg = bg_color or self.theme.colors.primary_light
        border = border_color or self.theme.colors.primary

        card = add_rounded_card(
            self.slide,
            self.margin_left,
            foot_top,
            foot_width,
            foot_height,
            fill_color=bg,
            border_color=border,
            border_width_pt=0.75,
            name="TakeawayFooter",
        )
        tf = card.text_frame
        set_text_frame_margins(tf, top=0.06, bottom=0.06, left=0.15, right=0.15)
        p = tf.paragraphs[0]

        r_badge = p.add_run()
        format_run(
            r_badge,
            f"【{label}】 ",
            self.theme.fonts.font_body,
            self.theme.fonts.size_body,
            self.theme.colors.primary_dark,
            bold=True,
        )

        r_text = p.add_run()
        format_run(
            r_text,
            takeaway_text,
            self.theme.fonts.font_body,
            self.theme.fonts.size_body,
            self.theme.colors.text_primary,
            bold=False,
        )
        return card

    def compute_columns(
        self,
        count: int,
        gap: float = 0.15,
        top: Optional[float] = None,
        height: Optional[float] = None,
    ) -> List[Tuple[float, float, float, float]]:
        """
        Compute bounding boxes `(left, top, width, height)` for N equal columns.
        """
        y = top if top is not None else self.body_top
        h = height if height is not None else self.body_height

        total_gap = gap * (count - 1)
        col_width = (self.content_width - total_gap) / count

        cols = []
        for i in range(count):
            col_left = self.margin_left + i * (col_width + gap)
            cols.append((col_left, y, col_width, h))
        return cols

    def compute_grid(
        self,
        rows: int,
        cols: int,
        gap_x: float = 0.15,
        gap_y: float = 0.15,
        top: Optional[float] = None,
        height: Optional[float] = None,
    ) -> List[List[Tuple[float, float, float, float]]]:
        """
        Compute bounding boxes for a 2D grid `[row][col] -> (left, top, width, height)`.
        """
        y_start = top if top is not None else self.body_top
        total_h = height if height is not None else self.body_height

        col_w = (self.content_width - gap_x * (cols - 1)) / cols
        row_h = (total_h - gap_y * (rows - 1)) / rows

        grid = []
        for r in range(rows):
            row_items = []
            curr_y = y_start + r * (row_h + gap_y)
            for c in range(cols):
                curr_x = self.margin_left + c * (col_w + gap_x)
                row_items.append((curr_x, curr_y, col_w, row_h))
            grid.append(row_items)
        return grid
