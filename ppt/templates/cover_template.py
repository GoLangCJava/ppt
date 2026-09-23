"""
Cover slide templates (Dark Executive & Minimalist Light).
"""

from typing import Optional
from pptx.slide import Slide
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

from ppt.core.canvas import SlideCanvas
from ppt.core.shapes import add_rounded_card, add_badge, add_rect_card
from ppt.core.typography import format_run, set_text_frame_margins


def render_cover_slide(
    canvas: SlideCanvas,
    title: str,
    subtitle: Optional[str] = None,
    client_tag: Optional[str] = None,
    author: Optional[str] = None,
    date_str: Optional[str] = None,
    dark_mode: bool = False,
):
    """
    Render an executive cover slide matching consulting aesthetic.
    """
    theme = canvas.theme

    if dark_mode:
        canvas.set_background(theme.colors.bg_dark)
        text_title_col = theme.colors.text_white
        text_sub_col = theme.colors.primary_light
        text_muted_col = theme.colors.text_muted
    else:
        canvas.set_background(theme.colors.bg_card)
        text_title_col = theme.colors.text_primary
        text_sub_col = theme.colors.primary_dark
        text_muted_col = theme.colors.text_muted

    # Accent decorative bar on top or left
    bar_width = 0.15
    bar = add_rect_card(
        canvas.slide,
        canvas.margin_left,
        1.20,
        bar_width,
        2.60,
        fill_color=theme.colors.primary,
        border_color=None,
        name="CoverAccentBar",
    )

    # Content box next to bar
    content_x = canvas.margin_left + bar_width + 0.25
    content_w = canvas.content_width - (bar_width + 0.25)

    # Client Tag / Category
    curr_y = 1.30
    if client_tag:
        tag_box = canvas.slide.shapes.add_textbox(
            Inches(content_x),
            Inches(curr_y),
            Inches(content_w),
            Inches(0.30),
        )
        tf_tag = tag_box.text_frame
        set_text_frame_margins(tf_tag, top=0, bottom=0, left=0, right=0)
        p_tag = tf_tag.paragraphs[0]
        r_tag = p_tag.add_run()
        format_run(
            r_tag,
            client_tag.upper(),
            theme.fonts.font_title,
            theme.fonts.size_action_title,
            theme.colors.primary,
            bold=True,
        )
        curr_y += 0.40

    # Main Big Title
    title_box = canvas.slide.shapes.add_textbox(
        Inches(content_x),
        Inches(curr_y),
        Inches(content_w),
        Inches(1.10),
    )
    tf_title = title_box.text_frame
    set_text_frame_margins(tf_title, top=0, bottom=0, left=0, right=0)
    p_title = tf_title.paragraphs[0]
    r_title = p_title.add_run()
    format_run(
        r_title,
        title,
        theme.fonts.font_title,
        28.0,
        text_title_col,
        bold=True,
    )
    curr_y += 1.15

    # Subtitle
    if subtitle:
        sub_box = canvas.slide.shapes.add_textbox(
            Inches(content_x),
            Inches(curr_y),
            Inches(content_w),
            Inches(0.40),
        )
        tf_sub = sub_box.text_frame
        set_text_frame_margins(tf_sub, top=0, bottom=0, left=0, right=0)
        p_sub = tf_sub.paragraphs[0]
        r_sub = p_sub.add_run()
        format_run(
            r_sub,
            subtitle,
            theme.fonts.font_body,
            theme.fonts.size_action_title,
            text_sub_col,
            bold=False,
        )
        curr_y += 0.50

    # Metadata at bottom: Author, Date, Version
    meta_y = canvas.slide_height - canvas.margin_bottom - 0.45
    meta_box = canvas.slide.shapes.add_textbox(
        Inches(content_x),
        Inches(meta_y),
        Inches(content_w),
        Inches(0.35),
    )
    tf_meta = meta_box.text_frame
    set_text_frame_margins(tf_meta, top=0, bottom=0, left=0, right=0)
    p_meta = tf_meta.paragraphs[0]

    meta_parts = []
    if author:
        meta_parts.append(author)
    if date_str:
        meta_parts.append(date_str)
    meta_text = "  |  ".join(meta_parts)

    r_meta = p_meta.add_run()
    format_run(
        r_meta,
        meta_text,
        theme.fonts.font_body,
        theme.fonts.size_body,
        text_muted_col,
        bold=False,
    )
