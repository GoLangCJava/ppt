#!/usr/bin/env python3
"""
Executive Presentation Generator: Databricks vs. Apache Doris for Enterprise Data Archiving.
Adhering to McKinsey Structured Communication, Consulting Visual Standards, and Automated Agent SOP.
Strict Zero-Overlap Geometry & 16:9 Widescreen Layout.
"""

import sys
sys.path.insert(0, '/home/user/ppt_deps')
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN


# ==============================================================================
# 1. COLOR TOKENS & VISUAL PALETTE (Consulting 60-30-10 Enterprise Standards)
# ==============================================================================
class C:
    # 60% Surfaces & Neutrals
    BG_CANVAS        = RGBColor(248, 250, 252)  # #F8FAFC - Main slide background
    BG_WHITE         = RGBColor(255, 255, 255)  # #FFFFFF - Card white surface
    BG_SLATE_LIGHT   = RGBColor(241, 245, 249)  # #F1F5F9 - Metric card surface
    SOFT_BORDER      = RGBColor(203, 213, 225)  # #CBD5E1 - 0.75pt border
    LINE_SUBTLE      = RGBColor(226, 232, 240)  # #E2E8F0 - Divider line

    # 30% Structural & Brand Colors
    TEXT_MAIN        = RGBColor(15, 23, 42)     # #0F172A - Pitch Black Slate
    TEXT_BODY        = RGBColor(51, 65, 85)     # #334155 - Slate Body
    TEXT_MUTED       = RGBColor(100, 116, 139)  # #64748B - Gray Captions
    AZURE_BLUE       = RGBColor(0, 120, 212)    # #0078D4 - Lead Note & Primary Accent
    ICE_BLUE         = RGBColor(235, 245, 255)  # #EBF5FF - Databricks Highlight Tint

    # 10% Visual Accents & Status
    DATABRICKS_RED   = RGBColor(234, 56, 36)    # #EA3824 - Databricks Flame Red
    DORIS_CYAN       = RGBColor(2, 132, 199)    # #0284C7 - Apache Doris Ocean Blue
    DORIS_LIGHT      = RGBColor(240, 249, 255)  # #F0F9FF - Doris Tint
    SUCCESS          = RGBColor(16, 185, 129)   # #10B981 - Green
    WARNING          = RGBColor(245, 158, 11)   # #F59E0B - Amber
    DANGER           = RGBColor(239, 68, 68)    # #EF4444 - Red


class F:
    SERIF = "Georgia"
    SANS  = "Microsoft YaHei"
    MONO  = "Consolas"


# ==============================================================================
# 2. GEOMETRY CONFIGURATION (10.0 x 5.625 INCHES)
# ==============================================================================
class G:
    WIDTH = 10.0
    HEIGHT = 5.625
    MARGIN_LEFT = 0.50
    MARGIN_RIGHT = 0.50
    CONTENT_W = WIDTH - MARGIN_LEFT - MARGIN_RIGHT  # 9.00 inches

    # Vertical Clearance Zones
    TRACKER_Y = 0.26
    TRACKER_H = 0.18    # ends at 0.44

    TITLE_Y   = 0.46
    TITLE_H   = 0.36    # ends at 0.82

    LEAD_Y    = 0.84
    LEAD_H    = 0.28    # ends at 1.12

    DIVIDER_Y = 1.16

    BODY_Y    = 1.26
    BODY_H    = 3.52    # ends at 4.78

    FOOTER_Y  = 4.88
    FOOTER_H  = 0.42    # ends at 5.30


def draw_rect(slide, left, top, width, height, fill_rgb=None, line_rgb=None, line_pt=1.0, shape_type=MSO_SHAPE.RECTANGLE):
    shape = slide.shapes.add_shape(
        shape_type, Inches(left), Inches(top), Inches(width), Inches(height)
    )
    if fill_rgb:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_rgb
    else:
        shape.fill.background()

    if line_rgb:
        shape.line.color.rgb = line_rgb
        shape.line.width = Pt(line_pt)
    else:
        shape.line.fill.background()
    return shape


def build_databricks_vs_doris_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 1. Subtle Canvas Background
    draw_rect(slide, 0, 0, G.WIDTH, G.HEIGHT, fill_rgb=C.BG_CANVAS)

    # 2. Top Header - Tracker
    tb_tr = slide.shapes.add_textbox(Inches(G.MARGIN_LEFT), Inches(G.TRACKER_Y), Inches(G.CONTENT_W), Inches(G.TRACKER_H))
    tf_tr = tb_tr.text_frame
    tf_tr.word_wrap = True
    tf_tr.margin_top = tf_tr.margin_bottom = tf_tr.margin_left = tf_tr.margin_right = Inches(0)
    p_tr = tf_tr.paragraphs[0]
    r_tr = p_tr.add_run()
    r_tr.text = "ENTERPRISE DATA ARCHITECTURE | TECHNOLOGY EVALUATION & ARCHIVING STRATEGY"
    r_tr.font.name = F.SERIF
    r_tr.font.size = Pt(7.5)
    r_tr.font.color.rgb = C.TEXT_MUTED
    r_tr.font.bold = True

    # 3. Top Header - Action Title (McKinsey BLUF: Complete governing thought)
    tb_ti = slide.shapes.add_textbox(Inches(G.MARGIN_LEFT), Inches(G.TITLE_Y), Inches(G.CONTENT_W), Inches(G.TITLE_H))
    tf_ti = tb_ti.text_frame
    tf_ti.word_wrap = True
    tf_ti.margin_top = tf_ti.margin_bottom = tf_ti.margin_left = tf_ti.margin_right = Inches(0)
    p_ti = tf_ti.paragraphs[0]
    r_ti = p_ti.add_run()
    r_ti.text = "企业数据归档选型：海量合规首选 Databricks 湖仓，秒级高频检索协同 Doris 加速"
    r_ti.font.name = F.SERIF
    r_ti.font.size = Pt(14.5)
    r_ti.font.color.rgb = C.TEXT_MAIN
    r_ti.font.bold = True

    # 4. Top Header - Lead Note
    tb_le = slide.shapes.add_textbox(Inches(G.MARGIN_LEFT), Inches(G.LEAD_Y), Inches(G.CONTENT_W), Inches(G.LEAD_H))
    tf_le = tb_le.text_frame
    tf_le.word_wrap = True
    tf_le.margin_top = tf_le.margin_bottom = tf_le.margin_left = tf_le.margin_right = Inches(0)
    p_le = tf_le.paragraphs[0]
    r_le = p_le.add_run()
    r_le.text = "基于冷热分层法则：Databricks 承载 90%+ 海量数据低成本安全封存；Doris 承接 10% 业务高频并发点查与交互式钻取。"
    r_le.font.name = F.SANS
    r_le.font.size = Pt(9.5)
    r_le.font.color.rgb = C.AZURE_BLUE

    # 5. Header Divider Line
    div = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(G.MARGIN_LEFT), Inches(G.DIVIDER_Y), Inches(G.CONTENT_W), Pt(0.75))
    div.fill.solid()
    div.fill.fore_color.rgb = C.LINE_SUBTLE
    div.line.fill.background()

    # --------------------------------------------------------------------------
    # 6. BODY CANVAS (Height: 3.52 in) - Split into 3 Columns
    # Col 1: Databricks (3.40 in)
    # Col 2: Doris (3.40 in)
    # Col 3: Scorecard & Metrics (1.90 in)
    # --------------------------------------------------------------------------
    col1_w = 3.40
    col2_w = 3.40
    col3_w = 1.90
    gap = 0.15

    col1_x = G.MARGIN_LEFT
    col2_x = col1_x + col1_w + gap
    col3_x = col2_x + col2_w + gap

    # ========================== COLUMN 1: DATABRICKS ==========================
    card_db = draw_rect(
        slide, col1_x, G.BODY_Y, col1_w, G.BODY_H,
        fill_rgb=C.ICE_BLUE, line_rgb=C.AZURE_BLUE, line_pt=1.5,
        shape_type=MSO_SHAPE.ROUNDED_RECTANGLE
    )
    # Top Accent Strip (Databricks Flame Red)
    draw_rect(slide, col1_x + 0.10, G.BODY_Y + 0.08, col1_w - 0.20, 0.05, fill_rgb=C.DATABRICKS_RED)

    tf_db = card_db.text_frame
    tf_db.word_wrap = True
    tf_db.margin_top = Inches(0.18)
    tf_db.margin_bottom = Inches(0.08)
    tf_db.margin_left = tf_db.margin_right = Inches(0.14)

    # Pill Badge
    p_db_pill = tf_db.paragraphs[0]
    r_db_pill = p_db_pill.add_run()
    r_db_pill.text = "【👑 主归档核心推荐 · 90%+ 海量数据】\n"
    r_db_pill.font.name = F.SANS
    r_db_pill.font.size = Pt(8.5)
    r_db_pill.font.color.rgb = C.DATABRICKS_RED
    r_db_pill.font.bold = True

    # Title
    p_db_title = tf_db.add_paragraph()
    r_db_title = p_db_title.add_run()
    r_db_title.text = "Azure Databricks Lakehouse\n"
    r_db_title.font.name = F.SERIF
    r_db_title.font.size = Pt(12.0)
    r_db_title.font.bold = True
    r_db_title.font.color.rgb = C.TEXT_MAIN

    r_db_sub = p_db_title.add_run()
    r_db_sub.text = "存算解耦 · 开放标准 · 法律合规底座"
    r_db_sub.font.name = F.SANS
    r_db_sub.font.size = Pt(7.5)
    r_db_sub.font.color.rgb = C.TEXT_MUTED

    # Strengths Bullets
    bullets_db = [
        ("📦 极低持有成本 (TCO)：", "基于 Delta/Parquet 开放格式，直接写入对象存储 Archive 冷层（~$0.001/GB/月），无查询时计算资源弹性缩容归零。"),
        ("🛡️ 审计与法律合规 (WORM)：", "原生满足监管不可篡改要求；时间旅行 (Time Travel) 精确重现历史审计时点状态；支持 GDPR 遗忘权精准清除。"),
        ("🌐 跨系统统一治理：", "Unity Catalog 提供贯穿全企业的表/行/列级权限管控、动态脱敏与端到端数据血缘，彻底杜绝数据孤岛与泄露。"),
        ("🔄 架构解耦无绑定：", "底层数据 100% 开放格式存储，避免被专有数据库文件格式绑架，随时可供 Spark, Presto 或外部引擎读取。"),
    ]

    for prefix, body in bullets_db:
        pb = tf_db.add_paragraph()
        pb.space_before = Pt(3.0)
        rb1 = pb.add_run()
        rb1.text = f"• {prefix}"
        rb1.font.name = F.SANS
        rb1.font.size = Pt(7.0)
        rb1.font.bold = True
        rb1.font.color.rgb = C.TEXT_MAIN

        rb2 = pb.add_run()
        rb2.text = body
        rb2.font.name = F.SANS
        rb2.font.size = Pt(6.8)
        rb2.font.color.rgb = C.TEXT_BODY

    # Best-for & Boundary
    pg_db = tf_db.add_paragraph()
    pg_db.space_before = Pt(6.0)
    r_g1 = pg_db.add_run()
    r_g1.text = "✓ 最佳场景: "
    r_g1.font.name = F.SANS
    r_g1.font.size = Pt(7.0)
    r_g1.font.color.rgb = C.SUCCESS
    r_g1.font.bold = True
    r_g2 = pg_db.add_run()
    r_g2.text = "退役系统全量历史归档、5~10年合规审计抽查、离线年报多维穿透分析。\n"
    r_g2.font.name = F.SANS
    r_g2.font.size = Pt(6.8)
    r_g2.font.color.rgb = C.TEXT_BODY

    r_b1 = pg_db.add_run()
    r_b1.text = "✕ 边界限制: "
    r_b1.font.name = F.SANS
    r_b1.font.size = Pt(7.0)
    r_b1.font.color.rgb = C.DANGER
    r_b1.font.bold = True
    r_b2 = pg_db.add_run()
    r_b2.text = "Serverless 冷启动时延在数秒级，不适合千人级前台高并发毫秒点查。"
    r_b2.font.name = F.SANS
    r_b2.font.size = Pt(6.8)
    r_b2.font.color.rgb = C.TEXT_BODY

    # ========================== COLUMN 2: APACHE DORIS ==========================
    card_doris = draw_rect(
        slide, col2_x, G.BODY_Y, col2_w, G.BODY_H,
        fill_rgb=C.BG_WHITE, line_rgb=C.SOFT_BORDER, line_pt=1.0,
        shape_type=MSO_SHAPE.ROUNDED_RECTANGLE
    )
    # Top Accent Strip (Doris Ocean Blue)
    draw_rect(slide, col2_x + 0.10, G.BODY_Y + 0.08, col2_w - 0.20, 0.05, fill_rgb=C.DORIS_CYAN)

    tf_dr = card_doris.text_frame
    tf_dr.word_wrap = True
    tf_dr.margin_top = Inches(0.18)
    tf_dr.margin_bottom = Inches(0.08)
    tf_dr.margin_left = tf_dr.margin_right = Inches(0.14)

    # Pill Badge
    p_dr_pill = tf_dr.paragraphs[0]
    r_dr_pill = p_dr_pill.add_run()
    r_dr_pill.text = "【⚡ 高频交互检索引擎 · 10% 温热数据】\n"
    r_dr_pill.font.name = F.SANS
    r_dr_pill.font.size = Pt(8.5)
    r_dr_pill.font.color.rgb = C.DORIS_CYAN
    r_dr_pill.font.bold = True

    # Title
    p_dr_title = tf_dr.add_paragraph()
    r_dr_title = p_dr_title.add_run()
    r_dr_title.text = "Apache Doris MPP Engine\n"
    r_dr_title.font.name = F.SERIF
    r_dr_title.font.size = Pt(12.0)
    r_dr_title.font.bold = True
    r_dr_title.font.color.rgb = C.TEXT_MAIN

    r_dr_sub = p_dr_title.add_run()
    r_dr_sub.text = "极速响应 · 倒排索引 · 实时交互引擎"
    r_dr_sub.font.name = F.SANS
    r_dr_sub.font.size = Pt(7.5)
    r_dr_sub.font.color.rgb = C.TEXT_MUTED

    # Strengths Bullets
    bullets_dr = [
        ("⚡ 毫秒级极速响应：", "MPP 分布式计算 + 全向量化执行引擎，数亿行单表点查与多维聚合耗时 < 500ms，轻松承载数千并发查询。"),
        ("🔍 文本与倒排索引：", "内置 Inverted Index，支持海量文本日志、合同流水号与物料编码跨字段快速模糊匹配检索。"),
        ("🏢 业务前台实时反哺：", "高度兼容 MySQL 协议，可直接对接前台 CRM/ERP 业务系统或开放 API，实现老数据准实时在线调阅。"),
        ("📈 冷热分层存算分离：", "支持将历史冷分区异步转储至 S3/对象存储，兼顾部分存储成本与本地 SSD 缓存加速。"),
    ]

    for prefix, body in bullets_dr:
        pb = tf_dr.add_paragraph()
        pb.space_before = Pt(3.0)
        rb1 = pb.add_run()
        rb1.text = f"• {prefix}"
        rb1.font.name = F.SANS
        rb1.font.size = Pt(7.0)
        rb1.font.bold = True
        rb1.font.color.rgb = C.TEXT_MAIN

        rb2 = pb.add_run()
        rb2.text = body
        rb2.font.name = F.SANS
        rb2.font.size = Pt(6.8)
        rb2.font.color.rgb = C.TEXT_BODY

    # Best-for & Boundary
    pg_dr = tf_dr.add_paragraph()
    pg_dr.space_before = Pt(6.0)
    r_dg1 = pg_dr.add_run()
    r_dg1.text = "✓ 最佳场景: "
    r_dg1.font.name = F.SANS
    r_dg1.font.size = Pt(7.0)
    r_dg1.font.color.rgb = C.SUCCESS
    r_dg1.font.bold = True
    r_dg2 = pg_dr.add_run()
    r_dg2.text = "近 1~3 年高频订单检索、前台客户历史交易秒查、交互式大屏即席分析。\n"
    r_dg2.font.name = F.SANS
    r_dg2.font.size = Pt(6.8)
    r_dg2.font.color.rgb = C.TEXT_BODY

    r_db1 = pg_dr.add_run()
    r_db1.text = "✕ 边界限制: "
    r_db1.font.name = F.SANS
    r_db1.font.size = Pt(7.0)
    r_db1.font.color.rgb = C.DANGER
    r_db1.font.bold = True
    r_db2 = pg_dr.add_run()
    r_db2.text = "需常驻 BE/FE 节点保障服务，PB 级长期纯冷封存硬件成本远高于对象存储。"
    r_db2.font.name = F.SANS
    r_db2.font.size = Pt(6.8)
    r_db2.font.color.rgb = C.TEXT_BODY

    # ========================== COLUMN 3: SCORECARD & METRICS ==========================
    card_sc = draw_rect(
        slide, col3_x, G.BODY_Y, col3_w, G.BODY_H,
        fill_rgb=C.BG_SLATE_LIGHT, line_rgb=C.SOFT_BORDER, line_pt=1.0,
        shape_type=MSO_SHAPE.ROUNDED_RECTANGLE
    )

    tf_sc = card_sc.text_frame
    tf_sc.word_wrap = True
    tf_sc.margin_top = Inches(0.12)
    tf_sc.margin_bottom = Inches(0.08)
    tf_sc.margin_left = tf_sc.margin_right = Inches(0.10)

    p_sc_title = tf_sc.paragraphs[0]
    r_sc_title = p_sc_title.add_run()
    r_sc_title.text = "选型评估矩阵 (Scorecard)\n"
    r_sc_title.font.name = F.SERIF
    r_sc_title.font.size = Pt(9.5)
    r_sc_title.font.bold = True
    r_sc_title.font.color.rgb = C.AZURE_BLUE

    # 4 Dimensions with Harvey Balls
    eval_items = [
        ("长期存储 TCO", "Databricks: ● 极低\nDoris: ▲ 中等 (节点开销)"),
        ("闲时算力成本", "Databricks: ● 弹性归零\nDoris: ▲ 常驻集群"),
        ("毫秒级并发点查", "Databricks: ▲ 秒级响应\nDoris: ● <500ms 极速"),
        ("全域血缘与治理", "Databricks: ● 原生 UC\nDoris: ○ 需外挂系统"),
    ]

    for dim, note in eval_items:
        p_item = tf_sc.add_paragraph()
        p_item.space_before = Pt(3.0)
        
        r_dim = p_item.add_run()
        r_dim.text = f"【{dim}】\n"
        r_dim.font.name = F.SANS
        r_dim.font.size = Pt(7.0)
        r_dim.font.bold = True
        r_dim.font.color.rgb = C.TEXT_MAIN

        r_note = p_item.add_run()
        r_note.text = note
        r_note.font.name = F.SANS
        r_note.font.size = Pt(6.5)
        r_note.font.color.rgb = C.TEXT_BODY

    # Bottom Quantitative Impact Badge inside Column 3
    p_kpi = tf_sc.add_paragraph()
    p_kpi.space_before = Pt(6.0)

    r_kpi_num = p_kpi.add_run()
    r_kpi_num.text = "📉 -75%\n"
    r_kpi_num.font.name = F.SERIF
    r_kpi_num.font.size = Pt(16.0)
    r_kpi_num.font.bold = True
    r_kpi_num.font.color.rgb = C.AZURE_BLUE

    r_kpi_desc = p_kpi.add_run()
    r_kpi_desc.text = "10年期归档综合 TCO 节省\n"
    r_kpi_desc.font.name = F.SANS
    r_kpi_desc.font.size = Pt(7.0)
    r_kpi_desc.font.bold = True
    r_kpi_desc.font.color.rgb = C.TEXT_MAIN

    r_kpi_sub = p_kpi.add_run()
    r_kpi_sub.text = "存算分离 + 开放 Delta 格式，杜绝专有数仓常驻节点浪费。"
    r_kpi_sub.font.name = F.SANS
    r_kpi_sub.font.size = Pt(6.2)
    r_kpi_sub.font.color.rgb = C.TEXT_MUTED

    # --------------------------------------------------------------------------
    # 7. FOOTER TAKEAWAY BANNER (Height: 0.42 in)
    # --------------------------------------------------------------------------
    card_ft = draw_rect(
        slide, G.MARGIN_LEFT, G.FOOTER_Y, G.CONTENT_W, G.FOOTER_H,
        fill_rgb=C.ICE_BLUE, line_rgb=C.AZURE_BLUE, line_pt=1.0,
        shape_type=MSO_SHAPE.ROUNDED_RECTANGLE
    )
    tf_ft = card_ft.text_frame
    tf_ft.word_wrap = True
    tf_ft.margin_top = Inches(0.06)
    tf_ft.margin_bottom = Inches(0.06)
    tf_ft.margin_left = Inches(0.14)
    tf_ft.margin_right = Inches(0.14)

    p_ft = tf_ft.paragraphs[0]
    r_ft1 = p_ft.add_run()
    r_ft1.text = "【麦肯锡咨询建议：双擎冷热分层协同】 "
    r_ft1.font.name = F.SANS
    r_ft1.font.size = Pt(8.5)
    r_ft1.font.bold = True
    r_ft1.font.color.rgb = C.AZURE_BLUE

    r_ft2 = p_ft.add_run()
    r_ft2.text = (
        "以 Databricks Lakehouse 作为企业 90%+ 历史数据的单一事实源（SSOT）低成本安全封存，满足 5~10 年法律审计合规；"
        "针对近 1~3 年高频业务调阅，按需加载温热索引与主题明细至 Doris，兼顾极低 TCO 与亚秒级极速交互。"
    )
    r_ft2.font.name = F.SANS
    r_ft2.font.size = Pt(7.5)
    r_ft2.font.color.rgb = C.TEXT_MAIN


def main():
    prs = Presentation()
    prs.slide_width = Inches(G.WIDTH)
    prs.slide_height = Inches(G.HEIGHT)

    print("Building One-Page PPT: Databricks vs. Apache Doris for Enterprise Data Archiving...")
    build_databricks_vs_doris_slide(prs)

    out_file = Path("Databricks_vs_Doris_Archiving_Selection.pptx")
    prs.save(str(out_file))
    print(f"✅ Generated single-page presentation to {out_file} successfully!")


if __name__ == "__main__":
    main()
