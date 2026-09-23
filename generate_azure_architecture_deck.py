#!/usr/bin/env python3
"""
Custom Executive Presentation Generator: Azure Databricks & ADF & Synapse Architecture.
High-Creativity, Modern Consulting Layout with Zero Overlaps & Precise Geometry.
"""

from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN


# ==============================================================================
# 1. COLOR TOKENS & VISUAL PALETTE (Microsoft Azure & Databricks Enterprise Aesthetic)
# ==============================================================================
class C:
    # Primary Azure Brand Blues
    AZURE_BLUE       = RGBColor(0, 120, 212)    # #0078D4 - Classic Azure Blue
    NAVY_DARK        = RGBColor(10, 25, 47)     # #0A192F - Deep Slate Navy Background
    NAVY_SURFACE     = RGBColor(15, 33, 64)     # #0F2140 - Dark Card Surface
    ICE_BLUE         = RGBColor(235, 245, 255)  # #EBF5FF - High-light Tint Fill
    SOFT_BORDER      = RGBColor(203, 213, 225)  # #CBD5E1 - Card Border Line
    LINE_SUBTLE      = RGBColor(226, 232, 240)  # #E2E8F0 - Divider Line

    # Product Brand Accents
    ADF_PURPLE       = RGBColor(92, 45, 145)    # #5C2D91 - Data Factory Identity
    ADF_LIGHT        = RGBColor(245, 240, 255)  # #F5F0FF - ADF Tint
    DATABRICKS_RED   = RGBColor(234, 56, 36)    # #EA3824 - Databricks Flame
    DATABRICKS_LIGHT = RGBColor(254, 242, 242)  # #FEF2F2 - Databricks Tint
    SYNAPSE_TEAL     = RGBColor(0, 130, 114)    # #008272 - Synapse Cyan/Teal
    SYNAPSE_LIGHT    = RGBColor(240, 253, 250)  # #F0FDFA - Synapse Tint

    # Medallion Lakehouse Tones
    BRONZE           = RGBColor(180, 83, 9)     # #B45309 - Bronze Layer
    BRONZE_LIGHT     = RGBColor(254, 243, 199)  # #FEF3C7 - Bronze Tint
    SILVER           = RGBColor(71, 85, 105)    # #475569 - Silver Layer
    SILVER_LIGHT     = RGBColor(241, 245, 249)  # #F1F5F9 - Silver Tint
    GOLD             = RGBColor(217, 119, 6)    # #D97706 - Gold Layer
    GOLD_LIGHT       = RGBColor(254, 249, 195)  # #FEF9C3 - Gold Tint

    # Functional State Colors
    SUCCESS          = RGBColor(16, 185, 129)   # #10B981 - Green
    SUCCESS_LIGHT    = RGBColor(236, 253, 245)  # #ECFDF5
    WARNING          = RGBColor(245, 158, 11)   # #F59E0B - Amber
    DANGER           = RGBColor(239, 68, 68)    # #EF4444 - Red

    # Neutrals
    BG_WHITE         = RGBColor(255, 255, 255)
    BG_SLATE_LIGHT   = RGBColor(248, 250, 252)  # #F8FAFC - Main Slide Background
    TEXT_MAIN        = RGBColor(15, 23, 42)     # #0F172A - Pitch Black Slate
    TEXT_MUTED       = RGBColor(100, 116, 139)  # #64748B - Gray Captions
    TEXT_BODY        = RGBColor(51, 65, 85)     # #334155 - Body Gray
    TEXT_WHITE       = RGBColor(255, 255, 255)


class F:
    TITLE = "Georgia"
    BODY  = "Microsoft YaHei"
    MONO  = "Consolas"


# ==============================================================================
# 2. GEOMETRY & GLOBAL COORDINATE RULES (10.0 x 5.625 INCHES)
# ==============================================================================
class G:
    WIDTH = 10.0
    HEIGHT = 5.625
    MARGIN_LEFT = 0.50
    MARGIN_RIGHT = 0.50
    CONTENT_W = WIDTH - MARGIN_LEFT - MARGIN_RIGHT  # 9.00 inches

    # Header Zones (Guaranteed disjoint Y-ranges)
    TRACKER_Y = 0.26
    TRACKER_H = 0.18    # ends at 0.44

    TITLE_Y   = 0.46
    TITLE_H   = 0.36    # ends at 0.82

    LEAD_Y    = 0.84
    LEAD_H    = 0.28    # ends at 1.12

    DIVIDER_Y = 1.16

    # Body Zone
    BODY_Y    = 1.26
    BODY_H    = 3.52    # ends at 4.78

    # Footer Takeaway Zone
    FOOTER_Y  = 4.88
    FOOTER_H  = 0.42    # ends at 5.30


# ==============================================================================
# 3. HELPER FUNCTIONS
# ==============================================================================
def draw_rect(slide, left, top, width, height, fill_rgb=None, line_rgb=None, line_pt=1.0, shape_type=MSO_SHAPE.RECTANGLE):
    """Draw a vector rectangle or rounded rectangle."""
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


def add_header(slide, tracker: str, action_title: str, lead_note: str):
    """Render top 3-tier header with strictly zero overlapping boxes."""
    # 1. Tracker
    tb_tr = slide.shapes.add_textbox(Inches(G.MARGIN_LEFT), Inches(G.TRACKER_Y), Inches(G.CONTENT_W), Inches(G.TRACKER_H))
    tf_tr = tb_tr.text_frame
    tf_tr.word_wrap = True
    tf_tr.margin_top = tf_tr.margin_bottom = tf_tr.margin_left = tf_tr.margin_right = Inches(0)
    p_tr = tf_tr.paragraphs[0]
    r_tr = p_tr.add_run()
    r_tr.text = tracker.upper()
    r_tr.font.name = F.TITLE
    r_tr.font.size = Pt(7.5)
    r_tr.font.color.rgb = C.TEXT_MUTED

    # 2. Action Title
    tb_ti = slide.shapes.add_textbox(Inches(G.MARGIN_LEFT), Inches(G.TITLE_Y), Inches(G.CONTENT_W), Inches(G.TITLE_H))
    tf_ti = tb_ti.text_frame
    tf_ti.word_wrap = True
    tf_ti.margin_top = tf_ti.margin_bottom = tf_ti.margin_left = tf_ti.margin_right = Inches(0)
    p_ti = tf_ti.paragraphs[0]
    r_ti = p_ti.add_run()
    r_ti.text = action_title
    r_ti.font.name = F.TITLE
    r_ti.font.size = Pt(14.5)
    r_ti.font.color.rgb = C.TEXT_MAIN
    r_ti.font.bold = True

    # 3. Lead Note
    tb_le = slide.shapes.add_textbox(Inches(G.MARGIN_LEFT), Inches(G.LEAD_Y), Inches(G.CONTENT_W), Inches(G.LEAD_H))
    tf_le = tb_le.text_frame
    tf_le.word_wrap = True
    tf_le.margin_top = tf_le.margin_bottom = tf_le.margin_left = tf_le.margin_right = Inches(0)
    p_le = tf_le.paragraphs[0]
    r_le = p_le.add_run()
    r_le.text = lead_note
    r_le.font.name = F.BODY
    r_le.font.size = Pt(10.0)
    r_le.font.color.rgb = C.AZURE_BLUE
    r_le.font.bold = False

    # 4. Divider Line
    div = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(G.MARGIN_LEFT), Inches(G.DIVIDER_Y), Inches(G.CONTENT_W), Pt(0.75))
    div.fill.solid()
    div.fill.fore_color.rgb = C.LINE_SUBTLE
    div.line.fill.background()


def add_takeaway(slide, title: str, text: str):
    """Render footer takeaway banner."""
    card = draw_rect(slide, G.MARGIN_LEFT, G.FOOTER_Y, G.CONTENT_W, G.FOOTER_H, fill_rgb=C.ICE_BLUE, line_rgb=C.AZURE_BLUE, line_pt=1.0, shape_type=MSO_SHAPE.ROUNDED_RECTANGLE)
    tf = card.text_frame
    tf.word_wrap = True
    tf.margin_top = Inches(0.06)
    tf.margin_bottom = Inches(0.06)
    tf.margin_left = Inches(0.15)
    tf.margin_right = Inches(0.15)
    p = tf.paragraphs[0]

    r1 = p.add_run()
    r1.text = f"【{title}】 "
    r1.font.name = F.BODY
    r1.font.size = Pt(8.5)
    r1.font.color.rgb = C.AZURE_BLUE
    r1.font.bold = True

    r2 = p.add_run()
    r2.text = text
    r2.font.name = F.BODY
    r2.font.size = Pt(8.0)
    r2.font.color.rgb = C.TEXT_MAIN


# ==============================================================================
# 4. SLIDE 1: CREATIVE HERO COVER SLIDE
# ==============================================================================
def make_slide_1_cover(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Deep Slate-Navy Background
    bg = draw_rect(slide, 0, 0, G.WIDTH, G.HEIGHT, fill_rgb=C.NAVY_DARK)

    # Decorative Cyan/Blue Top Accent Line
    top_line = draw_rect(slide, 0, 0, G.WIDTH, 0.08, fill_rgb=C.AZURE_BLUE)

    # Left Branding Vertical Accent Pill
    accent_bar = draw_rect(slide, G.MARGIN_LEFT, 1.20, 0.16, 2.90, fill_rgb=C.AZURE_BLUE)

    # Left Content Box (60% width)
    left_x = G.MARGIN_LEFT + 0.35
    left_w = 5.20
    tb = slide.shapes.add_textbox(Inches(left_x), Inches(1.22), Inches(left_w), Inches(3.00))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_top = tf.margin_bottom = tf.margin_left = tf.margin_right = Inches(0)

    # 1. Kicker
    p0 = tf.paragraphs[0]
    r0 = p0.add_run()
    r0.text = "AZURE DATA ARCHITECTURE STRATEGY\n"
    r0.font.name = F.TITLE
    r0.font.size = Pt(10.5)
    r0.font.color.rgb = C.AZURE_BLUE
    r0.font.bold = True

    # 2. Main Title
    p1 = tf.add_paragraph()
    p1.space_before = Pt(6.0)
    r1 = p1.add_run()
    r1.text = "Azure Modern Data\nPlatform Architecture"
    r1.font.name = F.TITLE
    r1.font.size = Pt(26.0)
    r1.font.color.rgb = C.TEXT_WHITE
    r1.font.bold = True

    # 3. Subtitle
    p2 = tf.add_paragraph()
    p2.space_before = Pt(10.0)
    r2 = p2.add_run()
    r2.text = "Azure Data Factory (ADF) · Azure Databricks · Azure Synapse\n三驾马车协同技术架构与湖仓一体化方案建议书"
    r2.font.name = F.BODY
    r2.font.size = Pt(11.5)
    r2.font.color.rgb = C.ICE_BLUE

    # 4. Meta
    p3 = tf.add_paragraph()
    p3.space_before = Pt(16.0)
    r3 = p3.add_run()
    r3.text = "Cloud Architecture & Data Practice  |  September 2026  |  Confidential"
    r3.font.name = F.BODY
    r3.font.size = Pt(8.5)
    r3.font.color.rgb = C.TEXT_MUTED

    # Right Side: Creative "Triad" Architecture Visual Blueprint Preview Card (35% width)
    right_x = G.MARGIN_LEFT + left_w + 0.50
    right_w = G.CONTENT_W - left_w - 0.50  # ~3.30 in
    card_right = draw_rect(slide, right_x, 1.15, right_w, 3.40, fill_rgb=C.NAVY_SURFACE, line_rgb=C.AZURE_BLUE, line_pt=1.5, shape_type=MSO_SHAPE.ROUNDED_RECTANGLE)

    tf_r = card_right.text_frame
    tf_r.word_wrap = True
    tf_r.margin_top = Inches(0.12)
    tf_r.margin_left = tf_r.margin_right = Inches(0.14)

    pr0 = tf_r.paragraphs[0]
    rr0 = pr0.add_run()
    rr0.text = "THE ARCHITECTURE TRIAD (三驾马车)\n"
    rr0.font.name = F.TITLE
    rr0.font.size = Pt(9.5)
    rr0.font.color.rgb = C.AZURE_BLUE
    rr0.font.bold = True

    pillars = [
        ("⚡ Azure Data Factory", "Ingestion & Multi-Source Orchestration", "全域混合调度纽带 · 100+ 连接器", C.ADF_PURPLE),
        ("💎 Azure Databricks", "Lakehouse Core & AI Computing", "Medallion 湖仓计算核心 · Photon 向量化", C.DATABRICKS_RED),
        ("🏛️ Azure Synapse", "Enterprise Data Warehousing & Serving", "专用 SQL 池 MPP · Power BI 极速呈现", C.SYNAPSE_TEAL),
    ]

    for p_name, p_en, p_desc, p_color in pillars:
        pr_item = tf_r.add_paragraph()
        pr_item.space_before = Pt(8.0)
        
        r_name = pr_item.add_run()
        r_name.text = f"{p_name}\n"
        r_name.font.name = F.BODY
        r_name.font.size = Pt(10.0)
        r_name.font.color.rgb = C.TEXT_WHITE
        r_name.font.bold = True

        r_en = pr_item.add_run()
        r_en.text = f"{p_en}\n"
        r_en.font.name = F.MONO
        r_en.font.size = Pt(7.0)
        r_en.font.color.rgb = C.AZURE_BLUE

        r_desc = pr_item.add_run()
        r_desc.text = p_desc
        r_desc.font.name = F.BODY
        r_desc.font.size = Pt(7.5)
        r_desc.font.color.rgb = C.SOFT_BORDER

    # Bottom Foundation note in right card
    pr_bot = tf_r.add_paragraph()
    pr_bot.space_before = Pt(10.0)
    r_bot = pr_bot.add_run()
    r_bot.text = "统一治理中枢：Unity Catalog + Microsoft Purview"
    r_bot.font.name = F.BODY
    r_bot.font.size = Pt(7.0)
    r_bot.font.color.rgb = C.TEXT_MUTED


# ==============================================================================
# 5. SLIDE 2: THE TRIAD ROLE MATRIX (三驾马车：定位与能力深度对比)
# ==============================================================================
def make_slide_2_roles(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(
        slide,
        tracker="Azure Data Platform Strategy | Core Roles & Positioning",
        action_title="三驾马车：ADF、Databricks 与 Synapse 各司其职，打破企业选型困惑",
        lead_note="明确三者在企业数据价值链中的专业分工：ADF管管道调度、Databricks管湖仓计算与AI、Synapse管数仓与报表消费。"
    )

    col_gap = 0.20
    col_w = (G.CONTENT_W - col_gap * 2) / 3
    card_y = G.BODY_Y
    card_h = G.BODY_H - 0.48

    columns_data = [
        {
            "tag": "全域调度接入纽带",
            "name": "Azure Data Factory",
            "en": "ORCHESTRATION & INGESTION",
            "badge_color": C.ADF_PURPLE,
            "badge_light": C.ADF_LIGHT,
            "highlight": False,
            "bullets": [
                ("多源连接生态：", "内置 100+ 原生成熟连接器，打通本地数据库、云平台与 SaaS。"),
                ("Serverless 调度编排：", "基于事件、时间轮询与动态参数化管道控制，无需常驻服务器。"),
                ("混合内网穿透：", "自托管运行时 (SHIR) 打通机房专线与 Azure VNet 私有网络。"),
            ],
            "best_for": "跨网络数据抽取、轻量文件落湖、作业触发与全流程自动化编排。",
            "avoid_for": "极复杂大数据清洗、非结构化音视频处理与多分析师交互查询。",
        },
        {
            "tag": "👑 湖仓一体计算核心",
            "name": "Azure Databricks",
            "en": "LAKEHOUSE ENGINE & AI",
            "badge_color": C.DATABRICKS_RED,
            "badge_light": C.DATABRICKS_LIGHT,
            "highlight": True,
            "bullets": [
                ("Medallion 奖章架构：", "驱动 Bronze 原始 -> Silver 清洗 -> Gold 业务宽表的逐层精炼。"),
                ("极致大数据性能：", "Photon 向量化执行引擎 + Delta Lake ACID 事务与时间旅行。"),
                ("统一数据与 AI 治理：", "Unity Catalog 提供行/列级细粒度权限控制，原生集成 MLflow。"),
            ],
            "best_for": "流批一体计算、TB/PB 级复杂清洗转换、特征工程与企业私域 AI 建模。",
            "avoid_for": "简单的点对点文件搬运调度与数千并发人员同时在线点选看板。",
        },
        {
            "tag": "企业数仓与报表消费",
            "name": "Azure Synapse Analytics",
            "en": "ENTERPRISE SERVING & EDW",
            "badge_color": C.SYNAPSE_TEAL,
            "badge_light": C.SYNAPSE_LIGHT,
            "highlight": False,
            "bullets": [
                ("专用 SQL 池 (Dedicated)：", "MPP 分布式高并发架构，承载企业级数据仓库与星型模型。"),
                ("无服务器 SQL (Serverless)：", "针对湖内文件直接执行原生 T-SQL 探索，按查询数据量计费。"),
                ("无缝 Power BI 体验：", "深度集成 DirectLake 模式，物化视图与缓存保障秒级展现。"),
            ],
            "best_for": "高并发管理驾驶舱、复杂维度建模与企业财务级统一指标口径。",
            "avoid_for": "非结构化数据处理与前沿深度学习模型的大规模分布式训练。",
        },
    ]

    for i, col in enumerate(columns_data):
        cx = G.MARGIN_LEFT + i * (col_w + col_gap)
        is_hl = col["highlight"]

        # Card container with top accent band
        card = draw_rect(slide, cx, card_y, col_w, card_h, fill_rgb=C.ICE_BLUE if is_hl else C.BG_WHITE, line_rgb=C.AZURE_BLUE if is_hl else C.SOFT_BORDER, line_pt=1.5 if is_hl else 1.0, shape_type=MSO_SHAPE.ROUNDED_RECTANGLE)

        # Top Accent Color Bar on card
        top_bar = draw_rect(slide, cx + 0.08, card_y + 0.08, col_w - 0.16, 0.05, fill_rgb=col["badge_color"])

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_top = Inches(0.18)
        tf.margin_bottom = Inches(0.10)
        tf.margin_left = tf.margin_right = Inches(0.14)

        # Tag Pill
        p_tag = tf.paragraphs[0]
        r_tag = p_tag.add_run()
        r_tag.text = f"【{col['tag']}】\n"
        r_tag.font.name = F.BODY
        r_tag.font.size = Pt(8.5)
        r_tag.font.color.rgb = col["badge_color"]
        r_tag.font.bold = True

        # Product Title
        p_name = tf.add_paragraph()
        r_name = p_name.add_run()
        r_name.text = f"{col['name']}\n"
        r_name.font.name = F.TITLE
        r_name.font.size = Pt(11.5)
        r_name.font.color.rgb = C.TEXT_MAIN
        r_name.font.bold = True

        r_en = p_name.add_run()
        r_en.text = col["en"]
        r_en.font.name = F.MONO
        r_en.font.size = Pt(6.5)
        r_en.font.color.rgb = C.TEXT_MUTED

        # Bullets
        for prefix, body in col["bullets"]:
            pb = tf.add_paragraph()
            pb.space_before = Pt(3.0)
            
            rb1 = pb.add_run()
            rb1.text = f"• {prefix}"
            rb1.font.name = F.BODY
            rb1.font.size = Pt(7.5)
            rb1.font.color.rgb = C.TEXT_MAIN
            rb1.font.bold = True

            rb2 = pb.add_run()
            rb2.text = body
            rb2.font.name = F.BODY
            rb2.font.size = Pt(7.5)
            rb2.font.color.rgb = C.TEXT_BODY

        # Scenario Guide
        pg = tf.add_paragraph()
        pg.space_before = Pt(6.0)

        rg1 = pg.add_run()
        rg1.text = "✓ 最佳适用: "
        rg1.font.name = F.BODY
        rg1.font.size = Pt(7.0)
        rg1.font.color.rgb = C.SUCCESS
        rg1.font.bold = True

        rg2 = pg.add_run()
        rg2.text = f"{col['best_for']}\n"
        rg2.font.name = F.BODY
        rg2.font.size = Pt(7.0)
        rg2.font.color.rgb = C.TEXT_BODY

        rg3 = pg.add_run()
        rg3.text = "✕ 边界限制: "
        rg3.font.name = F.BODY
        rg3.font.size = Pt(7.0)
        rg3.font.color.rgb = C.DANGER
        rg3.font.bold = True

        rg4 = pg.add_run()
        rg4.text = col["avoid_for"]
        rg4.font.name = F.BODY
        rg4.font.size = Pt(7.0)
        rg4.font.color.rgb = C.TEXT_BODY

    add_takeaway(
        slide,
        "架构协同原则",
        "绝非三选一的替代关系：以 ADF 为全域数据管道纽带，以 Databricks 为中枢加工单一事实源 (SSOT)，以 Synapse 为窗口保障企业级报表毫秒级响应。"
    )


# ==============================================================================
# 6. SLIDE 3: BLUEPRINT-GRADE TECHNICAL ARCHITECTURE (真正蓝图级端到端架构图)
# ==============================================================================
def make_slide_3_architecture(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(
        slide,
        tracker="Azure Data Platform Architecture | End-to-End Blueprint",
        action_title="技术架构全景图：基于 Medallion 奖章体系的 Azure 四层现代化数据架构",
        lead_note="四层解耦：多源接入、湖仓精炼、企业数仓建模与全场景消费深度协同，贯穿端到端安全与治理基线。"
    )

    # 4 Layers + 3 Protocol Strips + 1 Foundation Bar
    layers = [
        {
            "num": "01",
            "name": "消费与洞察层",
            "en": "CONSUMPTION",
            "color": C.NAVY_DARK,
            "strip": "── ↕ DirectLake / DirectQuery 毫秒直读 · HTTPS 443 · REST APIs · OData ──",
            "modules": [
                ("📊 Power BI 决策看板", "直连 Synapse 与 Databricks SQL，秒级多维钻取分析与自助看板"),
                ("📑 企业级合规与财务报表", "固定像素报表、多维穿透钻取、审计流水核对，支持定时自动分发"),
                ("🤖 AI 与大模型应用", "基于 Azure OpenAI 私域语义问答与 MLflow 模型在线推理服务"),
                ("🔌 开放 API 与逆向数据流", "API Management 统一封装标准数据微服务，反哺业务系统"),
            ]
        },
        {
            "num": "02",
            "name": "数仓与服务层",
            "en": "SERVING & EDW",
            "color": C.SYNAPSE_TEAL,
            "strip": "── ↕ Synapse Link · T-SQL 存储过程 · Delta Sharing 跨组织安全共享 ──",
            "modules": [
                ("🏛️ Synapse 专用 SQL 池", "MPP 分布式企业数仓，承载星型模型 (Fact/Dim)，毫秒高并发支持"),
                ("⚡ Synapse Serverless", "针对数据湖开放文件直接执行原生 T-SQL，零驻留成本即席探索"),
                ("🎯 Databricks SQL 仓库", "Serverless SQL Warehouse，提供高性价比湖仓交互式查询与语义层"),
                ("📐 统一指标与度量中枢", "统一维护全局商业指标口径与语义层，彻底消除跨部门数据歧义"),
            ]
        },
        {
            "num": "03",
            "name": "湖仓计算层",
            "en": "LAKEHOUSE ENGINE",
            "color": C.DATABRICKS_RED,
            "highlight": True,
            "strip": "── ↕ ABFS (ADLS Gen2) · Delta Lake 开放格式 · Apache Spark 3.x · Auto Loader ──",
            "modules": [
                ("🥉 Bronze 原始湖 (Raw)", "保留源端 1:1 快照与增量日志，Auto Loader 毫秒入湖，版本回溯"),
                ("🥈 Silver 洁净层 (Refined)", "执行清洗去重、类型推断、Schema 演进验证与异常数据分流 (Delta)"),
                ("🥇 Gold 业务层 (Curated)", "高度精炼的主题宽表、业务指标预计算与算法特征库，ACID 保障"),
                ("🛡️ Unity Catalog 统一治理", "跨工作区统一数据目录，支持表/行/列级授权与端到端血缘追踪"),
            ]
        },
        {
            "num": "04",
            "name": "接入与基建层",
            "en": "INGESTION & INFRA",
            "color": C.ADF_PURPLE,
            "strip": None,
            "modules": [
                ("⚡ Azure Data Factory (ADF)", "全域统一调度编排流水线，支持 100+ 异构连接器与断点续传"),
                ("📡 Event Hubs 实时流接入", "高吞吐接入 IoT 设备埋点、应用日志与 Kafka 实时数据流"),
                ("🗄️ ADLS Gen2 湖存储底座", "高性价比分层存储 (Hot/Cool/Archive)，多协议兼容与静态加密"),
                ("🏢 企业多源业务系统", "SAP ERP、本地核心数据库、SaaS CRM 与各类无结构附件文件"),
            ]
        },
    ]

    curr_y = G.BODY_Y
    layer_h = 0.64
    proto_h = 0.12
    label_w = 1.30
    mod_start_x = G.MARGIN_LEFT + label_w + 0.10
    mod_total_w = G.CONTENT_W - label_w - 0.10

    for l_idx, layer in enumerate(layers):
        is_hl = layer.get("highlight", False)

        # 1. Left Label Pillar
        lbl_box = draw_rect(slide, G.MARGIN_LEFT, curr_y, label_w, layer_h, fill_rgb=C.ICE_BLUE if is_hl else C.BG_SLATE_LIGHT, line_rgb=layer["color"], line_pt=1.5 if is_hl else 0.75, shape_type=MSO_SHAPE.ROUNDED_RECTANGLE)
        tf_l = lbl_box.text_frame
        tf_l.word_wrap = True
        tf_l.margin_top = Inches(0.06)
        tf_l.margin_bottom = Inches(0.04)
        tf_l.margin_left = tf_l.margin_right = Inches(0.06)

        p_ln = tf_l.paragraphs[0]
        r_ln = p_ln.add_run()
        r_ln.text = f"{layer['num']} {layer['name']}\n"
        r_ln.font.name = F.TITLE
        r_ln.font.size = Pt(8.5)
        r_ln.font.color.rgb = layer["color"]
        r_ln.font.bold = True

        p_le = tf_l.add_paragraph()
        r_le = p_le.add_run()
        r_le.text = layer["en"]
        r_le.font.name = F.MONO
        r_le.font.size = Pt(6.5)
        r_le.font.color.rgb = C.TEXT_MUTED

        # 2. Right Module Cards (4 items per layer)
        mods = layer["modules"]
        num_m = len(mods)
        mod_gap = 0.08
        mod_w = (mod_total_w - (mod_gap * (num_m - 1))) / num_m

        for m_idx, (m_title, m_desc) in enumerate(mods):
            mx = mod_start_x + m_idx * (mod_w + mod_gap)
            m_box = draw_rect(slide, mx, curr_y, mod_w, layer_h, fill_rgb=C.BG_WHITE, line_rgb=C.AZURE_BLUE if is_hl else C.SOFT_BORDER, line_pt=1.0 if is_hl else 0.5, shape_type=MSO_SHAPE.ROUNDED_RECTANGLE)

            tf_m = m_box.text_frame
            tf_m.word_wrap = True
            tf_m.margin_top = Inches(0.04)
            tf_m.margin_bottom = Inches(0.03)
            tf_m.margin_left = tf_m.margin_right = Inches(0.06)

            p_mt = tf_m.paragraphs[0]
            r_mt = p_mt.add_run()
            r_mt.text = f"{m_title}\n"
            r_mt.font.name = F.TITLE
            r_mt.font.size = Pt(7.5)
            r_mt.font.color.rgb = C.TEXT_MAIN
            r_mt.font.bold = True

            p_md = tf_m.add_paragraph()
            r_md = p_md.add_run()
            r_md.text = m_desc
            r_md.font.name = F.BODY
            r_md.font.size = Pt(6.5)
            r_md.font.color.rgb = C.TEXT_BODY

        curr_y += layer_h

        # 3. Protocol connector between layers
        if layer["strip"]:
            tb_proto = slide.shapes.add_textbox(Inches(mod_start_x), Inches(curr_y), Inches(mod_total_w), Inches(proto_h))
            tf_p = tb_proto.text_frame
            tf_p.word_wrap = False
            tf_p.margin_top = tf_p.margin_bottom = tf_p.margin_left = tf_p.margin_right = Inches(0)
            p_pr = tf_p.paragraphs[0]
            p_pr.alignment = PP_ALIGN.CENTER
            r_pr = p_pr.add_run()
            r_pr.text = layer["strip"]
            r_pr.font.name = F.MONO
            r_pr.font.size = Pt(6.5)
            r_pr.font.color.rgb = C.TEXT_MUTED
            curr_y += proto_h

    # Cross-Cutting Governance & Security Foundation Bar
    curr_y += 0.04
    gov_bar = draw_rect(slide, G.MARGIN_LEFT, curr_y, G.CONTENT_W, 0.24, fill_rgb=C.SILVER_LIGHT, line_rgb=C.SOFT_BORDER, line_pt=0.75, shape_type=MSO_SHAPE.ROUNDED_RECTANGLE)
    tf_g = gov_bar.text_frame
    tf_g.word_wrap = True
    tf_g.margin_top = Inches(0.02)
    tf_g.margin_bottom = Inches(0.02)
    p_g = tf_g.paragraphs[0]
    p_g.alignment = PP_ALIGN.CENTER
    r_g = p_g.add_run()
    r_g.text = "🔐 全局安全与治理底座：Microsoft Purview 数据编织 · Azure Key Vault 统一密钥 · Entra ID 单点鉴权 · 全私有端点 (Private Link)"
    r_g.font.name = F.BODY
    r_g.font.size = Pt(7.0)
    r_g.font.color.rgb = C.AZURE_BLUE
    r_g.font.bold = True

    add_takeaway(
        slide,
        "核心设计原则",
        "存算分离、计算弹性归零、单一事实源 (Single Copy of Truth)、全链路私网互联与企业级零信任访问。"
    )


# ==============================================================================
# 7. SLIDE 4: SCENARIO-BASED COMPARISON MATRIX (多维场景深度选型矩阵)
# ==============================================================================
def make_slide_4_comparison(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(
        slide,
        tracker="Azure Data Platform Strategy | Technology Evaluation",
        action_title="场景深度选型：在数据接入、复杂转换、湖仓治理与 BI 报表场景下的边界权衡",
        lead_note="结合实际企业生产痛点，对 ADF、Databricks 与 Synapse 在五大核心技术场景进行系统权衡与评级。"
    )

    options = [
        {"name": "Azure Data Factory", "sub": "全域调度编排中枢", "highlight": False},
        {"name": "Azure Databricks", "sub": "现代湖仓与计算核心", "highlight": True},
        {"name": "Azure Synapse", "sub": "企业数仓与报表服务", "highlight": False},
    ]

    rows_data = [
        {
            "dim": "01 异构多源集成与编排",
            "desc": "连接器生态、定时轮询与事件依赖编排能力",
            "c": [
                ("●", "顶级编排中枢", "内置 100+ 原生连接器，支持图形化与 SHIR 穿透内网"),
                ("▲", "轻量作业编排", "Workflows 支持轻量调度，但外部异构连接需自写代码"),
                ("●", "继承 ADF 能力", "内置 Synapse Pipelines 具备与 ADF 一致的集成调度能力"),
            ]
        },
        {
            "dim": "02 大数据清洗与转换 (ETL)",
            "desc": "TB/PB 级吞吐、复杂业务逻辑、半/非结构化处理",
            "c": [
                ("▲", "轻量搬运为主", "Mapping Data Flow 成本高，更适合作为 ELT 搬运工具"),
                ("●", "公认性能霸主", "Photon 向量化执行提效 3-5x，PySpark 表达复杂业务逻辑"),
                ("▲", "适合 T-SQL 转换", "适合库内 SQL 存储过程，对非结构化与前沿算法支持受限"),
            ]
        },
        {
            "dim": "03 湖仓一体与流批统一",
            "desc": "ACID 事务、时间旅行回溯、流批一体化支持",
            "c": [
                ("○", "无存储处理能力", "仅作为数据搬运管道，不参与底层文件格式优化与治理"),
                ("●", "Delta 原创引领者", "原生支持 Delta Lake / Iceberg，Auto Loader 毫秒入湖"),
                ("▲", "支持部分 Delta", "Serverless SQL 支持读取 Delta，但深度优化依赖外部作业"),
            ]
        },
        {
            "dim": "04 高并发即席查询与 EDW",
            "desc": "千人并发点选、星型维度建模、亚秒级报表聚合",
            "c": [
                ("○", "不提供查询服务", "不具备 SQL 查询引擎，仅输出文件至下游目标库"),
                ("▲", "并发能力持续提升", "Databricks SQL 极佳，但超大规模高并发报表开销相对较高"),
                ("●", "企业数仓绝对主力", "专用 SQL 池 MPP 架构抗压强，物化视图与结果集缓存完善"),
            ]
        },
        {
            "dim": "05 AI / 机器学习与高级分析",
            "desc": "算法模型训练、特征工程、端到端 MLOps",
            "c": [
                ("○", "不支持算法建模", "无算法运行环境，仅为算法训练搬运样本文件"),
                ("●", "一站式数据智能底座", "原生集成 MLflow 模型注册与部署，内置 GPU 分布式算力"),
                ("▲", "集成 Azure ML", "具备轻量预测函数，但前沿深度学习生态成熟度不如 Databricks"),
            ]
        },
    ]

    header_y = G.BODY_Y
    header_h = 0.28
    dim_w = 1.70
    opt_w = (G.CONTENT_W - dim_w) / 3

    # Column Headers
    curr_x = G.MARGIN_LEFT + dim_w
    for opt in options:
        is_hl = opt["highlight"]
        h_box = draw_rect(slide, curr_x, header_y, opt_w - 0.05, header_h, fill_rgb=C.AZURE_BLUE if is_hl else C.BG_SLATE_LIGHT, line_rgb=C.AZURE_BLUE if is_hl else C.SOFT_BORDER, line_pt=1.0, shape_type=MSO_SHAPE.ROUNDED_RECTANGLE)
        tf = h_box.text_frame
        tf.word_wrap = True
        tf.margin_top = Inches(0.04)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        prefix = "★ 核心推荐: " if is_hl else ""
        r.text = f"{prefix}{opt['name']} ({opt['sub']})"
        r.font.name = F.BODY
        r.font.size = Pt(8.0)
        r.font.color.rgb = C.TEXT_WHITE if is_hl else C.TEXT_MAIN
        r.font.bold = True
        curr_x += opt_w

    # Rows
    curr_y = header_y + header_h + 0.04
    row_h = (G.BODY_H - header_h - 0.65) / len(rows_data)

    for row in rows_data:
        # Dimension Label
        dim_box = draw_rect(slide, G.MARGIN_LEFT, curr_y, dim_w - 0.05, row_h - 0.04, fill_rgb=C.BG_SLATE_LIGHT, line_rgb=C.LINE_SUBTLE, line_pt=0.75, shape_type=MSO_SHAPE.ROUNDED_RECTANGLE)
        tf_d = dim_box.text_frame
        tf_d.word_wrap = True
        tf_d.margin_top = Inches(0.04)
        tf_d.margin_left = tf_d.margin_right = Inches(0.06)
        p_dt = tf_d.paragraphs[0]
        r_dt = p_dt.add_run()
        r_dt.text = f"{row['dim']}\n"
        r_dt.font.name = F.TITLE
        r_dt.font.size = Pt(7.5)
        r_dt.font.color.rgb = C.AZURE_BLUE
        r_dt.font.bold = True

        p_dd = tf_d.add_paragraph()
        r_dd = p_dd.add_run()
        r_dd.text = row["desc"]
        r_dd.font.name = F.BODY
        r_dd.font.size = Pt(6.5)
        r_dd.font.color.rgb = C.TEXT_MUTED

        # Cells
        curr_cx = G.MARGIN_LEFT + dim_w
        for opt_idx, (sym, verd, detail) in enumerate(row["c"]):
            is_hl = options[opt_idx]["highlight"]
            c_box = draw_rect(slide, curr_cx, curr_y, opt_w - 0.05, row_h - 0.04, fill_rgb=C.ICE_BLUE if is_hl else C.BG_WHITE, line_rgb=C.AZURE_BLUE if is_hl else C.LINE_SUBTLE, line_pt=1.0 if is_hl else 0.5, shape_type=MSO_SHAPE.ROUNDED_RECTANGLE)
            tf_c = c_box.text_frame
            tf_c.word_wrap = True
            tf_c.margin_top = Inches(0.04)
            tf_c.margin_left = tf_c.margin_right = Inches(0.06)

            p_cs = tf_c.paragraphs[0]
            r_sym = p_cs.add_run()
            r_sym.text = f"{sym}  {verd}\n"
            r_sym.font.name = F.BODY
            r_sym.font.size = Pt(7.5)
            r_sym.font.bold = True
            if sym == "●":
                r_sym.font.color.rgb = C.SUCCESS
            elif sym == "▲":
                r_sym.font.color.rgb = C.WARNING
            else:
                r_sym.font.color.rgb = C.DANGER

            p_cd = tf_c.add_paragraph()
            r_cd = p_cd.add_run()
            r_cd.text = detail
            r_cd.font.name = F.BODY
            r_cd.font.size = Pt(6.5)
            r_cd.font.color.rgb = C.TEXT_BODY

            curr_cx += opt_w

        curr_y += row_h

    # Bottom Recommendation & Trade-off Box
    curr_y += 0.04
    rec_w = G.CONTENT_W * 0.60
    tra_w = G.CONTENT_W - rec_w - 0.12

    r_box = draw_rect(slide, G.MARGIN_LEFT, curr_y, rec_w, 0.45, fill_rgb=C.ICE_BLUE, line_rgb=C.AZURE_BLUE, line_pt=1.0, shape_type=MSO_SHAPE.ROUNDED_RECTANGLE)
    tf_r = r_box.text_frame
    tf_r.word_wrap = True
    tf_r.margin_left = Inches(0.08)
    p_r = tf_r.paragraphs[0]
    r_rt = p_r.add_run()
    r_rt.text = "★ 协同推荐结论：三位一体协同现代化架构\n"
    r_rt.font.name = F.BODY
    r_rt.font.size = Pt(8.0)
    r_rt.font.color.rgb = C.AZURE_BLUE
    r_rt.font.bold = True
    r_rb = p_r.add_run()
    r_rb.text = "ADF 负责全域调度与源端抽取，Databricks 负责湖仓 Medallion 流批处理与 AI，Synapse 承载企业数仓建模与报表。"
    r_rb.font.name = F.BODY
    r_rb.font.size = Pt(7.0)
    r_rb.font.color.rgb = C.TEXT_MAIN

    t_box = draw_rect(slide, G.MARGIN_LEFT + rec_w + 0.12, curr_y, tra_w, 0.45, fill_rgb=C.BG_SLATE_LIGHT, line_rgb=C.SOFT_BORDER, line_pt=0.75, shape_type=MSO_SHAPE.ROUNDED_RECTANGLE)
    tf_t = t_box.text_frame
    tf_t.word_wrap = True
    tf_t.margin_left = Inches(0.08)
    p_t = tf_t.paragraphs[0]
    r_tt = p_t.add_run()
    r_tt.text = "⚖ 关键取舍 (Trade-off)\n"
    r_tt.font.name = F.BODY
    r_tt.font.size = Pt(8.0)
    r_tt.font.color.rgb = C.TEXT_MAIN
    r_tt.font.bold = True
    r_tb = p_t.add_run()
    r_tb.text = "三者结合需配置统一元数据治理 (Purview) 与网络私有端点，但换取了极佳的性能与长远可演进性。"
    r_tb.font.name = F.BODY
    r_tb.font.size = Pt(7.0)
    r_tb.font.color.rgb = C.TEXT_BODY


# ==============================================================================
# 8. SLIDE 5: MODERN DATA LIFECYCLE & MEDALLION PIPELINE (端到端数据流水线流转)
# ==============================================================================
def make_slide_5_pipeline(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(
        slide,
        tracker="Azure Data Platform Implementation | Data Lifecycle",
        action_title="数据流转闭环：从源端抽取到商业洞察的端到端 4 阶段落地工作流",
        lead_note="清晰界定 ADF 调度触发、Databricks 增量精炼与 Synapse 视图建模的职责交接面与流转机制。"
    )

    stages = [
        {
            "num": "STAGE 01",
            "name": "多源统一抽取与落地",
            "tech": "ADF · Event Hubs · Landing Zone",
            "bullets": [
                "ADF 定时轮询源端增量数据 (Watermark)",
                "落入 ADLS Gen2 Landing/Raw 分区目录",
                "写入抽取审计流水，核对行数与校验和",
            ],
            "highlight": False,
        },
        {
            "num": "STAGE 02",
            "name": "湖仓自动化精炼清洗",
            "tech": "Databricks · Auto Loader · Delta",
            "bullets": [
                "Auto Loader 事件感应新文件即时入湖",
                "写入 Bronze -> Silver 执行清洗、去重与类型约束",
                "Great Expectations 质检规则拦截异常脏数据",
            ],
            "highlight": True,
        },
        {
            "num": "STAGE 03",
            "name": "业务聚合建模与特征",
            "tech": "Databricks · Gold Delta · Features",
            "bullets": [
                "跨系统关联组装客户、订单、产品等主题宽表",
                "统一口径预计算核心 KPI 指标库与特征库",
                "Unity Catalog 登记元数据并对外发布共享",
            ],
            "highlight": False,
        },
        {
            "num": "STAGE 04",
            "name": "数仓加速与商业呈现",
            "tech": "Synapse Dedicated · Power BI",
            "bullets": [
                "PolyBase / Synapse Link 高速加载专用池",
                "Power BI DirectLake 毫秒级直连直读湖仓",
                "物化视图加速保障数千业务用户并发无卡顿",
            ],
            "highlight": False,
        },
    ]

    card_y = G.BODY_Y
    card_h = 1.85
    card_gap = 0.15
    card_w = (G.CONTENT_W - card_gap * 3) / 4

    for i, st in enumerate(stages):
        cx = G.MARGIN_LEFT + i * (card_w + card_gap)
        is_hl = st["highlight"]

        s_box = draw_rect(slide, cx, card_y, card_w, card_h, fill_rgb=C.ICE_BLUE if is_hl else C.BG_WHITE, line_rgb=C.AZURE_BLUE if is_hl else C.SOFT_BORDER, line_pt=1.5 if is_hl else 1.0, shape_type=MSO_SHAPE.ROUNDED_RECTANGLE)

        tf = s_box.text_frame
        tf.word_wrap = True
        tf.margin_top = Inches(0.08)
        tf.margin_bottom = Inches(0.06)
        tf.margin_left = tf.margin_right = Inches(0.08)

        p0 = tf.paragraphs[0]
        r0 = p0.add_run()
        r0.text = f"{st['num']}\n"
        r0.font.name = F.TITLE
        r0.font.size = Pt(8.5)
        r0.font.color.rgb = C.AZURE_BLUE
        r0.font.bold = True

        p1 = tf.add_paragraph()
        r1 = p1.add_run()
        r1.text = f"{st['name']}\n"
        r1.font.name = F.TITLE
        r1.font.size = Pt(9.5)
        r1.font.color.rgb = C.TEXT_MAIN
        r1.font.bold = True

        p2 = tf.add_paragraph()
        p2.space_before = Pt(2.0)
        p2.space_after = Pt(4.0)
        r2 = p2.add_run()
        r2.text = st["tech"]
        r2.font.name = F.MONO
        r2.font.size = Pt(6.5)
        r2.font.color.rgb = C.AZURE_BLUE if is_hl else C.TEXT_MUTED

        for b in st["bullets"]:
            pb = tf.add_paragraph()
            pb.space_before = Pt(2.0)
            rb = pb.add_run()
            rb.text = f"• {b}"
            rb.font.name = F.BODY
            rb.font.size = Pt(7.0)
            rb.font.color.rgb = C.TEXT_BODY

    # Lower Half: 3 Creative Quantitative Metric Badges
    kpis_y = card_y + card_h + 0.15
    kpis_h = 1.30
    kpi_gap = 0.20
    kpi_w = (G.CONTENT_W - kpi_gap * 2) / 3

    kpis = [
        ("⚡ 15 分钟", "数据端到端入库就绪时延", "从传统 T+1 日级批处理跨越至分钟级增量可用，全面支撑业务敏捷运营与实时分析。"),
        ("💰 -35%", "云端计算与存储 TCO 降低", "存算分离架构 + Serverless 弹性缩容归零，彻底杜绝常驻空转算力与冗余存储浪费。"),
        ("📉 -70%", "手工运维排错工单压缩", "端到端自动化管道 + 统一元数据治理 (Purview)，大幅减少数据对账与链路排查负担。"),
    ]

    for j, (val, title, desc) in enumerate(kpis):
        kx = G.MARGIN_LEFT + j * (kpi_w + kpi_gap)
        k_box = draw_rect(slide, kx, kpis_y, kpi_w, kpis_h, fill_rgb=C.BG_SLATE_LIGHT, line_rgb=C.SOFT_BORDER, line_pt=0.75, shape_type=MSO_SHAPE.ROUNDED_RECTANGLE)

        tf_k = k_box.text_frame
        tf_k.word_wrap = True
        tf_k.margin_top = Inches(0.08)
        tf_k.margin_left = tf_k.margin_right = Inches(0.10)

        pk0 = tf_k.paragraphs[0]
        rk0 = pk0.add_run()
        rk0.text = f"{val}\n"
        rk0.font.name = F.TITLE
        rk0.font.size = Pt(18.0)
        rk0.font.color.rgb = C.AZURE_BLUE
        rk0.font.bold = True

        pk1 = tf_k.add_paragraph()
        pk1.space_before = Pt(2.0)
        rk1 = pk1.add_run()
        rk1.text = f"{title}\n"
        rk1.font.name = F.TITLE
        rk1.font.size = Pt(9.0)
        rk1.font.color.rgb = C.TEXT_MAIN
        rk1.font.bold = True

        pk2 = tf_k.add_paragraph()
        pk2.space_before = Pt(2.0)
        rk2 = pk2.add_run()
        rk2.text = desc
        rk2.font.name = F.BODY
        rk2.font.size = Pt(7.0)
        rk2.font.color.rgb = C.TEXT_BODY

    add_takeaway(
        slide,
        "流程收益结论",
        "以流批一体闭环代替散落手工脚本，打通从源端感知到业务报表毫秒呈现的全流程，兼备敏捷度、高可靠性与合规性。"
    )


# ==============================================================================
# 9. MAIN ORCHESTRATOR
# ==============================================================================
def main():
    prs = Presentation()
    prs.slide_width = Inches(G.WIDTH)
    prs.slide_height = Inches(G.HEIGHT)

    print("=" * 60)
    print("Generating Creative Azure Architecture Presentation")
    print("=" * 60)

    print("Building Slide 1: Creative Hero Cover...")
    make_slide_1_cover(prs)

    print("Building Slide 2: The Triad Role Matrix...")
    make_slide_2_roles(prs)

    print("Building Slide 3: Blueprint-Grade Technical Architecture...")
    make_slide_3_architecture(prs)

    print("Building Slide 4: Comparative Evaluation Matrix...")
    make_slide_4_comparison(prs)

    print("Building Slide 5: Modern Data Lifecycle & Pipeline...")
    make_slide_5_pipeline(prs)

    out_file = Path("Azure_Databricks_ADF_Synapse_Architecture.pptx")
    prs.save(str(out_file))
    print(f"✅ Generated {len(prs.slides)} slides to {out_file}!")
    print("=" * 60)


if __name__ == "__main__":
    main()
