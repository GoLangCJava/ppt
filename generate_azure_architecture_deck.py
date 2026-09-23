#!/usr/bin/env python3
"""
Custom Presentation Generator: Azure Databricks & ADF & Synapse Architecture.
Strict Content-Driven Layout Engine with Zero-Overlap Geometry Guarantee.
"""

from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN


# ==========================================
# 1. COLOR & DESIGN SYSTEM TOKENS
# ==========================================
class Palette:
    # Primary Consulting Blues
    PRIMARY = RGBColor(0, 156, 222)       # #009CDE - Vibrant Azure/Consulting Blue
    PRIMARY_DARK = RGBColor(0, 90, 156)   # #005A9C - Deep Navy Blue
    PRIMARY_LIGHT = RGBColor(234, 244, 251) # #EAF4FB - Light Tint Background
    PRIMARY_BORDER = RGBColor(127, 205, 238) # #7FCDEE - Soft Accent Border
    
    # Secondary & Accent Colors
    ACCENT_PURPLE = RGBColor(124, 58, 237) # #7C3AED - AI / ML Highlight
    SUCCESS = RGBColor(16, 185, 129)       # #10B981 - Strong Capability
    WARNING = RGBColor(245, 158, 11)       # #F59E0B - Moderate / Partial
    DANGER = RGBColor(239, 68, 68)         # #EF4444 - Limitation / Not Applicable
    
    # Surfaces & Backgrounds
    BG_DARK = RGBColor(15, 23, 42)         # #0F172A - Executive Slate Dark
    BG_CARD = RGBColor(255, 255, 255)      # Pure White Card
    BG_TINT = RGBColor(248, 250, 252)      # #F8FAFC - Soft Slate White
    BG_CONTAINER = RGBColor(241, 245, 249) # #F1F5F9 - Container Gray
    
    # Borders & Lines
    BORDER_SUBTLE = RGBColor(226, 232, 240) # #E2E8F0 - Subtle Line
    BORDER_MUTED = RGBColor(203, 213, 225)  # #CBD5E1 - Muted Border
    
    # Text Hierarchy
    TEXT_PRIMARY = RGBColor(15, 23, 42)    # #0F172A - Primary Title / Dark Body
    TEXT_SECONDARY = RGBColor(51, 65, 85)  # #334155 - Secondary Body
    TEXT_MUTED = RGBColor(100, 116, 139)   # #64748B - Captions & Trackers
    TEXT_WHITE = RGBColor(255, 255, 255)   # White


class Typography:
    FONT_TITLE = "Georgia"
    FONT_BODY = "Microsoft YaHei"
    FONT_MONO = "Consolas"


# ==========================================
# 2. GEOMETRY & LAYOUT ENGINE (ZERO OVERLAP)
# ==========================================
class Layout:
    WIDTH = 10.0
    HEIGHT = 5.625
    MARGIN_LEFT = 0.55
    MARGIN_RIGHT = 0.55
    CONTENT_WIDTH = WIDTH - MARGIN_LEFT - MARGIN_RIGHT  # 8.90 inches

    # Header Zones (Strictly Disjoint Y-ranges)
    TRACKER_TOP = 0.28
    TRACKER_HEIGHT = 0.18    # ends at 0.46
    
    TITLE_TOP = 0.49
    TITLE_HEIGHT = 0.38      # ends at 0.87
    
    LEAD_TOP = 0.90
    LEAD_HEIGHT = 0.28       # ends at 1.18
    
    DIVIDER_Y = 1.22
    
    # Body Zone
    BODY_TOP = 1.32
    BODY_HEIGHT = 3.48       # ends at 4.80
    
    # Footer Zone
    FOOTER_TOP = 4.88
    FOOTER_HEIGHT = 0.42     # ends at 5.30


def add_header(slide, tracker: str, action_title: str, lead_note: str):
    """Draws standard consulting 3-tier header with strictly zero overlapping boxes."""
    # 1. Tracker Box
    tb_track = slide.shapes.add_textbox(
        Inches(Layout.MARGIN_LEFT),
        Inches(Layout.TRACKER_TOP),
        Inches(Layout.CONTENT_WIDTH),
        Inches(Layout.TRACKER_HEIGHT),
    )
    tf_tr = tb_track.text_frame
    tf_tr.word_wrap = True
    tf_tr.margin_top = Inches(0)
    tf_tr.margin_bottom = Inches(0)
    tf_tr.margin_left = Inches(0)
    tf_tr.margin_right = Inches(0)
    p_tr = tf_tr.paragraphs[0]
    r_tr = p_tr.add_run()
    r_tr.text = tracker.upper()
    r_tr.font.name = Typography.FONT_TITLE
    r_tr.font.size = Pt(7.5)
    r_tr.font.color.rgb = Palette.TEXT_MUTED
    r_tr.font.bold = False

    # 2. Main Action Title
    tb_title = slide.shapes.add_textbox(
        Inches(Layout.MARGIN_LEFT),
        Inches(Layout.TITLE_TOP),
        Inches(Layout.CONTENT_WIDTH),
        Inches(Layout.TITLE_HEIGHT),
    )
    tf_t = tb_title.text_frame
    tf_t.word_wrap = True
    tf_t.margin_top = Inches(0)
    tf_t.margin_bottom = Inches(0)
    tf_t.margin_left = Inches(0)
    tf_t.margin_right = Inches(0)
    p_t = tf_t.paragraphs[0]
    r_t = p_t.add_run()
    r_t.text = action_title
    r_t.font.name = Typography.FONT_TITLE
    r_t.font.size = Pt(15.0)
    r_t.font.color.rgb = Palette.TEXT_PRIMARY
    r_t.font.bold = True

    # 3. Lead Note
    tb_lead = slide.shapes.add_textbox(
        Inches(Layout.MARGIN_LEFT),
        Inches(Layout.LEAD_TOP),
        Inches(Layout.CONTENT_WIDTH),
        Inches(Layout.LEAD_HEIGHT),
    )
    tf_l = tb_lead.text_frame
    tf_l.word_wrap = True
    tf_l.margin_top = Inches(0)
    tf_l.margin_bottom = Inches(0)
    tf_l.margin_left = Inches(0)
    tf_l.margin_right = Inches(0)
    p_l = tf_l.paragraphs[0]
    r_l = p_l.add_run()
    r_l.text = lead_note
    r_l.font.name = Typography.FONT_BODY
    r_l.font.size = Pt(10.5)
    r_l.font.color.rgb = Palette.PRIMARY_DARK
    r_l.font.bold = False

    # 4. Divider Line
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(Layout.MARGIN_LEFT),
        Inches(Layout.DIVIDER_Y),
        Inches(Layout.CONTENT_WIDTH),
        Pt(0.75),
    )
    line.fill.solid()
    line.fill.fore_color.rgb = Palette.BORDER_SUBTLE
    line.line.fill.background()


def add_takeaway(slide, title: str, text: str):
    """Draws standard takeaway card at the bottom of the slide."""
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(Layout.MARGIN_LEFT),
        Inches(Layout.FOOTER_TOP),
        Inches(Layout.CONTENT_WIDTH),
        Inches(Layout.FOOTER_HEIGHT),
    )
    card.fill.solid()
    card.fill.fore_color.rgb = Palette.PRIMARY_LIGHT
    card.line.color.rgb = Palette.PRIMARY
    card.line.width = Pt(1.0)

    tf = card.text_frame
    tf.word_wrap = True
    tf.margin_top = Inches(0.06)
    tf.margin_bottom = Inches(0.06)
    tf.margin_left = Inches(0.15)
    tf.margin_right = Inches(0.15)

    p = tf.paragraphs[0]
    r_t = p.add_run()
    r_t.text = f"【{title}】 "
    r_t.font.name = Typography.FONT_BODY
    r_t.font.size = Pt(8.5)
    r_t.font.color.rgb = Palette.PRIMARY_DARK
    r_t.font.bold = True

    r_b = p.add_run()
    r_b.text = text
    r_b.font.name = Typography.FONT_BODY
    r_b.font.size = Pt(8.5)
    r_b.font.color.rgb = Palette.TEXT_PRIMARY
    r_b.font.bold = False


# ==========================================
# 3. BESPOKE SLIDE BUILDERS
# ==========================================

def build_slide_1_cover(prs):
    """Slide 1: Executive Cover Slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Dark Executive Background
    bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0), Inches(0), Inches(Layout.WIDTH), Inches(Layout.HEIGHT)
    )
    bg.fill.solid()
    bg.fill.fore_color.rgb = Palette.BG_DARK
    bg.line.fill.background()

    # Left Vibrant Blue Accent Bar
    bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(Layout.MARGIN_LEFT), Inches(1.20), Inches(0.16), Inches(2.80)
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = Palette.PRIMARY
    bar.line.fill.background()

    content_x = Layout.MARGIN_LEFT + 0.35
    content_w = Layout.CONTENT_WIDTH - 0.35

    # Content in a single text box for natural vertical flow
    tb = slide.shapes.add_textbox(
        Inches(content_x), Inches(1.25), Inches(content_w), Inches(3.20)
    )
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_top = Inches(0)
    tf.margin_bottom = Inches(0)
    tf.margin_left = Inches(0)
    tf.margin_right = Inches(0)

    # 1. Domain Tag
    p_tag = tf.paragraphs[0]
    r_tag = p_tag.add_run()
    r_tag.text = "ENTERPRISE CLOUD DATA ARCHITECTURE ADVISORY\n"
    r_tag.font.name = Typography.FONT_TITLE
    r_tag.font.size = Pt(11.0)
    r_tag.font.color.rgb = Palette.PRIMARY
    r_tag.font.bold = True

    # 2. Main Title
    p_title = tf.add_paragraph()
    p_title.space_before = Pt(8.0)
    r_title = p_title.add_run()
    r_title.text = "Azure Modern Data Platform Architecture"
    r_title.font.name = Typography.FONT_TITLE
    r_title.font.size = Pt(26.0)
    r_title.font.color.rgb = Palette.TEXT_WHITE
    r_title.font.bold = True

    # 3. Subtitle
    p_sub = tf.add_paragraph()
    p_sub.space_before = Pt(10.0)
    r_sub = p_sub.add_run()
    r_sub.text = "基于 Azure Databricks · ADF · Synapse 的端到端协同技术架构与选型方案建议书"
    r_sub.font.name = Typography.FONT_BODY
    r_sub.font.size = Pt(13.0)
    r_sub.font.color.rgb = Palette.PRIMARY_LIGHT
    r_sub.font.bold = False

    # 4. Metadata at bottom
    tb_meta = slide.shapes.add_textbox(
        Inches(content_x), Inches(4.70), Inches(content_w), Inches(0.40)
    )
    tf_m = tb_meta.text_frame
    tf_m.word_wrap = True
    p_m = tf_m.paragraphs[0]
    r_m = p_m.add_run()
    r_m.text = "Cloud Architecture & Data Practice  |  September 2026  |  Confidential"
    r_m.font.name = Typography.FONT_BODY
    r_m.font.size = Pt(9.0)
    r_m.font.color.rgb = Palette.TEXT_MUTED


def build_slide_2_positioning(prs):
    """Slide 2: Strategic Positioning & Differentiators."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(
        slide,
        tracker="Azure Data Platform Strategy | Core Roles & Positioning",
        action_title="战略定位：ADF、Databricks 与 Synapse 各司其职，构建现代湖仓一体分工体系",
        lead_note="明确三者在企业数据价值链中的专业分工，打破选型困惑：ADF管调度接入、Databricks管湖仓计算、Synapse管数仓与报表。"
    )

    col_gap = 0.20
    col_w = (Layout.CONTENT_WIDTH - col_gap * 2) / 3
    body_y = Layout.BODY_TOP
    body_h = Layout.BODY_HEIGHT - 0.52

    cards_data = [
        {
            "badge": "01",
            "name": "Azure Data Factory (ADF)",
            "role": "全域数据集成与调度编排中枢",
            "highlight": False,
            "bullets": [
                ("多源连接生态：", "内置 100+ 原生成熟连接器，覆盖数据库、多云与 SaaS 应用。"),
                ("Serverless 弹性调度：", "基于事件触发、定时轮询与参数化依赖编排，无基础设施开销。"),
                ("安全混合网络：", "自托管运行时 (SHIR) 打通内网与专线，全链路私有端点通信。"),
            ],
            "best_for": "跨网络数据抽取、轻量文件落湖、作业依赖触发与全域调度。",
            "not_for": "极复杂大数据清洗、非结构化计算与高并发即席查询。",
        },
        {
            "badge": "02",
            "name": "Azure Databricks",
            "role": "现代化湖仓一体与智能计算核心",
            "highlight": True,
            "bullets": [
                ("Medallion 奖章架构：", "驱动 Bronze 原始 -> Silver 清洗 -> Gold 业务宽表的逐层精炼。"),
                ("极致大数据性能：", "Photon 向量化执行引擎 + Delta Lake ACID 事务与时间旅行。"),
                ("统一数据与 AI 治理：", "Unity Catalog 提供表/行/列细粒度权限，原生集成 MLflow。"),
            ],
            "best_for": "流批一体计算、TB/PB 级复杂转换、特征工程与企业私域 AI 模型。",
            "not_for": "简单的点对点文件搬运与数千并发分析师同时在线交互点选。",
        },
        {
            "badge": "03",
            "name": "Azure Synapse Analytics",
            "role": "企业级数仓与高性能 BI 报表服务",
            "highlight": False,
            "bullets": [
                ("专用 SQL 池 (Dedicated)：", "MPP 分布式高并发架构，承载核心企业数仓与星型模型。"),
                ("无服务器 SQL (Serverless)：", "针对湖内文件直接执行原生 T-SQL 探索，按查询量计费。"),
                ("无缝 BI 体验：", "与 Power BI 深度集成，物化视图与 DirectLake 支撑秒级展现。"),
            ],
            "best_for": "高并发管理仪表盘、复杂维度建模、企业财务级口径治理。",
            "not_for": "底层非结构化音视频处理与前沿深度学习模型分布式训练。",
        },
    ]

    for i, c in enumerate(cards_data):
        col_x = Layout.MARGIN_LEFT + i * (col_w + col_gap)
        is_hl = c["highlight"]
        
        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(col_x), Inches(body_y), Inches(col_w), Inches(body_h)
        )
        card.fill.solid()
        card.fill.fore_color.rgb = Palette.PRIMARY_LIGHT if is_hl else Palette.BG_CARD
        card.line.color.rgb = Palette.PRIMARY if is_hl else Palette.BORDER_SUBTLE
        card.line.width = Pt(1.5 if is_hl else 1.0)

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_top = Inches(0.12)
        tf.margin_bottom = Inches(0.10)
        tf.margin_left = Inches(0.14)
        tf.margin_right = Inches(0.14)

        # 1. Badge & Product Name
        p0 = tf.paragraphs[0]
        r_badge = p0.add_run()
        r_badge.text = f"[{c['badge']}]  "
        r_badge.font.name = Typography.FONT_TITLE
        r_badge.font.size = Pt(10.0)
        r_badge.font.color.rgb = Palette.PRIMARY_DARK
        r_badge.font.bold = True

        r_name = p0.add_run()
        r_name.text = c["name"]
        r_name.font.name = Typography.FONT_TITLE
        r_name.font.size = Pt(11.0)
        r_name.font.color.rgb = Palette.TEXT_PRIMARY
        r_name.font.bold = True

        # 2. Core Role
        p_role = tf.add_paragraph()
        p_role.space_before = Pt(3.0)
        p_role.space_after = Pt(8.0)
        r_role = p_role.add_run()
        r_role.text = c["role"]
        r_role.font.name = Typography.FONT_BODY
        r_role.font.size = Pt(8.5)
        r_role.font.color.rgb = Palette.PRIMARY_DARK if is_hl else Palette.TEXT_MUTED
        r_role.font.bold = True

        # 3. Bullets
        for prefix, body in c["bullets"]:
            p_b = tf.add_paragraph()
            p_b.space_before = Pt(3.0)
            p_b.space_after = Pt(2.0)
            
            r_bp = p_b.add_run()
            r_bp.text = f"• {prefix}"
            r_bp.font.name = Typography.FONT_BODY
            r_bp.font.size = Pt(7.5)
            r_bp.font.color.rgb = Palette.TEXT_PRIMARY
            r_bp.font.bold = True

            r_bb = p_b.add_run()
            r_bb.text = body
            r_bb.font.name = Typography.FONT_BODY
            r_bb.font.size = Pt(7.5)
            r_bb.font.color.rgb = Palette.TEXT_SECONDARY
            r_bb.font.bold = False

        # 4. Scenario Guide Box
        p_guide = tf.add_paragraph()
        p_guide.space_before = Pt(8.0)
        r_g1 = p_guide.add_run()
        r_g1.text = "✓ 适用场景: "
        r_g1.font.name = Typography.FONT_BODY
        r_g1.font.size = Pt(7.0)
        r_g1.font.color.rgb = Palette.SUCCESS
        r_g1.font.bold = True

        r_g2 = p_guide.add_run()
        r_g2.text = f"{c['best_for']}\n"
        r_g2.font.name = Typography.FONT_BODY
        r_g2.font.size = Pt(7.0)
        r_g2.font.color.rgb = Palette.TEXT_SECONDARY

        r_g3 = p_guide.add_run()
        r_g3.text = "✕ 边界限制: "
        r_g3.font.name = Typography.FONT_BODY
        r_g3.font.size = Pt(7.0)
        r_g3.font.color.rgb = Palette.DANGER
        r_g3.font.bold = True

        r_g4 = p_guide.add_run()
        r_g4.text = c["not_for"]
        r_g4.font.name = Typography.FONT_BODY
        r_g4.font.size = Pt(7.0)
        r_g4.font.color.rgb = Palette.TEXT_SECONDARY

    add_takeaway(
        slide,
        "架构协同原则",
        "绝非三选一的替代关系：以 ADF 为全域数据管道纽带，以 Databricks 为中枢加工单一事实源 (SSOT)，以 Synapse 为窗口保障企业级报表毫秒级响应。"
    )


def build_slide_3_architecture(prs):
    """Slide 3: End-to-End Enterprise Technical Architecture Blueprint."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(
        slide,
        tracker="Azure Data Platform Architecture | Technical Blueprint",
        action_title="技术架构全景图：基于 Medallion 奖章体系的 Azure 四层现代化数据架构",
        lead_note="分层解耦：多源接入、湖仓精炼、企业数仓建模与全场景消费深度协同，贯穿端到端安全与治理基线。"
    )

    # 4 Layers + 3 Protocol Connectors + 1 Cross-cutting Baseline
    layers_data = [
        {
            "num": "01",
            "name": "消费洞察层",
            "en": "CONSUMPTION",
            "connector": "HTTPS 443 · DirectLake / DirectQuery · REST APIs · OData",
            "modules": [
                ("Power BI 敏捷可视化", "直连 Synapse 与 Databricks SQL，秒级多维钻取分析与自助看板"),
                ("企业级报表与仪表盘", "合规财务报表、运营日报与管理驾驶舱，支持定时分发与导出"),
                ("AI 与大模型应用", "基于 Azure OpenAI 知识库语义问答与 MLflow 模型在线推理"),
                ("开放 API 与逆向 ETL", "API Management 统一封装标准数据微服务，反哺业务系统"),
            ]
        },
        {
            "num": "02",
            "name": "建模服务层",
            "en": "SERVING & EDW",
            "connector": "Synapse Link · JDBC / ODBC · T-SQL · Delta Sharing 联邦访问",
            "modules": [
                ("Synapse 专用 SQL 池", "MPP 分布式企业数仓，承载星型模型 (Fact/Dim)，毫秒高并发"),
                ("Synapse Serverless", "针对数据湖开放文件直接执行原生 T-SQL，零驻留成本即席探索"),
                ("Databricks SQL 仓库", "Serverless SQL Warehouse，提供高性价比湖仓交互式查询"),
                ("统一指标管理中枢", "统一维护全局商业指标口径与语义层，消除跨部门数据歧义"),
            ]
        },
        {
            "num": "03",
            "name": "湖仓计算层",
            "en": "LAKEHOUSE ENGINE",
            "connector": "ABFS (ADLS Gen2) · Delta Lake 开放格式 · Apache Spark 3.x · Auto Loader",
            "highlight": True,
            "modules": [
                ("Bronze 原始湖 (Raw)", "保留源端 1:1 原始快照与增量日志，Auto Loader 毫秒入湖"),
                ("Silver 洁净层 (Refined)", "执行清洗、去重、类型推断、Schema 校验与实体关联 (Delta 表)"),
                ("Gold 业务层 (Curated)", "高度精炼的主题汇总宽表、业务指标预计算与特征库，ACID 保障"),
                ("Unity Catalog 统一治理", "跨工作区统一数据目录，支持表/行/列级授权与端到端血缘追踪"),
            ]
        },
        {
            "num": "04",
            "name": "接入基建层",
            "en": "INGESTION & INFRA",
            "connector": None,
            "modules": [
                ("Azure Data Factory (ADF)", "全域统一调度编排流水线，支持 100+ 连接器与断点续传"),
                ("Event Hubs 实时流接入", "高吞吐接入 IoT 设备埋点、应用日志与 Kafka 实时数据流"),
                ("ADLS Gen2 湖存储底座", "高性价比分层存储 (Hot/Cool/Archive)，多协议兼容与加密"),
                ("多源异构业务系统", "SAP ERP、本地核心数据库、SaaS CRM、无结构音视频附件"),
            ]
        },
    ]

    curr_y = Layout.BODY_TOP
    layer_h = 0.64
    proto_h = 0.12
    label_w = 1.30
    mod_start_x = Layout.MARGIN_LEFT + label_w + 0.12
    mod_total_w = Layout.CONTENT_WIDTH - label_w - 0.12

    for l_idx, layer in enumerate(layers_data):
        is_hl = layer.get("highlight", False)

        # 1. Left Layer Label Box
        l_box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(Layout.MARGIN_LEFT), Inches(curr_y), Inches(label_w), Inches(layer_h)
        )
        l_box.fill.solid()
        l_box.fill.fore_color.rgb = Palette.PRIMARY_LIGHT if is_hl else Palette.BG_TINT
        l_box.line.color.rgb = Palette.PRIMARY if is_hl else Palette.BORDER_MUTED
        l_box.line.width = Pt(1.5 if is_hl else 0.75)

        tf_l = l_box.text_frame
        tf_l.word_wrap = True
        tf_l.margin_top = Inches(0.06)
        tf_l.margin_bottom = Inches(0.04)
        tf_l.margin_left = Inches(0.08)
        tf_l.margin_right = Inches(0.06)

        p_ln = tf_l.paragraphs[0]
        r_ln = p_ln.add_run()
        r_ln.text = f"{layer['num']} {layer['name']}\n"
        r_ln.font.name = Typography.FONT_TITLE
        r_ln.font.size = Pt(8.5)
        r_ln.font.color.rgb = Palette.PRIMARY_DARK
        r_ln.font.bold = True

        p_le = tf_l.add_paragraph()
        r_le = p_le.add_run()
        r_le.text = layer["en"]
        r_le.font.name = Typography.FONT_BODY
        r_le.font.size = Pt(6.5)
        r_le.font.color.rgb = Palette.TEXT_MUTED

        # 2. Modules in Layer
        mods = layer["modules"]
        num_m = len(mods)
        mod_gap = 0.08
        mod_w = (mod_total_w - (mod_gap * (num_m - 1))) / num_m

        for m_idx, (m_title, m_desc) in enumerate(mods):
            m_x = mod_start_x + m_idx * (mod_w + mod_gap)
            m_box = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(m_x), Inches(curr_y), Inches(mod_w), Inches(layer_h)
            )
            m_box.fill.solid()
            m_box.fill.fore_color.rgb = Palette.BG_CARD
            m_box.line.color.rgb = Palette.PRIMARY_BORDER if is_hl else Palette.BORDER_SUBTLE
            m_box.line.width = Pt(1.0 if is_hl else 0.75)

            tf_m = m_box.text_frame
            tf_m.word_wrap = True
            tf_m.margin_top = Inches(0.05)
            tf_m.margin_bottom = Inches(0.04)
            tf_m.margin_left = Inches(0.06)
            tf_m.margin_right = Inches(0.06)

            p_mt = tf_m.paragraphs[0]
            r_mt = p_mt.add_run()
            r_mt.text = m_title
            r_mt.font.name = Typography.FONT_TITLE
            r_mt.font.size = Pt(7.5)
            r_mt.font.color.rgb = Palette.PRIMARY_DARK
            r_mt.font.bold = True

            p_md = tf_m.add_paragraph()
            p_md.space_before = Pt(1.5)
            r_md = p_md.add_run()
            r_md.text = m_desc
            r_md.font.name = Typography.FONT_BODY
            r_md.font.size = Pt(6.5)
            r_md.font.color.rgb = Palette.TEXT_SECONDARY

        curr_y += layer_h

        # 3. Protocol connector line between layers
        if layer["connector"]:
            tb_proto = slide.shapes.add_textbox(
                Inches(mod_start_x), Inches(curr_y), Inches(mod_total_w), Inches(proto_h)
            )
            tf_p = tb_proto.text_frame
            tf_p.word_wrap = False
            tf_p.margin_top = Inches(0)
            tf_p.margin_bottom = Inches(0)
            tf_p.margin_left = Inches(0)
            tf_p.margin_right = Inches(0)
            p_pr = tf_p.paragraphs[0]
            p_pr.alignment = PP_ALIGN.CENTER
            r_pr = p_pr.add_run()
            r_pr.text = f"↕  {layer['connector']}  ↕"
            r_pr.font.name = Typography.FONT_MONO
            r_pr.font.size = Pt(6.5)
            r_pr.font.color.rgb = Palette.TEXT_MUTED
            curr_y += proto_h

    # Cross-cutting Security & Governance Foundation Bar
    curr_y += 0.04
    gov_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(Layout.MARGIN_LEFT), Inches(curr_y), Inches(Layout.CONTENT_WIDTH), Inches(0.24)
    )
    gov_box.fill.solid()
    gov_box.fill.fore_color.rgb = Palette.BG_CONTAINER
    gov_box.line.color.rgb = Palette.BORDER_MUTED
    gov_box.line.width = Pt(0.75)

    tf_g = gov_box.text_frame
    tf_g.word_wrap = True
    tf_g.margin_top = Inches(0.02)
    tf_g.margin_bottom = Inches(0.02)
    tf_g.margin_left = Inches(0.10)
    p_g = tf_g.paragraphs[0]
    p_g.alignment = PP_ALIGN.CENTER
    r_g = p_g.add_run()
    r_g.text = "🛡️ 统一治理与安全底座：Microsoft Purview 数据编织 · Azure Key Vault 统一密钥 · Entra ID 单点鉴权 · 全私有端点 (Private Link)"
    r_g.font.name = Typography.FONT_BODY
    r_g.font.size = Pt(7.0)
    r_g.font.color.rgb = Palette.PRIMARY_DARK
    r_g.font.bold = True

    add_takeaway(
        slide,
        "核心设计原则",
        "存算分离、计算弹性归零、单一事实源 (Single Copy of Truth)、全链路私网互联与企业级零信任访问。"
    )


def build_slide_4_comparison(prs):
    """Slide 4: Deep Comparative Evaluation Matrix."""
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

    header_y = Layout.BODY_TOP
    header_h = 0.28
    dim_w = 1.70
    opt_w = (Layout.CONTENT_WIDTH - dim_w) / 3

    # Column Headers
    curr_x = Layout.MARGIN_LEFT + dim_w
    for opt in options:
        is_hl = opt["highlight"]
        h_box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(curr_x), Inches(header_y), Inches(opt_w - 0.05), Inches(header_h)
        )
        h_box.fill.solid()
        h_box.fill.fore_color.rgb = Palette.PRIMARY if is_hl else Palette.BG_TINT
        h_box.line.color.rgb = Palette.PRIMARY if is_hl else Palette.BORDER_MUTED
        h_box.line.width = Pt(1.0)

        tf = h_box.text_frame
        tf.word_wrap = True
        tf.margin_top = Inches(0.04)
        tf.margin_bottom = Inches(0.02)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        prefix = "★ 核心推荐: " if is_hl else ""
        r.text = f"{prefix}{opt['name']} ({opt['sub']})"
        r.font.name = Typography.FONT_BODY
        r.font.size = Pt(8.0)
        r.font.color.rgb = Palette.TEXT_WHITE if is_hl else Palette.TEXT_PRIMARY
        r.font.bold = True
        curr_x += opt_w

    # Rows
    curr_y = header_y + header_h + 0.04
    row_h = (Layout.BODY_HEIGHT - header_h - 0.72) / len(rows_data)

    for row in rows_data:
        # Dimension Label
        dim_box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(Layout.MARGIN_LEFT), Inches(curr_y), Inches(dim_w - 0.05), Inches(row_h - 0.04)
        )
        dim_box.fill.solid()
        dim_box.fill.fore_color.rgb = Palette.BG_TINT
        dim_box.line.color.rgb = Palette.BORDER_SUBTLE
        dim_box.line.width = Pt(0.75)

        tf_d = dim_box.text_frame
        tf_d.word_wrap = True
        tf_d.margin_top = Inches(0.04)
        tf_d.margin_bottom = Inches(0.02)
        tf_d.margin_left = Inches(0.06)
        tf_d.margin_right = Inches(0.06)
        p_dt = tf_d.paragraphs[0]
        r_dt = p_dt.add_run()
        r_dt.text = f"{row['dim']}\n"
        r_dt.font.name = Typography.FONT_TITLE
        r_dt.font.size = Pt(7.5)
        r_dt.font.color.rgb = Palette.PRIMARY_DARK
        r_dt.font.bold = True

        p_dd = tf_d.add_paragraph()
        r_dd = p_dd.add_run()
        r_dd.text = row["desc"]
        r_dd.font.name = Typography.FONT_BODY
        r_dd.font.size = Pt(6.5)
        r_dd.font.color.rgb = Palette.TEXT_MUTED

        # Cells
        curr_cx = Layout.MARGIN_LEFT + dim_w
        for opt_idx, (sym, verd, detail) in enumerate(row["c"]):
            is_hl = options[opt_idx]["highlight"]
            c_box = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(curr_cx), Inches(curr_y), Inches(opt_w - 0.05), Inches(row_h - 0.04)
            )
            c_box.fill.solid()
            c_box.fill.fore_color.rgb = Palette.PRIMARY_LIGHT if is_hl else Palette.BG_CARD
            c_box.line.color.rgb = Palette.PRIMARY if is_hl else Palette.BORDER_SUBTLE
            c_box.line.width = Pt(1.0 if is_hl else 0.5)

            tf_c = c_box.text_frame
            tf_c.word_wrap = True
            tf_c.margin_top = Inches(0.04)
            tf_c.margin_bottom = Inches(0.02)
            tf_c.margin_left = Inches(0.06)
            tf_c.margin_right = Inches(0.06)

            p_cs = tf_c.paragraphs[0]
            r_sym = p_cs.add_run()
            r_sym.text = f"{sym}  {verd}\n"
            r_sym.font.name = Typography.FONT_BODY
            r_sym.font.size = Pt(7.5)
            r_sym.font.bold = True
            if sym == "●":
                r_sym.font.color.rgb = Palette.SUCCESS
            elif sym == "▲":
                r_sym.font.color.rgb = Palette.WARNING
            else:
                r_sym.font.color.rgb = Palette.DANGER

            p_cd = tf_c.add_paragraph()
            r_cd = p_cd.add_run()
            r_cd.text = detail
            r_cd.font.name = Typography.FONT_BODY
            r_cd.font.size = Pt(6.5)
            r_cd.font.color.rgb = Palette.TEXT_SECONDARY

            curr_cx += opt_w

        curr_y += row_h

    # Bottom Recommendation & Trade-off Box
    curr_y += 0.04
    rec_w = Layout.CONTENT_WIDTH * 0.60
    tra_w = Layout.CONTENT_WIDTH - rec_w - 0.12

    r_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(Layout.MARGIN_LEFT), Inches(curr_y), Inches(rec_w), Inches(0.50)
    )
    r_box.fill.solid()
    r_box.fill.fore_color.rgb = Palette.PRIMARY_LIGHT
    r_box.line.color.rgb = Palette.PRIMARY
    r_box.line.width = Pt(1.0)
    tf_r = r_box.text_frame
    tf_r.word_wrap = True
    tf_r.margin_left = Inches(0.08)
    p_r = tf_r.paragraphs[0]
    r_rt = p_r.add_run()
    r_rt.text = "★ 协同推荐结论：三位一体协同现代化架构\n"
    r_rt.font.name = Typography.FONT_BODY
    r_rt.font.size = Pt(8.0)
    r_rt.font.color.rgb = Palette.PRIMARY_DARK
    r_rt.font.bold = True
    r_rb = p_r.add_run()
    r_rb.text = "ADF 负责全域调度与源端抽取，Databricks 负责湖仓 Medallion 流批处理与 AI，Synapse 承载企业数仓建模与报表。"
    r_rb.font.name = Typography.FONT_BODY
    r_rb.font.size = Pt(7.0)
    r_rb.font.color.rgb = Palette.TEXT_PRIMARY

    t_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(Layout.MARGIN_LEFT + rec_w + 0.12), Inches(curr_y), Inches(tra_w), Inches(0.50)
    )
    t_box.fill.solid()
    t_box.fill.fore_color.rgb = Palette.BG_TINT
    t_box.line.color.rgb = Palette.BORDER_MUTED
    t_box.line.width = Pt(0.75)
    tf_t = t_box.text_frame
    tf_t.word_wrap = True
    tf_t.margin_left = Inches(0.08)
    p_t = tf_t.paragraphs[0]
    r_tt = p_t.add_run()
    r_tt.text = "⚖ 关键取舍 (Trade-off)\n"
    r_tt.font.name = Typography.FONT_BODY
    r_tt.font.size = Pt(8.0)
    r_tt.font.color.rgb = Palette.TEXT_PRIMARY
    r_tt.font.bold = True
    r_tb = p_t.add_run()
    r_tb.text = "三者结合需配置统一元数据治理 (Purview) 与网络私有端点，但换取了极佳的性能与长远可演进性。"
    r_tb.font.name = Typography.FONT_BODY
    r_tb.font.size = Pt(7.0)
    r_tb.font.color.rgb = Palette.TEXT_SECONDARY


def build_slide_5_pipeline(prs):
    """Slide 5: End-to-End Pipeline & Medallion Data Lifecycle."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(
        slide,
        tracker="Azure Data Platform Implementation | Data Lifecycle",
        action_title="数据流转闭环：从源端抽取到商业洞察的端到端 4 阶段落地工作流",
        lead_note="清晰界定 ADF 调度触发、Databricks 增量精炼与 Synapse 视图建模的职责交接面与流转机制。"
    )

    # 4 Horizontal Process Cards
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

    card_y = Layout.BODY_TOP
    card_h = 1.85
    card_gap = 0.15
    card_w = (Layout.CONTENT_WIDTH - card_gap * 3) / 4

    for i, st in enumerate(stages):
        cx = Layout.MARGIN_LEFT + i * (card_w + card_gap)
        is_hl = st["highlight"]

        s_box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(cx), Inches(card_y), Inches(card_w), Inches(card_h)
        )
        s_box.fill.solid()
        s_box.fill.fore_color.rgb = Palette.PRIMARY_LIGHT if is_hl else Palette.BG_CARD
        s_box.line.color.rgb = Palette.PRIMARY if is_hl else Palette.BORDER_SUBTLE
        s_box.line.width = Pt(1.5 if is_hl else 1.0)

        tf = s_box.text_frame
        tf.word_wrap = True
        tf.margin_top = Inches(0.08)
        tf.margin_bottom = Inches(0.06)
        tf.margin_left = Inches(0.08)
        tf.margin_right = Inches(0.08)

        # Stage Header
        p0 = tf.paragraphs[0]
        r0 = p0.add_run()
        r0.text = f"{st['num']}\n"
        r0.font.name = Typography.FONT_TITLE
        r0.font.size = Pt(8.5)
        r0.font.color.rgb = Palette.PRIMARY_DARK
        r0.font.bold = True

        p1 = tf.add_paragraph()
        r1 = p1.add_run()
        r1.text = f"{st['name']}\n"
        r1.font.name = Typography.FONT_TITLE
        r1.font.size = Pt(9.5)
        r1.font.color.rgb = Palette.TEXT_PRIMARY
        r1.font.bold = True

        p2 = tf.add_paragraph()
        p2.space_before = Pt(2.0)
        p2.space_after = Pt(4.0)
        r2 = p2.add_run()
        r2.text = st["tech"]
        r2.font.name = Typography.FONT_MONO
        r2.font.size = Pt(6.5)
        r2.font.color.rgb = Palette.PRIMARY_DARK if is_hl else Palette.TEXT_MUTED

        for b in st["bullets"]:
            pb = tf.add_paragraph()
            pb.space_before = Pt(2.0)
            rb = pb.add_run()
            rb.text = f"• {b}"
            rb.font.name = Typography.FONT_BODY
            rb.font.size = Pt(7.0)
            rb.font.color.rgb = Palette.TEXT_SECONDARY

    # Lower Quantitative Impact Cards (3 Big Stats)
    kpis_y = card_y + card_h + 0.15
    kpis_h = 1.30
    kpi_gap = 0.20
    kpi_w = (Layout.CONTENT_WIDTH - kpi_gap * 2) / 3

    kpis = [
        ("15 分钟", "数据端到端入库就绪时延", "从传统 T+1 日级批处理跨越至分钟级增量可用，全面支撑实时运营决策。"),
        ("35%", "云端计算与存储 TCO 降低", "存算分离架构 + Serverless 弹性缩容归零，杜绝常驻空转算力开销。"),
        ("70%", "手工运维排错工单压缩", "端到端自动化管道 + 统一元数据治理 (Purview)，大幅减少数据对账工作量。"),
    ]

    for j, (val, title, desc) in enumerate(kpis):
        kx = Layout.MARGIN_LEFT + j * (kpi_w + kpi_gap)
        k_box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(kx), Inches(kpis_y), Inches(kpi_w), Inches(kpis_h)
        )
        k_box.fill.solid()
        k_box.fill.fore_color.rgb = Palette.BG_TINT
        k_box.line.color.rgb = Palette.BORDER_SUBTLE
        k_box.line.width = Pt(0.75)

        tf_k = k_box.text_frame
        tf_k.word_wrap = True
        tf_k.margin_top = Inches(0.08)
        tf_k.margin_bottom = Inches(0.06)
        tf_k.margin_left = Inches(0.10)
        tf_k.margin_right = Inches(0.10)

        pk0 = tf_k.paragraphs[0]
        rk0 = pk0.add_run()
        rk0.text = f"{val}\n"
        rk0.font.name = Typography.FONT_TITLE
        rk0.font.size = Pt(20.0)
        rk0.font.color.rgb = Palette.PRIMARY
        rk0.font.bold = True

        pk1 = tf_k.add_paragraph()
        pk1.space_before = Pt(2.0)
        rk1 = pk1.add_run()
        rk1.text = f"{title}\n"
        rk1.font.name = Typography.FONT_TITLE
        rk1.font.size = Pt(9.0)
        rk1.font.color.rgb = Palette.TEXT_PRIMARY
        rk1.font.bold = True

        pk2 = tf_k.add_paragraph()
        pk2.space_before = Pt(2.0)
        rk2 = pk2.add_run()
        rk2.text = desc
        rk2.font.name = Typography.FONT_BODY
        rk2.font.size = Pt(7.0)
        rk2.font.color.rgb = Palette.TEXT_SECONDARY

    add_takeaway(
        slide,
        "流程收益结论",
        "以流批一体闭环代替散落手工脚本，打通从源端感知到业务报表毫秒呈现的全流程，兼备敏捷度、高可靠性与合规性。"
    )


def build_slide_6_roadmap(prs):
    """Slide 6: Delivery Roadmap & Implementation Timeline."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(
        slide,
        tracker="Azure Data Platform Implementation | Project Roadmap",
        action_title="实施路线图：10 周搭建高弹性、高治理能力的现代化 Azure 湖仓平台",
        lead_note="遵循敏捷迭代与规范驱动开发原则，前置架构与治理基线，分四阶段稳步实现生产级投产与价值验证。"
    )

    phases = [
        {
            "id": "PHASE 1",
            "dur": "2 周",
            "title": "基础设施与安全基线",
            "tasks": [
                "VNet 网络与私有端点 (Private Link) 规划",
                "部署 ADF、Databricks 与 Synapse 工作区",
                "Key Vault 统一凭据托管与 Entra ID 授权",
            ],
            "exit": "基础设施自动化部署通过，安全基线评审无阻断",
            "highlight": False,
        },
        {
            "id": "PHASE 2",
            "dur": "4 周",
            "title": "核心流水线与湖仓精炼",
            "tasks": [
                "ADF 首批核心源系统抽取管道调优 (小时级/天级)",
                "编写 Databricks Auto Loader 与 Bronze/Silver 作业",
                "建立数据质量校验规则库与监控告警机制",
            ],
            "exit": "首批核心系统端到端冒烟测试全部通过，吞吐达标",
            "highlight": True,
        },
        {
            "id": "PHASE 3",
            "dur": "3 周",
            "title": "数仓建模与 BI 报表集成",
            "tasks": [
                "Synapse 专用 SQL 池星型模型建立 (Fact/Dim)",
                "调优 PolyBase 高速链路，Gold层加速分发",
                "Power BI 核心管理看板搭建与 DirectLake 性能压测",
            ],
            "exit": "关键业务用户开展 UAT 验收并签署确认书",
            "highlight": False,
        },
        {
            "id": "PHASE 4",
            "dur": "1 周",
            "title": "生产割接与平稳运维交接",
            "tasks": [
                "生产环境同构部署与网络策略最终放行",
                "全量历史数据一次性对账核对与增量追平",
                "上线投产、团队运维赋能培训与平稳交接",
            ],
            "exit": "生产环境稳定运行 72 小时无重大缺陷，进入维保",
            "highlight": False,
        },
    ]

    phase_y = Layout.BODY_TOP
    phase_h = 1.80
    phase_gap = 0.15
    phase_w = (Layout.CONTENT_WIDTH - phase_gap * 3) / 4

    for i, p_info in enumerate(phases):
        px = Layout.MARGIN_LEFT + i * (phase_w + phase_gap)
        is_hl = p_info["highlight"]

        p_card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(px), Inches(phase_y), Inches(phase_w), Inches(phase_h)
        )
        p_card.fill.solid()
        p_card.fill.fore_color.rgb = Palette.PRIMARY_LIGHT if is_hl else Palette.BG_CARD
        p_card.line.color.rgb = Palette.PRIMARY if is_hl else Palette.BORDER_SUBTLE
        p_card.line.width = Pt(1.5 if is_hl else 1.0)

        tf = p_card.text_frame
        tf.word_wrap = True
        tf.margin_top = Inches(0.08)
        tf.margin_bottom = Inches(0.06)
        tf.margin_left = Inches(0.08)
        tf.margin_right = Inches(0.08)

        p0 = tf.paragraphs[0]
        r0 = p0.add_run()
        r0.text = f"{p_info['id']} ｜ {p_info['dur']}\n"
        r0.font.name = Typography.FONT_TITLE
        r0.font.size = Pt(8.5)
        r0.font.color.rgb = Palette.PRIMARY_DARK
        r0.font.bold = True

        p1 = tf.add_paragraph()
        r1 = p1.add_run()
        r1.text = f"{p_info['title']}\n"
        r1.font.name = Typography.FONT_TITLE
        r1.font.size = Pt(9.5)
        r1.font.color.rgb = Palette.TEXT_PRIMARY
        r1.font.bold = True

        for t in p_info["tasks"]:
            pt = tf.add_paragraph()
            pt.space_before = Pt(2.0)
            rt = pt.add_run()
            rt.text = f"• {t}"
            rt.font.name = Typography.FONT_BODY
            rt.font.size = Pt(7.0)
            rt.font.color.rgb = Palette.TEXT_SECONDARY

        pe = tf.add_paragraph()
        pe.space_before = Pt(4.0)
        re1 = pe.add_run()
        re1.text = "退出标准: "
        re1.font.name = Typography.FONT_BODY
        re1.font.size = Pt(6.5)
        re1.font.color.rgb = Palette.PRIMARY_DARK
        re1.font.bold = True

        re2 = pe.add_run()
        re2.text = p_info["exit"]
        re2.font.name = Typography.FONT_BODY
        re2.font.size = Pt(6.5)
        re2.font.color.rgb = Palette.TEXT_MUTED

    # Lower Half: Effort Comparison & KPI Badges
    lower_y = phase_y + phase_h + 0.15
    lower_h = 1.35
    left_w = Layout.CONTENT_WIDTH * 0.58
    right_w = Layout.CONTENT_WIDTH - left_w - 0.15
    right_x = Layout.MARGIN_LEFT + left_w + 0.15

    # Left: Comparison
    e_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(Layout.MARGIN_LEFT), Inches(lower_y), Inches(left_w), Inches(lower_h)
    )
    e_box.fill.solid()
    e_box.fill.fore_color.rgb = Palette.BG_CARD
    e_box.line.color.rgb = Palette.BORDER_SUBTLE
    e_box.line.width = Pt(0.75)

    tf_e = e_box.text_frame
    tf_e.word_wrap = True
    tf_e.margin_top = Inches(0.08)
    tf_e.margin_left = Inches(0.12)
    tf_e.margin_right = Inches(0.12)

    pe0 = tf_e.paragraphs[0]
    re0 = pe0.add_run()
    re0.text = "实施投入与交付效率对比 (Delivery Efficiency)\n"
    re0.font.name = Typography.FONT_TITLE
    re0.font.size = Pt(9.5)
    re0.font.color.rgb = Palette.PRIMARY_DARK
    re0.font.bold = True

    pe1 = tf_e.add_paragraph()
    pe1.space_before = Pt(3.0)
    re1 = pe1.add_run()
    re1.text = "传统定制开发模式： 240 人天 (各组件手工粘合，周期长达 5 个月)\n"
    re1.font.name = Typography.FONT_BODY
    re1.font.size = Pt(8.0)
    re1.font.color.rgb = Palette.TEXT_SECONDARY

    re2 = pe1.add_run()
    re2.text = "现代化协同架构模式： 145 人天 (规范驱动 + PaaS全托管，工期提效 ~40%)\n"
    re2.font.name = Typography.FONT_BODY
    re2.font.size = Pt(8.0)
    re2.font.color.rgb = Palette.SUCCESS
    re2.font.bold = True

    pe2 = tf_e.add_paragraph()
    pe2.space_before = Pt(3.0)
    re3 = pe2.add_run()
    re3.text = "💡 前置完成网络与治理基线建设，通过 Medallion 标准分层大幅避免后期返工与组件重复建设。"
    re3.font.name = Typography.FONT_BODY
    re3.font.size = Pt(7.0)
    re3.font.color.rgb = Palette.TEXT_MUTED

    # Right: 3 KPI Badges
    badges = [
        ("95 人天", "~40% 交付工期与人天净节省"),
        ("10 周", "约 2.5 个月完成全平台投产上线"),
        ("3-4 人", "精益核心交付团队 (云架构+数仓+BI)"),
    ]
    b_h = (lower_h - 0.08 * 2) / 3
    for b_idx, (b_val, b_sub) in enumerate(badges):
        by = lower_y + b_idx * (b_h + 0.08)
        b_box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(right_x), Inches(by), Inches(right_w), Inches(b_h)
        )
        b_box.fill.solid()
        b_box.fill.fore_color.rgb = Palette.PRIMARY_LIGHT if b_idx == 0 else Palette.BG_TINT
        b_box.line.color.rgb = Palette.PRIMARY if b_idx == 0 else Palette.BORDER_MUTED
        b_box.line.width = Pt(1.0 if b_idx == 0 else 0.5)

        tf_b = b_box.text_frame
        tf_b.word_wrap = True
        tf_b.margin_top = Inches(0.04)
        tf_b.margin_bottom = Inches(0.02)
        tf_b.margin_left = Inches(0.10)
        p_b = tf_b.paragraphs[0]
        p_b.alignment = PP_ALIGN.CENTER
        rb1 = p_b.add_run()
        rb1.text = f"{b_val}  "
        rb1.font.name = Typography.FONT_TITLE
        rb1.font.size = Pt(11.0)
        rb1.font.color.rgb = Palette.PRIMARY_DARK
        rb1.font.bold = True

        rb2 = p_b.add_run()
        rb2.text = b_sub
        rb2.font.name = Typography.FONT_BODY
        rb2.font.size = Pt(7.5)
        rb2.font.color.rgb = Palette.TEXT_SECONDARY

    add_takeaway(
        slide,
        "实施建议",
        "建议采取敏捷迭代交付模式：以首批高频业务系统打通全链路闭环，形成可复制标准，后续系统接入即可实现流水线式快速迁移。"
    )


# ==========================================
# 4. MAIN ORCHESTRATOR
# ==========================================
def main():
    prs = Presentation()
    prs.slide_width = Inches(Layout.WIDTH)
    prs.slide_height = Inches(Layout.HEIGHT)

    print("=" * 60)
    print("Building Bespoke Azure Architecture Presentation")
    print("=" * 60)

    print("Building Slide 1: Cover...")
    build_slide_1_cover(prs)

    print("Building Slide 2: Strategic Positioning...")
    build_slide_2_positioning(prs)

    print("Building Slide 3: Technical Architecture Blueprint...")
    build_slide_3_architecture(prs)

    print("Building Slide 4: Comparative Evaluation Matrix...")
    build_slide_4_comparison(prs)

    print("Building Slide 5: End-to-End Pipeline Workflow...")
    build_slide_5_pipeline(prs)

    print("Building Slide 6: Delivery Roadmap & Implementation...")
    build_slide_6_roadmap(prs)

    output_path = Path("Azure_Databricks_ADF_Synapse_Architecture.pptx")
    prs.save(str(output_path))
    print(f"✅ Successfully generated {len(prs.slides)} slides to {output_path}!")
    print("=" * 60)


if __name__ == "__main__":
    main()
