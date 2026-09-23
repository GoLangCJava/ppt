---
name: consulting-ppt-generator
description: Use this skill when the user wants to generate executive, consulting-grade PowerPoint presentations (McKinsey, BCG, Bain, BTS style) using python-pptx. Produces native, 100% editable slides with crisp vector shapes, 16:9 widescreen layout, and structured consulting components (Multi-column Cards, N-Tier Architecture Blueprints, Multi-Dimensional Comparison Matrices, Delivery Roadmaps, and KPI Dashboards).
---

# Consulting PPT Generator Skill

## 1. Overview & Core Philosophy

This skill guides an AI Agent or developer to author **high-density, executive-level consulting presentations** using Python (`python-pptx`). Unlike image-based slide generators that output flat raster images with uneditable text, this skill generates **native, fully editable PowerPoint presentations (`.pptx`)**.

### Consulting Presentation Standards (Top-Tier Consulting DNA)
1. **Pyramid Principle (金字塔原理) & Conclusion First (结论先行)**:
   - Every content slide **must** have an **Action Title** (动词或判断句，直接陈述本页核心结论)，严禁使用纯名词分类作为主标题（例如：“现状分析” ❌ vs “现状分析：多类退役系统沉淀海量异构数据，亟需统一归档治理” ✅）。
   - 顶部设有 **Tracker（项目/章节导航器）** 与 **Lead Note（导读副标题）**。
2. **MECE 原则 (相互独立，完全穷尽)**:
   - 采用多列卡片 (2-5 列)、矩阵分类、分层堆栈等结构化容器承载信息，避免大段无结构文字。
3. **高信息密度与视觉层级 (Visual Hierarchy)**:
   - **字号阶梯**：封面标题 (28pt) > 幻灯片标题 (16pt) > 核心导读 (12pt) > 卡片标题 (10pt) > 正文 (8.5pt) > 标签/注脚 (6.5-7.5pt)。
   - **前缀加粗律**：项目符号列表首个逻辑短语或冒号前关键词加粗（如 `核心痛点：...`）。
4. **专业调色盘 (Enterprise Palette)**:
   - 经典咨询蓝 (`#009CDE`, `#005A9C`, `#EAF4FB`)，中性 Slate 深灰主字 (`#0F172A`)，背景微灰卡片 (`#F8FAFC`, `#FFFFFF`)，强调点缀 (`#7C3AED` AI 紫, `#10B981` 成功绿)。

---

## 2. Core Layout Archetypes (标准页面原型)

| Archetype | Component | Best For | Typical Structure |
|---|---|---|---|
| **Cover** | `cover` | 汇报封面 | 左侧竖条重音装饰 + 主标题 + 业务副标题 + 机构/作者/日期 |
| **Section** | `section` | 章节转场与目录过渡 | 居中或偏左超大章节编号 (`01`) + 章节标题 + 核心覆盖范围描述 |
| **Multi-Card** | `cards` | 现状痛点、战略支柱、能力矩阵 | 2-5 列结构卡片 (每张卡片含序号 Badge、标题、副标题、加粗要点、底部标签) + 底部结论框 |
| **KPI Metrics** | `kpis` | 现状盘点、量化收益、业务规模 | 3-4 个大数字看板 (超大字号数值 + 含义标签 + 解释说明 + 状态 Pill) |
| **Architecture** | `architecture` | 系统蓝图、技术架构、平台分层 | 3-5 层横向架构堆栈 (左侧层级标签，右侧模块网格，层间微型协议连线) |
| **Comparison** | `comparison` | 方案比选、产品评估、供应商选型 | 评级矩阵 (行: 评估维度, 列: 候选方案, 单元格: `● ▲ ○` Harvey Balls + 要点, 高亮推荐列 + 底部取舍分析) |
| **Roadmap** | `roadmap` | 实施排期、项目甘特、投入产出 | 上部阶段推进条 (阶段、周次、任务、里程碑退出标准) + 下部人天对比与 ROI 指标 |

---

## 3. Workflow (生成工作流)

When requested to create or update consulting presentations:

### Step 1: SCQA 结构澄清与内容大纲规划
1. **Situation (情境)**: 行业/企业背景、现状基线。
2. **Complication (冲突/痛点)**: 为什么现在必须做、面临的挑战、监管或业务压力。
3. **Question (关键问题)**: 核心决策点是什么？
4. **Answer (解法与落地)**: 架构方案、实施排期与收益预期。
5. 梳理幻灯片清单（通常 5-15 页完整汇报），为每一页指定 **Archetype 类型**与 **Action Title 核心论点**。

### Step 2: 编写声明式 Spec (YAML 或 JSON)
根据规划的大纲编写 `spec.yaml`，严格遵循 Pydantic 数据契约：

```yaml
title: "项目汇报方案建议书"
author: "咨询顾问团队"
date: "2026年9月"
theme: "consulting_blue"  # 可选: consulting_blue, executive_dark, strategy_navy
aspect_ratio: "16:9"

slides:
  - type: cover
    title: "企业数据归档平台建设方案"
    subtitle: "合规驱动的历史数据治理与现代化湖仓底座建议书"
    client_tag: "数字化转型顶层设计"
    dark_mode: true

  - type: cards
    tracker: "项目背景与现状分析"
    action_title: "现状分析：多类业务系统沉淀海量历史数据，面临合规与存储双重挑战"
    lead_note: "历史业务数据既是监管取证的核心凭证，也是业务经营洞察的沉淀资产。"
    takeaway: "Phase 1 建议先启动高频退役系统的数据识别与标准化归档入库。"
    cards:
      - title: "存量痛点"
        badge: "01"
        bullets:
          - "底层存储仅提供散落文件留存，缺乏统一元数据目录与索引"
          - "跨机房分散部署，查询通道陆续停运，面临外部监管审计风险"
        tags: ["存储分散", "合规风险"]
      - title: "资产价值"
        badge: "02"
        highlight: true
        bullets:
          - "完整沉淀核心交易记录与合规证据链，支撑下游经营复盘"
          - "为未来引入大模型自然语言问答与合规审计打下高质量语料基础"
        tags: ["数据资产", "AI就绪"]
      - title: "落地路径"
        badge: "03"
        bullets:
          - "分批分级推进：已具备结构系统先行，缺失元数据系统启动扫描"
          - "统一身份认证与数据脱敏，确保数据全生命周期安全闭环"
        tags: ["分期实施", "安全管控"]
```

### Step 3: 调用构建器编译生成 PPTX
在命令行直接运行：
```bash
ppt-gen -s spec.yaml -o output.pptx
```
或者在 Python 脚本中调用：
```python
from ppt.parser.builder import DeckBuilder

builder = DeckBuilder("spec.yaml")
builder.build("output.pptx")
```

### Step 4: 质量审核校验 (Consulting Quality Checklist)
生成后检查以下项目：
- [ ] **Action Title 完整性**：是否每个页面都有具体判断的陈述句？
- [ ] **信息密度**：是否留白适中，没有出现单行稀疏排版或文字溢出？
- [ ] **视觉焦点**：重点卡片或推荐方案是否设置了 `highlight: true`？
- [ ] **退出标准明确**：路线图是否附带了可验收的 Exit Criteria？
- [ ] **配色一致**：全篇主题、字体、行高、边距保持严谨一致。

---

## 4. Python Fluent API 参考 (代码编程式调用)

若需细粒度定制组件坐标或实现动态算法排版，可直接使用 Python API：

```python
from pptx import Presentation
from ppt.config.default_theme import load_theme
from ppt.core.canvas import SlideCanvas
from ppt.components.card_grid import CardItem, render_card_grid

# 1. 创建演示文稿
prs = Presentation()
theme = load_theme("consulting_blue")
prs.slide_width = prs.slide_width  # 16:9
slide = prs.slides.add_slide(prs.slide_layouts[6])

# 2. 挂载画板
canvas = SlideCanvas(slide, theme)

# 3. 绘制标准头部
canvas.add_header(
    tracker="企业数字化咨询建议书",
    action_title="战略选型：统一技术底座可有效降低 40% 的长期系统运维总成本",
    lead_note="通过标准化接口整合多源业务数据，释放端到端分析价值。",
)

# 4. 绘制卡片网格
cards = [
    CardItem(title="技术标准化", badge="01", bullets=["消除异构孤岛", "降低采购维护成本"]),
    CardItem(title="全链路可控", badge="02", bullets=["端到端数据加密", "统一权限中心"], highlight=True),
    CardItem(title="业务赋能", badge="03", bullets=["准实时报表交付", "为AI应用赋能"]),
]
render_card_grid(canvas, cards)

# 5. 添加底部结论
canvas.add_takeaway_footer("结论：优先实施核心数据层标准化，兼顾平滑过渡与长远可扩展性。")

# 6. 保存
prs.save("custom_deck.pptx")
```

---

## 5. Directory Layout

```
ppt/
├── SKILL.md                          # 本技能说明文档
├── README.md                         # 项目整体文档
├── pyproject.toml                    # Python 构建配置
├── requirements.txt                  # 依赖清单
├── ppt/                              # 核心 Python 库
│   ├── config/                       # 主题与设计规范 (consulting_blue, executive_dark, etc.)
│   ├── core/                         # 画布、栅格、几何图形、字形排版
│   ├── components/                   # 咨询级高阶原子与复合组件 (Cards, Arch, Table, Roadmap, KPI)
│   ├── templates/                    # 标准页面模板 (Cover, Section, Cards, Arch, Compare, Roadmap)
│   ├── parser/                       # 声明式 DSL (Pydantic Schema + DeckBuilder)
│   └── cli.py                        # 命令行工具 (ppt-gen)
├── examples/                         # 示例与测试规范
│   ├── sample_deck.yaml              # 完整的咨询提案 YAML 规范样例
│   ├── generate_demo.py              # 一键生成示例脚本
│   └── output/                       # 输出生成的 .pptx 文件
└── tests/                            # 自动化单元测试
```
