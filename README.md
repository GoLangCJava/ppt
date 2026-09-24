# Consulting PPT Generator (基于 Python 的咨询级 PPT 自动化生成引擎)

<div align="center">

**高质量 · 原生矢量 · 顶级咨询机构 (McKinsey / BCG / Bain / BTS) 风格 PPT 自动化生成引擎与 Agent Skill**

[![Python 3.9+](https://img.shields.io/badge/Python-3.9%2B-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Framework: python-pptx](https://img.shields.io/badge/Engine-python--pptx-orange.svg)](https://python-pptx.readthedocs.io/)
[![Standard: Agent Skill](https://img.shields.io/badge/Format-Agent%20Skill-blueviolet.svg)](SKILL.md)

</div>

---

## 📖 项目简介

本项目专为**企业级咨询汇报、数字化转型方案、技术架构设计与项目实施交付提案**打造。

与市面上常见的「大模型生成整张不可编辑位图图片再贴入 PPT」的方案不同，本项目通过 Python (`python-pptx`) 配合严苛的**咨询排版设计系统**，直接生成 **100% 原生矢量、文本可编辑、高信息密度、具备严谨视觉层级**的顶级咨询风格幻灯片。

项目内包含了：
1. **Agent Skill 规范 (`SKILL.md`)**：遵循 Claude Code / Deer-flow / Yuxi / MyAgent 标准，任何智能体均可阅读并直接执行该技能。
2. **声明式 DSL 解析器 (YAML / JSON)**：仅需编写结构化数据或让大模型输出 YAML，即可自动编译出完整 PPT。
3. **编程式 Fluent API**：支持在 Python 代码中按需灵活拼装自定义组件。
4. **黄金基准参考样例**：以仓库中附带的 48 页顶级咨询机构方案《`China RIMS Solution Proposal_V1.0_20260902.pptx`》为设计美学标杆，复刻其设计系统。

---

## 🎨 咨询级设计系统 (Consulting Design DNA)

- **金字塔原理与 Action Title 体系**：
  - 每页幻灯片强制配备 **Action Title（结论先行句）**，直接陈述本页核心洞察，辅以 **Lead Note（导读副标题）** 与 **Tracker（顶部导航器）**。
- **经典咨询色盘 (Enterprise Palette)**：
  - 咨询主蓝（`#009CDE`、`#005A9C`）、高对比 Slate 深灰文字（`#0F172A`）、浅色微卡片（`#F8FAFC`、`#FFFFFF`）、高亮强调色（`#7C3AED` AI 紫、`#10B981` 成功绿）。
- **高密度标准原型组件 (Components)**：
  - **Cover**：高阶管理层汇报封面（支持深色与浅色模式）。
  - **Section**：章节大编号转场过渡页。
  - **Cards Grid**：2-5 列结构化分析卡片（含编号徽章、加粗列表、微标签、高亮突出）。
  - **Architecture View**：N-Tier 多层系统架构蓝图（含层级标签、服务网格、协议连线）。
  - **Comparison Matrix**：多维方案比选与评估矩阵（含 Harvey balls `● ▲ ○` 评级球、推荐列高亮、底部关键取舍卡）。
  - **Roadmap & Gantt**：交付实施路线图（阶段推进、周次排期、任务列表、退出标准 Exit Criteria、投入人天对比与 ROI 看板）。
  - **KPI Metrics**：大数字量化看板。
  - **Executive Takeaway**：底部核心判断与结论横幅。

---

## 🚀 快速上手

### 1. 环境安装

```bash
git clone https://github.com/GoLangCJava/ppt.git
cd ppt

# 安装依赖
pip install -r requirements.txt

# 以可编辑模式安装 CLI 工具
pip install -e .
```

### 2. 命令行一键生成

直接使用项目内置的完整咨询案例文档 (`examples/sample_deck.yaml`)：

```bash
# 查看所有支持的主题
ppt-gen --list-themes

# 从 YAML 规范生成 PPTX
ppt-gen -s examples/sample_deck.yaml -o examples/output/consulting_demo.pptx
```

生成的 PPTX 文件位于 `examples/output/consulting_demo.pptx`，可直接用 PowerPoint / WPS 打开演示与二次编辑。

### 3. 运行 Python 演示脚本

```bash
python3 examples/generate_demo.py
```

### 4. 运行单元测试

```bash
pytest -v
```

---

## 📝 声明式编写示例 (YAML DSL)

你或任何大模型只需编写清晰的 YAML，例如：

```yaml
title: "数字化平台建设方案建议书"
author: "咨询顾问团队"
date: "2026年9月"
theme: "consulting_blue"

slides:
  # 封面
  - type: cover
    title: "China RIMS Solution Proposal"
    subtitle: "企业历史数据标准化归档与智能化治理技术方案建议书"
    dark_mode: true

  # 3 列卡片现状分析
  - type: cards
    tracker: "项目背景与现状分析"
    action_title: "现状分析：多类退役系统沉淀海量异构数据，亟需统一归档治理"
    lead_note: "历史业务数据是企业核心资产，也是满足监管审计、保障业务连续的关键依据。"
    takeaway: "Phase 1 建议先启动高频退役系统的数据识别与标准化归档入库。"
    cards:
      - title: "存量痛点"
        badge: "01"
        bullets:
          - "现有本地仅散落存储，缺乏全局元数据目录与索引"
          - "历史通道停运，面临外部监管审计与调阅合规风险"
        tags: ["架构分散", "合规风险"]
      - title: "资产价值"
        badge: "02"
        highlight: true
        bullets:
          - "完整沉淀业务历史记录，支撑下游经营复盘与策略优化"
          - "为未来大模型智能问答与合规挖掘打下高质量语料底座"
        tags: ["数据资产", "AI就绪"]
      - title: "实施路径"
        badge: "03"
        bullets:
          - "分批启动：具备结构系统先行，缺失元数据系统启动逆向扫描"
          - "全链路加密与动态脱敏，确保数据全生命周期合规"
        tags: ["分期实施", "安全管控"]
```

---

## 🛠️ Python 代码编程式调用 (Programmatic API)

```python
from pptx import Presentation
from ppt.config.default_theme import load_theme
from ppt.core.canvas import SlideCanvas
from ppt.components.card_grid import CardItem, render_card_grid

prs = Presentation()
theme = load_theme("consulting_blue")
slide = prs.slides.add_slide(prs.slide_layouts[6])

canvas = SlideCanvas(slide, theme)
canvas.add_header(
    tracker="企业数字化咨询建议书",
    action_title="战略选型：统一技术底座可有效降低 40% 的长期系统运维总成本",
    lead_note="通过标准化接口整合多源业务数据，释放端到端分析价值。",
)

cards = [
    CardItem(title="技术标准化", badge="01", bullets=["消除异构孤岛", "降低采购维护成本"]),
    CardItem(title="全链路可控", badge="02", bullets=["端到端数据加密", "统一权限中心"], highlight=True),
    CardItem(title="业务赋能", badge="03", bullets=["准实时报表交付", "为AI应用赋能"]),
]
render_card_grid(canvas, cards)
canvas.add_takeaway_footer("结论：优先实施核心数据层标准化，兼顾平滑过渡与长远可扩展性。")

prs.save("custom_output.pptx")
```

---

## 📁 目录结构与规范文档库

### 📚 核心设计规范与方法论文档 (Methodology & Standards)
- **`how-to-design-better-consulting-ppt.md`**：【核心推荐】**如何将咨询与技术 PPT 设计得更好看**——从“死板土味方框”到“顶级咨询视觉艺术”进阶指南（包含 9 大立竿见影美化秘诀与 Python 实战代码）。
- **`consulting-ppt-visual-chart-standards.md`**：咨询级 PPT 视觉图表与版式规范标准（60-30-10 配色法则、字号字重梯度、微圆角与边框标准、Harvey Balls、大数收益徽章）。
- **`mckinsey-structured-communication-guide.md`**：麦肯锡结构化沟通与金字塔原理指南（BLUF 原则、SCQA 叙事模型、MECE 分类法则、Action Title 公式）。
- **`automated-ppt-script-agent-sop.md`**：自动化 PPT 脚本智能体标准化作业程序（16:9 宽屏绝对隔离带、单 TextFrame 流式排版、自动化零碰撞检测门禁）。
- **`agent.md`**：AI Agent 咨询级 PPT 生成 SOP 与操作手册。

```
ppt/
├── SKILL.md                          # 供智能体 (Claude Code / Agent) 调用的标准化 Skill
├── README.md                         # 项目完整文档
├── agent.md                          # 智能体生成 SOP 与最佳实践
├── how-to-design-better-consulting-ppt.md # 视觉美化进阶指南 (如何设计得更好看)
├── consulting-ppt-visual-chart-standards.md # 咨询图表视觉规范
├── mckinsey-structured-communication-guide.md # 麦肯锡结构化沟通指南
├── automated-ppt-script-agent-sop.md # 自动化脚本智能体 SOP
├── generate_azure_architecture_deck.py # 专属定制 Azure 架构 PPT 生成脚本
├── Azure_Databricks_ADF_Synapse_Architecture.pptx # 原生矢量高颜值演示文稿成品
├── China RIMS Solution Proposal_V1.0_20260902.pptx # 48页黄金参考提案 (参考基准)
├── ppt/                              # Python 核心包
│   ├── config/                       # 主题色彩系统 (咨询蓝, 行政暗色, 经典海军蓝)
│   ├── core/                         # 栅格排版、画布管理器、矢量几何与字形渲染
│   ├── components/                   # 咨询级高阶组件 (Cards, Arch, Table, Roadmap, KPI, Takeaway)
│   ├── templates/                    # 完整页面模板 (Cover, Section, Cards, Arch, Compare, Roadmap)
│   ├── parser/                       # 声明式 DSL (Pydantic Schema + DeckBuilder)
│   └── cli.py                        # 命令行交互工具 (ppt-gen)
├── examples/                         # 示例演示
│   ├── sample_deck.yaml              # 完整的咨询提案 YAML 规范样例
│   ├── generate_demo.py              # 一键生成示例脚本
│   └── output/                       # 生成文件输出目录
└── tests/                            # 自动化单元测试
```

---

## 📄 许可证

[MIT License](LICENSE)
