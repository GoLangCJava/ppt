# 自动化 PPT 脚本智能体标准化作业程序 (Automated PPT Script Agent SOP)

> **版本**：v2.0 Enterprise  
> **适用对象**：AI 编码智能体（Claude Code, Cursor, Windsurf, deer-flow, Yuxi, MyAgent 等）及自动化生成工具  
> **核心原则**：**内容驱动排版 (Content-Driven)** · **零重叠几何控制 (Zero-Overlap)** · **原生矢量输出 (Native Vector PPTX)**

---

## 1. 智能体作业准则与范式转移 (Paradigm Shift)

传统自动化生成 PPT 的最大痛点是**模板硬套 (Template Stuffing)**：无论用户输入什么内容，智能体都强行塞入固定的 3 卡片或 4 卡片布局中。这种做法会导致：
1. **内容与形式脱节**：系统架构图变成了几张干瘪的散乱文本框，无法表达上下游流转与层次。
2. **文本重叠与溢出**：不同内容长度不一，导致卡片文字与底部容器、标题严重碰撞。
3. **页面毫无设计美感**：同质化矩形让人疲劳，达不到高端咨询汇报要求。

### 范式转移：内容优先的四步闭环
```
[ 用户业务诉求 ] 
       │
       ▼
1. 内容解构与论点拆解 (One Message Per Slide & Pyramid Principle)
       │
       ▼
2. 动态选择视觉构型 (Visual Archetype: Stack, Matrix, Chevron, Hero Card)
       │
       ▼
3. 严密几何分区计算 (Zero-Overlap Dynamic Geometry Engine)
       │
       ▼
4. 原生 Python-pptx 绘制与自动化碰撞验证 (Automated Bounding-Box Check)
```

---

## 2. 幻灯片几何安全网格 (Global Geometry Engine)

所有幻灯片采用标准 **16:9 宽屏布局**（宽度 `10.0 英寸`，高度 `5.625 英寸`）。禁止随意调整宽高比。

### 垂直安全分区规范（绝对隔离带，严禁穿透）
| 分区名称 | Y 轴起始 (Inches) | 高度 (Inches) | 结束位置 (Inches) | 容纳内容与排版规则 |
| :--- | :--- | :--- | :--- | :--- |
| **Tracker 分类栏** | `0.26 in` | `0.18 in` | `0.44 in` | 全局导航分类（如 `AZURE DATA ARCHITECTURE STRATEGY`），7.5pt 大写 |
| **Action Title 主标题** | `0.46 in` | `0.36 in` | `0.82 in` | 麦肯锡结论型行动标题，14~16pt 粗体，单一完整观点 |
| **Lead Note 导读句** | `0.84 in` | `0.28 in` | `1.12 in` | 支撑观点的背景或导读，9~10pt，主品牌色或深灰 |
| **Header 分割线** | `1.16 in` | `0.75 pt` | `1.17 in` | 浅灰分割细线 (`#E2E8F0`)，贯穿内容有效宽度 |
| **Body 核心画布区** | `1.26 in` | `3.52 in` | `4.78 in` | 核心图表、分层架构、矩阵表格、流程卡片 |
| **Footer 总结栏** | `4.88 in` | `0.42 in` | `5.30 in` | 底部 Takeaway 总结卡，浅蓝底色 + 粗体标签，严禁上移 |

> **关键铁律**：  
> `Body 结束位置 (4.78 in) < Footer 起始位置 (4.88 in)`。两者之间保留 `0.10 in` 的绝对缓冲区，保证无论如何渲染绝不重叠。

---

## 3. 常见视觉构型与代码实现标准 (Visual Archetypes)

### 3.1 四层系统架构全景图 (4-Tier Cloud Architecture Blueprint)
当用户输入包含“架构图、Draw an Architecture、系统蓝图”时，必须绘制蓝图级架构视图：
- **纵向分层**：
  - `01 消费与洞察层`（Power BI 看板、AI 应用、开放 API）
  - `02 数仓与服务层`（Synapse 专用池、Serverless、Databricks SQL）
  - `03 湖仓计算层`（Medallion 架构：Bronze ➔ Silver ➔ Gold + Unity Catalog）
  - `04 接入与基建层`（ADF 管道、Event Hubs、ADLS Gen2、企业混合源）
- **层间协议指示带**：
  - 严禁乱画交叉细线！在层与层之间添加居中的等宽协议条，如：  
    `── ↕ DirectLake / DirectQuery · HTTPS 443 · REST APIs ──`
- **横向贯穿安全治理底座**：
  - 架构底部放置贯穿整页的基准条：`Microsoft Purview 统一元数据治理 · Key Vault · Entra ID · Private Link`。

### 3.2 方案选型对比矩阵 (Scenario-based Comparison Matrix)
当用户需要评估多个技术组件时：
- **表头**：左侧为场景维度列（宽 `1.70 in`），右侧为候选组件列。重点推荐组件列全列赋予冰蓝高亮背景（`#EBF5FF`）与主色边框。
- **评级符号**：采用 Harvey Balls（`● 优势显著`、`▲ 部分支持`、`○ 不适用/需自研`），杜绝大篇纯文字。
- **底部双卡决策区**：左侧 60% 宽度放置【协同推荐结论】，右侧 40% 放置【关键取舍 (Trade-off)】。

### 3.3 数据生命周期流水线 (Lifecycle Pipeline & Metric Badges)
- **上半区**：4 步横向递进卡片（Stage 01 抽取 ➔ Stage 02 清洗 ➔ Stage 03 建模 ➔ Stage 04 呈现）。
- **下半区**：3 个定量业务收益大数徽章（如 `⚡ 15 分钟`、`💰 -35%`、`📉 -70%`），大号衬线数字形成视觉焦点。

---

## 4. 文本流式排版最佳实践 (Flow Inside Containers)

在容器或卡片内部排版时，**严禁使用多个重叠的独立 TextBox**。推荐使用以下两种范式：

### 范式 A：单个 TextFrame 流式排版（首选）
```python
card = draw_rect(slide, x, y, w, h, fill_rgb=C.BG_WHITE, line_rgb=C.SOFT_BORDER)
tf = card.text_frame
tf.word_wrap = True
tf.margin_top = Inches(0.08)
tf.margin_left = tf.margin_right = Inches(0.10)

# 段落 1：小标题
p0 = tf.paragraphs[0]
r0 = p0.add_run()
r0.text = "【核心湖仓底座】\n"
r0.font.bold = True
r0.font.size = Pt(8.5)
r0.font.color.rgb = C.DATABRICKS_RED

# 段落 2：大标题
p1 = tf.add_paragraph()
r1 = p1.add_run()
r1.text = "Azure Databricks\n"
r1.font.bold = True
r1.font.size = Pt(11.0)

# 段落 3：正文描述（自动依据前序内容向下顺推，绝不重叠）
p2 = tf.add_paragraph()
p2.space_before = Pt(3.0)
r2 = p2.add_run()
r2.text = "• 统一驱动 Bronze -> Silver -> Gold 逐层精炼。"
r2.font.size = Pt(7.5)
```

### 范式 B：动态累加式 Y 轴计算（多元素容器）
```python
curr_y = card_y + padding_top
for item in items:
    draw_element(slide, x, curr_y, w, item_h)
    curr_y += item_h + element_gap
```

---

## 5. 自动化质量验证门禁 (Quality Gate Checklist)

在保存演示文稿并交付前，必须运行自动化碰撞检测脚本：
```python
from pptx import Presentation

prs = Presentation('output.pptx')
for idx, slide in enumerate(prs.slides):
    shapes = [s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip()]
    for i, s1 in enumerate(shapes):
        b1 = (s1.left, s1.top, s1.left + s1.width, s1.top + s1.height)
        for j, s2 in enumerate(shapes[i+1:], start=i+1):
            b2 = (s2.left, s2.top, s2.left + s2.width, s2.top + s2.height)
            # 矩形碰撞检测：AABB 测试
            if not (b1[2] <= b2[0] or b2[2] <= b1[0] or b1[3] <= b2[1] or b2[3] <= b1[1]):
                raise AssertionError(f"Slide {idx+1} 检测到重叠: [{i}] 与 [{j}] 发生相交！")
print("🎉 自动化门禁检测通过：全量形状零重叠！")
```

---

## 6. 异常与反模式速查表 (Anti-Patterns to Avoid)

| 反模式 (Bad Practice) | 产生原因 | 正确做法 (Best Practice) |
| :--- | :--- | :--- |
| **卡片文字盖在底部 Takeaway 上** | 卡片高度固定写死，正文文字过多向下溢出 | 将 Body 区域高度严格限制在 3.52 in 以内，字号控制在 7~8pt |
| **标题与第一行卡片紧贴** | Header 各元素间距过小，缺少分割线 | 严格遵循 `0.26 / 0.46 / 0.84 / 1.16 / 1.26` 垂直坐标梯次 |
| **架构图只画 3 个并排长方形** | 缺乏技术分层思考与通信协议表达 | 采用 4 层堆叠 (N-Tier Stack) + 左侧层号 + 层间协议带 |
| **全页纯蓝纯灰，沉闷呆板** | 缺少视觉重心地带与品牌强调色 | 使用 60-30-10 配色法则，核心推荐组件赋予浅底高亮与醒目标签 |
