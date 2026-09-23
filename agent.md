# AGENT.MD — 咨询级 PPT 自动化生成智能体操作手册 (SOP)

> **定位**：供 AI Agent（如 Claude Code、Cursor、Windsurf、deer-flow、Yuxi、MyAgent 等）在承接咨询汇报类 PPT 需求时的**标准化作业程序 (SOP)** 与**知识库**。
> **核心原则**：**内容驱动排版 (Content-Driven Layout)**，拒绝生搬硬套死板模板；**零重叠排版 (Zero-Overlap Geometry)**，严谨计算坐标与字号阶梯；**原生矢量输出 (Native Vector PPTX)**，保持 100% 可二次编辑。

---

## 1. 核心思想：内容驱动，拒绝死板硬套 (Content-Driven vs Rigid Templates)

在生成咨询 PPT 时，**最常见且严重的错误就是“把所有内容硬塞进同一种卡片模板”**，导致：
- 架构图变成了几张干瘪的卡片，缺乏上下游连通关系与层级感；
- 文字过多时与卡片边框、底部标签发生灾难性的**重叠 (Overlap)**；
- 页面同质化严重，缺乏顶级咨询方案的视觉说服力。

### 正确的流程：四步内容优先法 (Content-First Process)
1. **明确每页的独立信息使命 (One Message Per Slide)**：
   - 架构方案汇报不是堆砌术语，而是解决客户选型痛点。例如本案（Azure Databricks & ADF & Synapse）：
     - 第 1 页：**封面**（立调性、定基调）。
     - 第 2 页：**战略定位与分工**（解决“为什么需要三个工具，各自边界在哪”）。
     - 第 3 页：**端到端技术蓝图 (Draw an Architecture)**（画出真实的 4 层系统架构蓝图与交互链路）。
     - 第 4 页：**多维场景选型比选**（通过矩阵表格定量对比 5 个核心场景的优劣）。
     - 第 5 页：**数据生命周期流水线**（横向流转，展现从源端抽取到 BI 展现的 4 阶段）。
     - 第 6 页：**实施排期与工期对比**（敏捷推进计划与人天节省看板）。
2. **依据内容形态动态选择视觉构型 (Visual Archetype Selection)**：
   - 系统蓝图 ➔ **多层容器堆栈 (N-Tier Stack) + 协议连线 + 侧边/底部安全边框**。
   - 方案选型 ➔ **多维评估矩阵表格 (Matrix Table) + Harvey Balls 评级球 + 推荐列高亮**。
   - 职责分工 ➔ **3 列对比卡片 (Columns) + 适用/不适用边界**。
   - 数据流水线 ➔ **横向流动步骤条 (Chevron/Process Flow)**。
3. **针对性编写 Python 绘图代码**：
   - 对特定主题定制专属的 Python 生成脚本，精确计算每个组件的坐标、边距和尺寸，确保每一个像素与字号都经过计算。

## 2. 视觉创意与高级感设计法则 (Visual Creativity & Aesthetics Rules)

高端咨询与云架构提案绝不是“枯燥的单色框图”，必须具备现代极简与极具表现力的**视觉创意 (Visual Creativity)**：

### 2.1 品牌色与多层次色彩系统 (Rich Thematic Color Palette)
- **主调与暗底冲击**：
  - 封面页采用 Deep Slate Navy (`#0A192F`)，配合 Azure Blue (`#0078D4`) 形成强烈科技感对比与尊享感。
- **产品原生色标体系 (Branded Accents)**：
  - **Azure Data Factory**：深紫/蓝靛 (`#5C2D91`)，浅底 (`#F5F0FF`)，传递“管道与编排纽带”特性。
  - **Azure Databricks**：烈焰橙红 (`#EA3824`)，浅底 (`#FEF2F2`)，凸显“Medallion 湖仓计算核心”地位。
  - **Azure Synapse**：蓝绿/青色 (`#008272`)，浅底 (`#F0FDFA`)，展现“企业级数仓与报表分析”稳健。
- **Medallion 湖仓金银铜质感阶梯**：
  - **Bronze Layer**：古铜琥珀 (`#B45309`)
  - **Silver Layer**：质感冷银 (`#475569`)
  - **Gold Layer**：黄金琥珀 (`#D97706`)

### 2.2 视觉容器结构创新 (Structural Creative Components)
1. **封面双分视觉架构卡 (Split Composition Hero Card)**：
   - 左侧 60% 主标题排版与咨询元数据，右侧 35% 绘制深色磨砂质感的“三驾马车蓝图缩略卡”，在封面第一眼就传达架构全貌。
2. **彩色顶条与标签徽章 (Top Accent Bars & Tag Pills)**：
   - 每张核心卡片顶部增加 0.05 英寸的产品专属色彩条，内嵌 `【👑 核心湖仓底座】`、`【全域调度接入纽带】` 等视觉胶囊标签。
3. **真实协议连接带 (Protocol Strips)**：
   - 层与层之间不再留白或画容易错位的交叉细线，而是使用居中的等宽协议条，例如 `── ↕ ABFS (ADLS Gen2) · Delta Lake 开放格式 · Apache Spark 3.x ──`，既具技术真实感又规范美观。
4. **定量收益大数卡 (Metric Badges)**：
   - 提取业务高光指标（如 `⚡ 15 分钟`、`💰 -35%`、`📉 -70%`），大号衬线数字配合紧凑解说，瞬间抓住高层管理者的注意力。

---

## 3. 零重叠排版几何法则 (Zero-Overlap Geometry Rules)

幻灯片重叠的根本原因在于**硬编码了绝对坐标而没有做垂直/水平累加**。必须严格遵循以下几何分区与坐标系：

### 16:9 页面全局分区 (10.0 × 5.625 英寸)
- **左边距 / 右边距**：`margin_left = 0.55 in`, `margin_right = 0.55 in`，有效内容宽度 `content_width = 8.90 in`。
- **顶部 Header 安全区 (0.28 in ~ 1.25 in)**：
  ```
  y = 0.28 ~ 0.46 in:  Tracker 导航器 (18-20pt 间距，7.5pt 字体，大写次级灰)
  y = 0.50 ~ 0.88 in:  Action Title 核心论点 (16pt Bold 结论句，深黑文字)
  y = 0.92 ~ 1.20 in:  Lead Note 导读副标题 (11-12pt 咨询蓝文字)
  y = 1.25 in:         细分割线 (Divider Line, 0.75pt, #E2E8F0)
  ```
  **严格校验**：`0.46 < 0.50`，`0.88 < 0.92`，`1.20 < 1.25`，保证 Header 内部绝无重叠！
- **中部 Body 安全区 (1.35 in ~ 4.80 in)**：
  - 有效高度：`3.45 in`。
  - 所有图表、分层架构、卡片网格均限定在此区间内部。
- **底部 Footer 安全区 (4.90 in ~ 5.35 in)**：
  - `footer_top = 4.90 in`, `height = 0.42 in`。
  - 放置 Executive Takeaway 核心总结横幅或关键取舍 (Trade-off)。
  - `4.80 < 4.90`，保证 Body 与 Footer 绝不重叠！

### 卡片内部的流式计算准则 (Flow Inside Cards)
在任何卡片或容器内部绘制标题、副标题、项目符号时，**严禁使用独立重合的 TextBox**！
- 方式一：**单个 TextFrame 承载**（最佳）。将卡片标题、副标题、正文段落全部写入同一个 `TextFrame` 中，通过 `space_before` 和 `space_after` 控制间距，PowerPoint 底层引擎会自动流式排版，天然杜绝文字互相重叠。
- 方式二：**累加式 Y 轴计算**。如果必须分形状绘制，必须：`curr_y = prev_y + prev_height + gap`，并动态预估文本折行后的实际高度。

---

## 4. 架构图专业绘制指南 (How to "Draw an Architecture" in Python)

当用户提出“画一个架构图”时，一个专业的咨询级架构幻灯片必须包含以下 5 大要素（参考 `China RIMS Solution Proposal` Slide 27 & 37）：

1. **分层容器 (Layer Containers)**：
   - 纵向划分 3-4 个层级（例如：消费层 ➔ 服务与数仓层 ➔ 湖仓精炼层 ➔ 接入与源端层）。
   - 左侧为**层级说明卡片**（编号 `01`、中文名、英文名）；右侧为该层内部的**服务组件卡片**。
2. **协议与流动连线 (Protocol & Flow Strips)**：
   - 层与层之间必须有细长的高亮连线或协议条，清晰标注接口与协议（如 `HTTPS 443 · DirectLake / REST APIs`、`ABFS · Delta Lake 格式 · Apache Spark 3.x`）。
3. **核心计算与存储差异化表现**：
   - 重点核心组件（如 Databricks Delta Lake、Synapse Dedicated Pool）使用**品牌色描边**与**微浅色背景**高亮突出。
4. **贯穿全局的治理与安全基线 (Cross-Cutting Foundation)**：
   - 在底部或侧边设置贯穿所有层级的安全基础设施条（如 `Microsoft Purview 元数据治理 · Azure Key Vault 密钥 · Entra ID 单点登录 · 全私网互通`）。
5. **底部 Takeaway 总结卡**：
   - 总结本架构的 3-4 个核心技术设计原则（存算分离、计算弹性归零、单一事实源 SSOT、零信任访问）。

---

## 5. 选型对比矩阵绘制指南 (How to Draw a Comparison Matrix)

当用户需要评估对比技术栈（如 ADF vs Databricks vs Synapse）时：
1. **维度明确**：横轴为候选方案，纵轴为业务核心维度（数据接入、清洗转换、湖仓一体、并发报表、AI高级分析）。
2. **Harvey Balls 权威评级**：
   - `●` (绿色/主色)：能力极强，天然原生支持。
   - `▲` (黄色/橙色)：部分支持，但有边界或需自建较多逻辑。
   - `○` (浅灰/红色)：不适用或非专长。
3. **重点列整列高亮**：为推荐的组合方案或核心推荐组件添加浅蓝背景底色与深蓝边框。
4. **底部必须包含“推荐结论”与“关键取舍 (Trade-off)”**。

---

## 6. 智能体重新生成 PPT 的具体执行步骤 (Action SOP)

遇到 PPT 生成或重构需求时，请按顺序执行：
1. **规划内容与页面结构**：在脑海或草稿中列出各页的大纲与核心论点。
2. **编写专属 Python 脚本**：创建针对该主题的生成脚本（例如 `generate_azure_architecture_deck.py`），导入 `python-pptx` 与设计系统主题。
3. **执行并验证**：运行 Python 脚本生成 `.pptx`，并通过检测脚本确认所有形状之间**零碰撞、零重叠**。
4. **提交与呈现**：提交至 Git 工作分支，并通过 `present_file` 打开生成的 PPT 或规范文件呈现给用户。
