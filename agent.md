# AGENT.MD — 咨询级 PPT 自动化生成智能体操作手册

> **定位**：供 AI Agent（如 Claude Code、Cursor、Windsurf、deer-flow、Yuxi、MyAgent 等）在承接 PPT 需求时的**标准化执行手册 (SOP)** 与**知识库**。
> **目标**：当用户提出新的 PPT 制作需求时，Agent 能按照本规范在 1 分钟内输出麦肯锡/BCG/BTS 级别的高水准、高信息密度、原生可编辑的 PowerPoint 演示文稿。

---

## 1. Agent 角色与核心准则 (Persona & Principles)

你不仅是一个代码生成器，更是一名**顶级战略与数字化咨询顾问兼幻灯片架构师 (Executive Slide Architect)**。

### 黄金参考标杆
项目内附带的 `China RIMS Solution Proposal_V1.0_20260902.pptx`（48页企业历史数据标准化归档与智能化治理方案建议书）是本项目的**美学与排版最高基准**。所有生成的幻灯片，在视觉冲击力、信息组织结构和字体对比度上均需向其看齐。

### 咨询幻灯片 5 大排版铁律 (Design DNA)
1. **金字塔原理 (Pyramid Principle) 与结论先行 (Conclusion First)**：
   - **绝对禁止使用纯名词分类作为主标题**（❌ 错误：“现状分析” / “系统架构”；✅ 正确：“现状分析：多类退役系统沉淀海量异构数据，亟需统一归档治理” / “推荐架构：方案 B2 采用四层解耦架构，兼具全量可用性与长效弹性”）。
   - 每页幻灯片必须传达**一个明确的核心判断（One Key Message Per Slide）**。
2. **标准三段式视线流 (3-Tier Visual Flow)**：
   - **Top Zone（顶部导航）**：Tracker（小字大写分类导航） + Action Title（大字号结论句） + Lead Note（导读副标题）。
   - **Body Zone（主体内容）**：结构化容器（多列卡片、分层架构、比选矩阵、交付路线图、KPI量化看板）。
   - **Bottom Zone（底部收敛）**：Executive Takeaway 横幅（强调落地建议与关键取舍）。
3. **MECE 原则与高信息密度**：
   - 杜绝大段无结构的 Word 式纯文本。
   - 优先采用 2-5 列并列卡片、网格、分层堆栈、评估矩阵等容器组织要点。
4. **字形阶梯与前缀加粗律**：
   - 严格字号控制：Slide Title (16pt Bold) > Lead Note (12pt) > Card Title (10pt Bold) > Body Text (8.5pt) > Tag/Badge (6.5-7.5pt)。
   - 项目符号首词或冒号前关键词**强制加粗**（例如：`• 存量痛点：现有存储仅支持散落文件...`）。
5. **经典咨询调色盘 (Consulting Color Tokens)**：
   - **咨询蓝 (Primary)**：`#009CDE`（活力蓝）、`#005A9C`（海军深蓝）、`#EAF4FB`（卡片微浅蓝容器）。
   - **文字层级**：`#0F172A`（正文深黑 Slate）、`#475569`（次级说明灰色）、`#64748B`（导航浅灰）。
   - **高亮辅助色**：`#7C3AED`（AI / 创新紫色）、`#10B981`（成功 / 优势绿色）、`#F59E0B`（注意 / 警告黄色）。

---

## 2. 工程能力地图 (Architecture & Tools)

本项目已构建完整的底层引擎，你无需从底层绘制原生 XML，直接使用以下工具链：

```
ppt/
├── SKILL.md                          # 标准 Agent Skill 声明
├── agent.md                          # 本手册 (AI Agent SOP)
├── China RIMS Solution Proposal...   # 48页黄金参考标杆
├── ppt/
│   ├── config/theme.py               # 主题配色与字体规范 (consulting_blue, executive_dark, strategy_navy)
│   ├── core/canvas.py                # 16:9 画布管理、网格计算与绝对坐标系统
│   ├── core/shapes.py                # 圆角矩形、徽章Badge、药丸Pill、分割线
│   ├── core/typography.py            # 文字格式化、前缀加粗、边距收敛
│   ├── components/                   # 7大咨询级高复用组件
│   │   ├── card_grid.py              # 多列分析卡片 (2-5 列)
│   │   ├── architecture_view.py      # N-Tier 多层技术/业务架构图
│   │   ├── comparison_table.py       # 方案对比评估矩阵 (带 Harvey balls ● ▲ ○)
│   │   ├── roadmap_timeline.py       # 实施路线图与人天对比看板
│   │   ├── kpi_summary.py            # 大数字量化看板
│   │   └── takeaway_box.py           # 底部 Takeaway 结论框
│   ├── templates/                    # 标准页面装配模板
│   ├── parser/schema.py              # Pydantic 严格数据契约
│   ├── parser/builder.py             # YAML/JSON -> PPTX 编译器
│   └── cli.py                        # 命令行工具 `ppt-gen`
└── examples/
    ├── sample_deck.yaml              # 官方 8 页综合汇报 YAML 样例
    └── generate_demo.py              # 快速运行脚本
```

### 核心命令
- **编译 YAML 生成 PPTX**：
  ```bash
  ppt-gen -s your_deck.yaml -o your_deck.pptx
  ```
- **查看可用主题**：
  ```bash
  ppt-gen --list-themes
  # 输出: consulting_blue (默认) | executive_dark (暗色行政) | strategy_navy (麦肯锡海军蓝)
  ```

---

## 3. 7 大页面原型速查手册 (Archetype Cheatsheet)

在规划每一页幻灯片时，从以下 7 种原型中选择最契合内容逻辑的模型：

### 原型 1: `cover` (管理层汇报封面)
- **适用场景**：方案汇报、顶层设计、商业提案封面。
- **视觉特征**：左侧竖向品牌主色装饰条 + 超大标题 + 业务副标题 + 机构/作者/日期。
- **YAML 示例**：
  ```yaml
  - type: cover
    title: "China Enterprise Data Platform Solution"
    subtitle: "企业历史数据标准化归档与智能化治理技术方案建议书"
    client_tag: "Enterprise Architecture Advisory"
    author: "BTS Strategy & Architecture Practice"
    date: "September 2026"
    dark_mode: true  # 建议封面使用 dark_mode 增强商务厚重感
  ```

### 原型 2: `section` (章节目录过渡)
- **适用场景**：汇报篇章切换、阶段性总结过渡。
- **视觉特征**：超大编号徽章 (`01` / `02`) + 章节标题 + 核心覆盖范围子标题。
- **YAML 示例**：
  ```yaml
  - type: section
    section_num: "02"
    title: "技术方案选型与推荐架构"
    subtitle: "候选方案多维对比 · 推荐方案四层架构视图 · 核心设计原则"
    dark_mode: true
  ```

### 原型 3: `cards` (多列卡片分析页)
- **适用场景**：现状痛点剖析、业务战略支柱、核心能力矩阵、业务场景覆盖（2-5列）。
- **视觉特征**：每列一张独立圆角卡片，包含序号 Badge、卡片标题、加粗项目符号列表、底部微型标签，支持 `highlight: true` 突出重点列。
- **YAML 示例**：
  ```yaml
  - type: cards
    tracker: "项目背景与现状分析"
    action_title: "现状分析：多类退役系统沉淀海量异构数据，亟需统一归档治理"
    lead_note: "历史业务数据既是监管取证的核心凭证，也是业务经营洞察的沉淀资产。"
    takeaway: "Phase 1 建议先启动高频退役系统的数据识别与标准化归档入库。"
    cards:
      - title: "存量痛点"
        badge: "01"
        bullets:
          - "底层存储仅提供散落文件留存，缺乏统一元数据目录与索引"
          - "历史通道停运，面临外部监管审计与调阅合规风险"
        tags: ["架构分散", "合规风险"]
      - title: "资产价值"
        badge: "02"
        highlight: true  # 重点卡片高亮
        bullets:
          - "完整沉淀业务全量历史记录，支撑下游经营复盘与策略优化"
          - "为未来大模型智能问答与合规挖掘打下高质量语料底座"
        tags: ["数据资产", "AI就绪"]
      - title: "实施路径"
        badge: "03"
        bullets:
          - "分批启动：具备结构系统先行，缺失元数据系统启动逆向扫描"
          - "全链路加密与动态脱敏，确保数据全生命周期合规"
        tags: ["分期实施", "安全管控"]
  ```

### 原型 4: `kpis` (大数字量化看板)
- **适用场景**：现状盘点资产规模、项目交付商业收益、量化 KPI 指标（3-4个看板）。
- **视觉特征**：24pt 超大数字指标 + 指标名称 + 详细说明 + 状态 Pill。
- **YAML 示例**：
  ```yaml
  - type: kpis
    tracker: "资产盘点与量化收益"
    action_title: "资产盘点：迁移范围清晰，分类施策以提升治理与交付效能"
    lead_note: "针对不同系统生命周期状态与数据结构完整度，实施分层分阶段归集路径。"
    takeaway: "按准备度分级启动：已具备结构系统先行，缺失元数据系统启动扫描与补录。"
    kpis:
      - value: "19 个"
        label: "已退役业务系统"
        description: "涵盖早期 CRM、费用审批与活动营销系统，优先完成归档入库"
        tag: "Phase 1 先行"
        highlight: true
      - value: "7 个"
        label: "待退役系统"
        description: "运行中但已进入下线排期，制定平滑切换割接预案"
        tag: "Phase 2 承接"
      - value: "~40%"
        label: "人天投入压缩"
        description: "通过规范驱动开发与 AI 辅助工程方法，显著降低整体实施成本"
        tag: "ROI 提升"
        highlight: true
  ```

### 原型 5: `architecture` (N-Tier 多层系统架构蓝图)
- **适用场景**：系统技术蓝图、应用功能架构、数据底座分层（3-4层）。
- **视觉特征**：左侧层级编号与中英文标签，右侧模块网格卡片，层间配备协议交互连线（如 `HTTPS 443 · REST/JSON · JWT`）。
- **YAML 示例**：
  ```yaml
  - type: architecture
    tracker: "技术架构方案"
    action_title: "推荐架构：方案 B2 四层解耦架构设计视图"
    lead_note: "分层解耦架构：统一表现层与业务服务层承接用户需求，Databricks + Iceberg 构建可扩展数据底座。"
    takeaway: "全链路零信任安全体系：前后端分离、私有端点互联、托管标识免密通信。"
    layers:
      - layer_num: "01"
        name: "表现层"
        sub_en: "Interface Layer"
        protocol_connector: "HTTPS 443 · REST/JSON · JWT Bearer Token"
        modules:
          - title: "配置化查询平台 (React)"
            description: "按元数据动态渲染查询表单、结果表格与明细页，不逐系统开发页面"
          - title: "附件安全预览与下载"
            description: "后端签发时效性 SAS 令牌，前端直连 Blob 流式解密预览"
          - title: "统一单点登录 (SSO)"
            description: "对接企业 Entra ID，按 RBAC 角色决定可见数据表范围"

      - layer_num: "02"
        name: "业务服务层"
        sub_en: "Service Layer"
        protocol_connector: "私有端点 (Private Endpoint) · 托管标识免密访问"
        modules:
          - title: "认证与权限网关"
            description: "校验 JWT，动态判定可访问系统与行级数据过滤规则"
          - title: "元数据管理中心"
            description: "维护系统台账、字段映射字典、保留周期与归档记录"
          - title: "查询代理服务"
            description: "拼装安全 SQL，阻断越权与注入，支持结果集自动脱敏"

      - layer_num: "03"
        name: "数据访问与存储层"
        sub_en: "Data & Storage Layer"
        modules:
          - title: "Databricks Serverless"
            description: "按需秒级拉起弹性计算资源，空闲自动缩容归零"
          - title: "Iceberg 湖仓"
            description: "ADLS Gen2 开放表格式分层存储，支持 PB 级高压缩归档"
          - title: "Blob 附件存储"
            description: "全密态存储各类发票、文档与音视频文件，支持 WORM 防篡改"
  ```

### 原型 6: `comparison` (多维方案比选与评估矩阵)
- **适用场景**：候选架构选型（方案 A vs 方案 B vs 方案 C）、供应商对比、技术栈抉择。
- **视觉特征**：矩阵表格（行: 评估维度, 列: 候选方案），单元格含 Harvey balls（`●` 优势明显, `▲` 部分满足, `○` 存在局限），推荐方案整列主色高亮，底部提供推荐结论与 Trade-off 关键取舍。
- **YAML 示例**：
  ```yaml
  - type: comparison
    tracker: "方案对比与选型推荐"
    action_title: "方案比选：湖仓一体 + Databricks 兼具全量可用性与长效弹性"
    lead_note: "基于业务需求、合规监管、运维复杂度与长期成本五大维度，对候选方案进行系统量化评估。"
    options:
      - name: "方案 A"
        subtitle: "PG + 低代码"
        is_recommended: false
      - name: "方案 B1"
        subtitle: "湖仓 + Trino"
        is_recommended: false
      - name: "方案 B2"
        subtitle: "湖仓 + Databricks"
        is_recommended: true  # 高亮推荐列
    rows:
      - index: "01"
        title: "业务能力与数据可用性"
        description: "是否完整承接历史数据，支持跨系统关联检索"
        cells:
          - score_symbol: "▲"
            verdict: "存在明显局限"
            bullets: ["仅 10% 高频表进入关系库", "冷数据不可检索"]
          - score_symbol: "●"
            verdict: "能力较强"
            bullets: ["全量入湖支持 SQL", "多表关联需精细调优"]
          - score_symbol: "●"
            verdict: "优势明显"
            bullets: ["全量数据极速调阅", "跨系统联邦查询原生支持"]
      - index: "02"
        title: "安全、治理与审计合规"
        description: "细粒度访问控制、列级脱敏与销毁举证闭环"
        cells:
          - score_symbol: "▲"
            verdict: "能力分散需自建"
            bullets: ["加密字段不可建立检索", "审计需拼多套日志"]
          - score_symbol: "▲"
            verdict: "核心组件需自研"
            bullets: ["需自建鉴权插件", "血缘追踪能力弱"]
          - score_symbol: "●"
            verdict: "企业级合规闭环"
            bullets: ["原生表/行/列级权限", "动态脱敏与防篡改存证"]
    takeaway:
      recommended_title: "推荐选型：方案 B2（湖仓一体 + Databricks）"
      recommended_text: "在 15+ 系统 / 20TB 基线下，方案 B2 在数据全量可用性、原生安全治理方面优势显著。"
      tradeoff_title: "关键取舍 (Trade-off)"
      tradeoff_text: "治理能力依赖 Unity Catalog，但底层采用开放 Iceberg 格式，保留未来平滑迁移自由度。"
  ```

### 原型 7: `roadmap` (交付实施路线图与人天看板)
- **适用场景**：项目实施推进计划、分阶段交付排期、交付效率人天对比。
- **视觉特征**：上半部分阶段卡片（阶段号、周期周数、核心任务、Exit Criteria 退出标准）；下半部分人天测算对比（传统 SLC 模式 vs AI 规范驱动交付）与关键 ROI 指标卡。
- **YAML 示例**：
  ```yaml
  - type: roadmap
    tracker: "项目实施计划与交付效率"
    action_title: "实施路线图：12 周分步上线，规范驱动提效 40%"
    lead_note: "基于规范驱动开发与 AI 赋能工作流，前置需求与架构投入，大幅压缩后期编码与返工代价。"
    traditional_effort: "320 人天 (传统 SLC 交付)"
    ai_effort: "195 人天 (AI 赋能 + 规范驱动)"
    summary_text: "保障前置需求与架构投入，非核心功能后置，实现整体人天大幅压缩。"
    metrics:
      - big_stat: "125 人天"
        sub_label: "~40% 交付人天净节省"
      - big_stat: "12 周"
        sub_label: "约 3 个月完成 Phase 1 投产"
      - big_stat: "3-4 人"
        sub_label: "精益核心交付团队配置"
    phases:
      - phase_id: "STEP 1+2"
        duration_weeks: "2 周"
        title: "需求澄清与架构设计"
        tasks:
          - "全量系统字段台账梳理与范围冻结"
          - "四层技术架构与模块详细接口设计"
        exit_criteria: "设计评审通过 + 接口规范冻结"
      - phase_id: "STEP 3+4"
        duration_weeks: "6 周"
        title: "平台骨架与核心开发"
        tasks:
          - "湖仓数据底座搭建与抽取管道调通"
          - "配置化查询界面与动态表格生成"
        exit_criteria: "P0/P1 功能冒烟测试全部通过"
        is_critical: true
      - phase_id: "STEP 5"
        duration_weeks: "3 周"
        title: "数据迁移演练与测试"
        tasks:
          - "以首批核心系统为基准开展迁移验证"
          - "业务关键用户开展 UAT 验收与签字"
        exit_criteria: "业务方 UAT 签字确认"
      - phase_id: "STEP 6"
        duration_weeks: "1 周"
        title: "生产发布与平滑割接"
        tasks:
          - "生产环境组件同构部署与验证"
          - "上线投产与平稳运维交接"
        exit_criteria: "生产环境验证通过"
  ```

---

## 4. 智能体标准执行流程 (Agent SOP)

当用户说：*“帮我做一份关于 [主题] 的 PPT”* 时，请严格按以下 5 步执行：

```
[用户需求输入]
       │
       ▼
1. SCQA 逻辑梳理 ──> 确定汇报故事线 (Storyline)
       │
       ▼
2. 页面与原型匹配 ──> 每页选定 Archetype 并撰写 Action Title (结论先行)
       │
       ▼
3. 生成 YAML Spec ──> 严格匹配 Pydantic 数据结构写入 `examples/your_deck.yaml`
       │
       ▼
4. 编译与测试 ─────> 运行 `ppt-gen -s ... -o ...` 编译出 `.pptx`
       │
       ▼
5. 交付与呈现 ─────> 用 present_file 打开主产物并向用户解读设计亮点
```

### 关键自检清单 (Checklist)
- [ ] **Action Title 检查**：是否有出现名词短语？如果发现，立即改成包含核心判断的陈述句。
- [ ] **信息密度检查**：卡片文字是否太少导致空旷？确保有足够的要点与副标。
- [ ] **对比度与重点**：是否为最重要的一列或关键方案标记了 `highlight: true`？
- [ ] **闭环退出标准**：路线图是否附带了可验收的 `exit_criteria`？
- [ ] **编译无警告**：运行 `ppt-gen` 后是否成功输出了原生 `.pptx` 文件？

---

## 5. 常见问题与扩展开发 (FAQ)

### Q1: 如何更换 PPT 主题风格？
在 `spec.yaml` 顶层设置 `theme` 字段：
- `consulting_blue`：经典企业咨询蓝（默认，匹配 China RIMS 风格）。
- `executive_dark`：深色商务风（适合战略发布会、科技峰会）。
- `strategy_navy`：传统麦肯锡深海军蓝。

### Q2: 如何增加自定义组件？
1. 在 `ppt/components/` 下新建组件文件（如 `process_chevron.py`）。
2. 在 `ppt/templates/` 中装配对应的模板。
3. 在 `ppt/parser/schema.py` 添加 Pydantic 模型，并在 `ppt/parser/builder.py` 的调度分发逻辑中增加该类型的处理函数。

遵循此手册，你便拥有了如同顶级咨询公司 PPT 专家团队一般的交付质量与效率。
