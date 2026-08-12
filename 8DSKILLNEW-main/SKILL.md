---
name: 8d-skill
description: "汽车行业IATF 16949 8D问题解决报告生成器。根据产品名和缺陷描述智能匹配5套行业模板（涂装/装配/焊接/尺寸/通用），自动生成单 Sheet .xlsx和标准.docx格式8D报告。覆盖D0-D8全流程：5Why完整预填路径、6M鱼骨图排查（含判定）、RC根因总结、CA纠正措施、D7 Yokoten横向展开。纯本地生成。Use when user needs 8D report for automotive quality issues, customer complaints, SCAR, or root cause analysis."
---

# 8D Skill — 汽车行业 8D 问题解决报告生成器

> **版本**：2.0.3
> **适用对象**：汽车零部件质量工程师、SQE、CQE、QE
> **输出物**：8D 报告 .xlsx（单 Sheet）+ 8D 报告 .docx（标准文档）
> **依赖**：openpyxl、python-docx（纯本地生成，无任何外部 API）

---

## 一、概述

8D（Eight Disciplines）问题解决法是汽车行业 IATF 16949 体系下处理客户投诉、供应商质量问题的标准方法。一份规范的 8D 报告既是质量改善工具，也是与客户/供应商沟通的正式技术文件。

本 Skill 的价值：

1. **降低报告撰写门槛**——质量工程师只需提供产品名、缺陷描述、客户等基础信息，Skill 自动补全 5Why 路径、6M 方向、遏制措施等结构化内容
2. **保证报告规范性**——所有 8 个章节均按 IATF 16949 与 VDA 框架预填模板，避免漏项
3. **同时输出 Excel 和 Word**——Excel 适合工程内部流转与追溯，Word 适合正式提交客户
4. **可按缺陷类型匹配模板**——内置 5 套行业模板（涂装/装配/焊接/尺寸/通用），覆盖汽车零部件 90% 以上常见缺陷
5. **融入真实行业参考资料**——模板 5Why 路径和 6M 分析方向基于 Quality-One 8D 方法论、AIAG CQI 标准和汽车行业真实案例设计

---

## 二、何时触发

满足以下任一条件即应触发本 Skill：

| 触发条件 | 示例用户输入 |
|---|---|
| 明确提到"8D 报告" | "客户投诉外观不良，需要 8D 报告" |
| 提到客户投诉 + 产品 + 缺陷 | "我是做保险杠的，XX 客户投诉漆面颗粒，要 8D" |
| 提到质量追溯/根因分析需求 | "请帮我做一份根因分析报告，客户要 8D 格式" |
| 提供了产品+缺陷，要求 Excel/Word 报告 | "生成一份注塑件尺寸超差的 8D 报告 Excel 和 Word" |

**反例（不应触发本 Skill）**：

- "讲解一下 8D 是什么" → 走知识问答，不生成文件
- "我们公司最近质量很差怎么办" → 缺少具体产品/缺陷信息，应先 ask 收集
- "生成一份 PPT 给客户讲 8D" → 应使用 pptx skill，不是本 Skill

---

## 三、工作流

主 Agent 接到 8D 报告生成请求后，按以下 5 步执行：

### Step 1：从用户输入提取关键信息

至少需要提取以下字段，缺失的字段必须用 `AskUserQuestion` 工具向用户追问：

| 字段 | 必填 | 示例 | 缺失时追问示例 |
|---|---|---|---|
| product（产品名） | ✅ | 前保险杠总成（涂装件） | "请告诉我具体的产品名称是什么？" |
| defect（缺陷描述） | ✅ | 漆面颗粒/杂质 | "请描述具体的缺陷现象（如颗粒、流挂、超差多少 mm）" |
| customer（客户名） | ✅ | XX 汽车有限公司 | "客户是哪家主机厂？" |
| defect_rate（不良率） | 推荐 | 500 PPM（0.05%） | "目前的不良率/不良数量是多少？（请用 PPM 或 ≤0.5% 的百分比表示）" |
| batch_size（8D 分析样本数） | 推荐 | 12 件 | "本次客户投诉/退货的本批次样本数量是多少件？" |
| output_dir（输出目录） | 可选 | ~/Desktop | 默认 ~/Desktop |

> ⚠️ **行业常识基准（填写示例值时严禁违反，详见第十章）**：
> - **不良率**：汽车零部件量产线 IATF 16949 目标 ≤50 PPM；客户投诉触发 8D 的典型量级是 100–2000 PPM（0.01%–0.2%）；>0.5%（5000 PPM）属于停线停发事故，不应作为「常规 8D 示例」。
> - **批次数量（8D 分析样本数）**：指本次客户投诉/退货的具体样本件数，典型 1–30 件（线束、ECU 等高价值件常为个位数）；**不是**生产批量、**不是**2000 件这类数字。
> - 向用户追问时给出的「示例值」必须落在上述区间内，**禁止**用 5%、8%、11.5%、2000、500 这类脱离常识的数字作为提示。

### Step 2：匹配最合适的模板

参考 `templates/INDEX.md`，按下表规则匹配模板：

| 用户描述关键字 | 选用模板 slug |
|---|---|
| 漆面、涂装、颗粒、流挂、色差、橘皮、缩孔 | `paint-defect` |
| 装配、间隙、面差、卡扣、异响、松动 | `assembly-defect` |
| 焊接、虚焊、焊穿、焊渣、焊点、强度 | `welding-defect` |
| 尺寸、超差、CPK、公差、变形、收缩 | `dimensional-defect` |
| 其他/无法明确分类 | `generic-defect` |

匹配规则：先按"缺陷描述"匹配，再按"产品类别"复核。两者冲突时优先按缺陷描述匹配。

### Step 3：模板融合逻辑说明（generate_8d_report_tool 自动完成，Agent 无需手动操作）

> ⚠️ **重要边界声明**：本节描述的是 `generate_8d.py` 脚本内部的模板融合逻辑，**仅供 Agent 理解原理用**。
> Agent **不应该**自己去读取 `template.json` 或手动做占位符替换——这些工作全部由 `generate_8d_report_tool` 调用脚本自动完成。
> Agent 唯一要做的就是：在 Step 4 调用 `generate_8d_report_tool(product, defect, customer, defect_rate, batch_size, template)` 并传入 6 个参数。
>
> ❌ **错误做法**：Agent 用 `python` 代码读取 `template.json` → 自己做 `replace({defect_type}, ...)` → 用 `export_xlsx_tool` 拼表格
> ✅ **正确做法**：Agent 直接调用 `generate_8d_report_tool`，脚本自动完成模板加载 + 占位符替换 + Excel/Word 生成

脚本 `generate_8d.py` 内部会做以下处理（Agent 不需要关心）：

1. **加载模板**：根据 `--template` 参数读取 `templates/<slug>/template.json`
2. **替换占位符**：将模板中所有 `{defect_type}`、`{product_name}`、`{customer}` 替换为用户实际输入
3. **填充 5W2H**：将 product、defect、customer、defect_rate、batch_size 填入 D0-D2 信息表
4. **生成 8D 编号**：格式 `8D-{YYYYMMDD}-{HHMMSS}`，例如 `8D-20260621-093015`
5. **PPM ↔ % 互转**：自动识别 `500PPM` / `0.05%` / `500` 等格式并双格式展示

> **知识增强提示（可选）**：如果用户描述的缺陷类型在模板 `defect_types` 中匹配度较低，或涉及特殊工艺/材料，Agent 可使用 `web_search_tool` 搜索以下关键词获取真实参考资料，**用于在对话中向用户解释根因方向**（不要修改 template.json，脚本生成时会用预填内容）：
> - `"{defect_type} root cause analysis automotive"`
> - `"{product_type} manufacturing defects 8D"`
> - `"IATF 16949 8D problem solving"`
> - `"AIAG CQI process control"`

> 🔴 **本节仅适用于 8D 报告场景**：本节描述的模板融合逻辑、`generate_8d_report_tool` 工具、`templates/` 目录下的 5 套模板，**全部仅用于 8D 报告生成**。
> 当用户需要生成 **DFMEA / PFMEA / FMEA / 控制计划 / CP / 工艺卡** 等其他类型文档时，**不应该**调用 `generate_8d_report_tool`，而应该用 `export_xlsx_tool`（Excel）或 `export_document_tool`（Word）按用户需求自己组织内容。
> 判断依据：用户明确提到「8D」「客户投诉+产品+缺陷」「SCAR」「根因分析报告」→ 8D → `generate_8d_report_tool`；其他场景（DFMEA/PFMEA/控制计划等）→ 用 `export_xlsx_tool` / `export_document_tool`。

### Step 3.5：动态 5Why 覆盖（可选，提升报告专业度）

> 💡 **背景**：模板预填的 5Why 路径是**通用版**，对具体缺陷的针对性不足。当 Agent 已基于用户描述的缺陷现象、工艺背景、初步线索推演出**更具体的 5Why 路径**时，可以通过 `generate_8d_report_tool` 的 `five_why_steps` 参数传入，覆盖模板预填内容。

**何时传入动态 5Why**：

| 场景 | 是否传入 | 原因 |
|------|---------|------|
| 用户给了产品+缺陷+工艺背景+初步线索 | ✅ 传入 | Agent 能推演出具体的 5Why，比模板预填更有价值 |
| 用户描述的缺陷不在任何模板的 `defect_types` 覆盖范围内 | ✅ 传入 | 模板预填的 5Why 不匹配，需要 Agent 自己分析 |
| 用户只给了产品+缺陷，无其他信息 | ❌ 不传入 | Agent 没有足够信息推演，用模板预填即可 |
| 用户明确说"你帮我分析根因" | ✅ 传入 | 用户期望 Agent 做深度分析 |

**5Why JSON 格式**（必须 6 步，answer 要具体）：

```json
[
  {"level": "问题", "question": "产品{product_name}为什么出现{defect_type}？", "answer": "——", "evidence": "——"},
  {"level": "Why 1", "question": "为什么出现{defect_type}？", "answer": "直接原因（如：铸造缩松导致局部壁厚不足）", "evidence": "X光探伤显示缩松区"},
  {"level": "Why 2", "question": "为什么{Why1的answer}？", "answer": "...", "evidence": "..."},
  {"level": "Why 3", "question": "为什么{Why2的answer}？", "answer": "...", "evidence": "..."},
  {"level": "Why 4", "question": "为什么{Why3的answer}？", "answer": "...", "evidence": "..."},
  {"level": "Why 5（根因）", "question": "为什么管理/系统层面存在漏洞？", "answer": "系统原因（如：新设备导入流程未将关键参数稳定性作为验收强制项）", "evidence": "..."}
]
```

**5Why 推演规则**：
1. 每一步的 question 必须基于上一步的 answer 继续追问（不能跳层）
2. 每一步的 answer 必须具体（不要用"请填写"），最好带可验证的证据
3. Why 5 必须定位到**管理/系统层面**的根因（如流程缺失、规范缺陷、体系漏洞），不能停留在设备/工艺层面
4. 推演深度参考 `references/5why_examples.md` 中的 5 个真实案例

### Step 4：调用 generate_8d_report_tool 生成报告（唯一入口）

> 🔴 **唯一正确入口**：Agent 必须通过调用工具 `generate_8d_report_tool(...)` 来生成 8D 报告，**不要**用 `subprocess` / `os.system` / `python -c` 等方式直接执行脚本。
>
> 工具内部会自动调用 `generate_8d.py` 脚本，传递参数，解析输出 JSON，返回下载链接给 Agent。

**调用方式**（Agent 在工具调用时这样传参）：

```
# 基础调用（用模板预填 5Why，空白处留 ____）
generate_8d_report_tool(
    product="前保险杠总成（涂装件）",
    defect="漆面颗粒/杂质",
    customer="XX 汽车有限公司",
    defect_rate="500PPM",
    batch_size="12",
    template="paint-defect"
)

# 自动填充模式（用户说"你帮我填"/"给我示例"时启用）
generate_8d_report_tool(
    product="前保险杠总成（涂装件）",
    defect="漆面颗粒/杂质",
    customer="XX 汽车有限公司",
    defect_rate="500PPM",
    batch_size="12",
    template="paint-defect",
    auto_fill=True
)

# 进阶调用（传入动态 5Why，覆盖模板预填）
generate_8d_report_tool(
    product="前保险杠总成（涂装件）",
    defect="漆面颗粒/杂质",
    customer="XX 汽车有限公司",
    defect_rate="500PPM",
    batch_size="12",
    template="paint-defect",
    five_why_steps='[{"level":"问题","question":"...","answer":"——","evidence":"——"},{"level":"Why 1","question":"...","answer":"...","evidence":"..."},...,{"level":"Why 5（根因）","question":"...","answer":"...","evidence":"..."}]'
)
```

工具内部自动完成：

- 调用 `python skills/8d-skill/scripts/generate_8d.py` 脚本（脚本路径自动定位，无需 Agent 关心）
- 检查并尝试自动 pip install openpyxl、python-docx
- 加载对应模板 `template.json`，替换占位符，填充 5W2H
- 生成 `8D报告_{product}_{defect}_模板.xlsx`（单 Sheet，带合并单元格+章节标题+根因高亮）
- 生成 `8D报告_{product}_{defect}.docx`（标准 8D Word 文档）
- 返回包含两个文件下载链接的字符串，Agent 在回复中原样展示给用户

### Step 5：向用户展示结果（先输出 D0-D8 完整内容，再展示下载链接）

**输出顺序（必须严格遵守）**：

1. **先在对话中输出 D0-D8 完整内容**（Markdown 格式）：
   - D0 准备 / D1 团队 / D2 问题描述（5W2H）/ D3 临时遏制措施
   - D4 根本原因分析（5Why 路径 + 6M 排查 + RC1/RC2/RC3 总结）
   - D5 永久纠正措施 / D6 实施计划 / D7 横向展开 / D8 关闭
   - 每个步骤都要包含预填的 5Why 路径、6M 排查表、CA 措施等关键信息
   - 用户能立即看到分析结果，不用等文件下载

2. **然后调用 `generate_8d_report_tool` 生成 xlsx + docx 文件**

3. **最后展示下载链接**：
   - xlsx 文件下载链接 + 一句话说明（"单 Sheet 完整 Excel，可继续编辑"）
   - docx 文件下载链接 + 一句话说明（"标准 Word 文档，可直接提交客户"）
   - 已匹配的模板 slug 和 5Why 路径摘要（≤3 句）
   - 提示用户："报告已生成，对话中的内容与文件一致，如需补充实际数据可直接编辑 xlsx 后提交客户"

**为什么要在对话里输出完整内容**：
- 用户能立即看到分析结果，不用等文件下载
- 对话内容与文件内容一致，方便用户对照
- 如果文件下载失败，用户至少能在对话里看到完整报告
- 体现 Agent 的分析能力（D4 5Why 推演、6M 排查等）

### 🔧 自动填充模式（auto_fill 参数）

> 💡 **背景**：用户明确说"你帮我填"、"给我个示例"、"看一下范例"时，8D 报告应该填充合理的示例值，而不是留 `____` 空白。

**Agent 判断逻辑**：
- ✅ **用户明确说**"你帮我填"、"帮我示例一下"、"看一下范例"、"其他不要问我"→ 启用 `auto_fill=True`
- ❌ **用户没明说**（只提供产品/缺陷/客户）→ 默认 `auto_fill=False`，保留 `____` 空白让用户填

**启用 auto_fill 后，脚本会自动填充以下内容**（不需要 Agent 传入）：

| 字段 | 自动填充规则 | 示例 |
|------|-------------|------|
| D1 团队姓名 | 按角色生成化名（张伟/李娜/王芳/刘强/陈静/赵磊） | 团队领导：张伟（质量工程师） |
| D1 联系方式 | 内部分机号 | 8001 / 8002 / 8003... |
| D3 责任人 | 按措施类型分配角色 | 库存全检→质量部（张伟）；在途拦截→物流部（周敏） |
| D3 完成时间 | 当前日期 + 2/3/5/7 天 | 2026-06-26 / 2026-06-28 |
| D5 责任人 | 按措施分配 | RC1→工艺部（李娜）；RC2→质量部（张伟） |
| D5/D6 完成时间 | 当前日期 + 7/14/30 天 | 2026-07-01 / 2026-07-15 |
| D7 责任人/完成时间 | 同 D5 | |
| D8 编制/审核/批准 | 化名 + 当天日期 | 编制：张伟 2026-06-24；审核：王芳 2026-06-24 |
| 8D 负责人 | 团队领导化名 | 张伟 |
| 报告发起人 | 团队领导化名 | 张伟 |

**注意**：
- 化名是固定的（张伟/李娜/王芳/刘强/陈静/赵磊/周敏/孙健），方便用户区分
- 日期基于当前日期计算，不用真实生产日期
- 这些都是**示例值**，用户拿到 xlsx 后可以替换成真实数据

**注意事项：**
- ✅ 对话中的 D0-D8 内容应与生成的 xlsx/docx 文件内容保持一致
- ✅ 如果用户提供了根因线索，D4 5Why 应基于线索推演（可同时通过 `five_why_steps` 参数传入脚本覆盖模板预填）
- ✅ 启用 `auto_fill=True` 时，对话中的内容也应反映填充后的示例值
- ⚠️ 对话内容用 Markdown 表格格式，文件内容用 Excel/Word 原生表格，两者数据一致但呈现形式不同

---

## 四、8D 八个步骤概述

| 步骤 | 名称 | 核心目的 | 本 Skill 输出位置 |
|---|---|---|---|
| D0 | 准备 | 立项、初步评估严重度 | Excel 单 Sheet / Word 第一章 |
| D1 | 团队组建 | 跨职能小组成立 | Excel 单 Sheet / Word 第一章 |
| D2 | 问题描述 | 5W2H 描述问题 | Excel 单 Sheet / Word 第一章 |
| D3 | 临时遏制措施 | 隔离不良品、保护客户 | Excel 单 Sheet / Word 第二章 |
| D4 | 根本原因分析 | 5Why（完整预填路径）+ 6M 排查（含判定）+ RC 总结 + 验证 | Excel 单 Sheet / Word 第三章 |
| D5 | 永久纠正措施制定 | CA 方案评估 | Excel 单 Sheet / Word 第四章 |
| D6 | 永久纠正措施实施 | 实施跟踪表 | Excel 单 Sheet / Word 第四章 |
| D7 | 预防再发生 | Yokoten 横向展开 | Excel 单 Sheet / Word 第五章 |
| D8 | 关闭与团队致谢 | 签名关闭 | Excel 单 Sheet / Word 第六章 |

---

## 五、模板选择指南

详见 `templates/INDEX.md`。简要选择流程：

```
用户描述缺陷
    ↓
是漆面/外观问题？→ paint-defect
    ↓ 否
是装配/间隙问题？→ assembly-defect
    ↓ 否
是焊接/连接问题？→ welding-defect
    ↓ 否
是尺寸/公差问题？→ dimensional-defect
    ↓ 否
generic-defect（兜底）
```

每个模板下都有 `intro.md` 说明适用场景，`template.json` 提供预填内容。

### template.json 数据结构（v2.0+）

每个 `template.json` 包含：

- `slug` / `name`：模板标识
- `defect_types` / `product_categories`：覆盖的缺陷类型和产品类别
- `d0_d2`：D0-D2 阶段提示（缺陷等级/发现位置/影响）
- `d3_template.containment_actions`：5 项预填遏制措施
- `d4_template.5why_path`：**完整 5Why 路径**（`steps` 数组，Why 1→5 连贯递进，含预填答案和证据）
- `d4_template.6m_analysis`：**6M 排查表**（数组格式，每项含 `m`/`finding`/`judgment`）
- `d4_template.root_cause_summary`：**预填根因总结**（RC1/RC2 含描述和类型）
- `d4_template.verification`：**预填验证结论**（含具体数据和判定）
- `d5_d6_template.permanent_actions`：预填 CA 方案（含责任人建议）
- `d7_template.yokoten`：预填横向展开措施

> 从 v2.0 起，5Why 从 hints（候选列表）升级为完整预填路径（steps），6M 增加 judgment 判定字段，D4 增加 root_cause_summary 和 verification。模板生成时会优先使用新格式，旧格式自动兼容。

### 行业标准参考（真实出处，撰写 D4/D5/D6 时可引用）

| 标准 | 出处 | 在 8D 中的引用场景 |
|---|---|---|
| **IATF 16949:2016** | 国际汽车工作组（IATF）| D2/D3/D4 条款对应（详见 `references/8d_guide.md` 末尾关联表） |
| **AIAG-VDA FMEA 1st Ed. (2019)** | AIAG + VDA 联合发布 | D4 根因分析后回填 PFMEA；D7 更新 PFMEA 的 RPN/AP |
| **AIAG PPAP 4th Ed.** | AIAG | D6 实施验证后向客户提交 PSW；D7 重大变更需重新 PPAP |
| **AIAG APQP 2nd Ed.** | AIAG | D7 预防措施纳入下一轮 APQP 阶段门评审 |
| **AIAG MSA 4th Ed.** | AIAG | D4 测量系统分析（GRR% ≤ 10% 接受，10%–30% 视情况）|
| **AIAG SPC 2nd Ed.** | AIAG | D4 关键参数 SPC 监控；D6 验证 Cpk ≥ 1.33 量产、≥ 1.67 安全件 |
| **VDA 6.3 (Process Audit)** | VDA-QMC | D7 横向展开过程审核；P6/P7 元素复查 |
| **Ford 8D (1986)** | 福特汽车，8D 方法论起源 | D0-D8 步骤框架的本源 |
| **Toyota 5Why / Yokoten** | 丰田生产方式 TPS | D4 5Why、D7 Yokoten 横向展开 |
| **Quality-One 8D** | Quality-One International | 8D 培训参考资料，5Why 范例常用来源 |
| **AIAG CQI-9 (热处理)/CQI-11 (电镀)/CQI-12 (涂装)/CQI-15 (焊接)/CQI-23 (成型)** | AIAG 特殊过程评估 | 对应模板（焊接→CQI-15，涂装→CQI-12，注塑→CQI-23）的 D4 根因方向与 D5 控制项依据 |

> 引用真实标准时必须保留发布机构与版本号。Agent 不应编造标准号或版本；如不确定，宁可不引用。

---

## 六、references/ 目录说明

| 文件 | 用途 | 何时查阅 |
|---|---|---|
| `8d_guide.md` | 8D 方法论详细指南，D0-D8 每步的关键活动、输出物、常见错误 | 撰写报告时对某一步骤不确定，或客户提出方法论质疑时 |
| `5why_examples.md` | 5 个行业/缺陷类型的完整 5Why 范例 + 常见断点提示 | 模板预填的 5Why 路径不适用，需要参考其他案例时 |
| `fishbone_guide.md` | 6M 详细说明（每个 M 含 5+ 条排查项）+ 汽车行业排查清单 | 需要细化鱼骨图分析方向时 |

---

## 七、完整调用示例

**用户输入**：

> 我是一家汽车外饰件厂，生产的后保险杠（注塑件）装配到车身后，客户反馈两侧间隙不一致，间隙超差约 2mm，本批次客户退货 8 件，按 4000 件出货量折算约 200 PPM，需要 8D 报告。

**Agent 处理**：

1. 提取信息：product="后保险杠（注塑件）"、defect="两侧间隙不一致，超差约 2mm"、customer="（缺失，需追问）"、defect_rate="200 PPM"、batch_size="8 件"
2. 追问客户名和批次数量
3. 匹配模板：缺陷描述包含"间隙超差"→ 候选 `assembly-defect` 和 `dimensional-defect`，产品类别为注塑件 → 最终选 `dimensional-defect`
4. 调用脚本生成文件
5. 返回文件路径

**预期输出**：

- `/home/z/my-project/download/8D报告_后保险杠（注塑件）_两侧间隙不一致超差约2mm_模板.xlsx`
- `/home/z/my-project/download/8D报告_后保险杠（注塑件）_两侧间隙不一致超差约2mm.docx`
- 5Why 路径应指向"注塑收缩率/模具定位/检具校准"
- 6M 方向应包含 Man（操作工装模经验）、Machine（注塑机参数稳定性）、Material（原料批次变异）、Method（保压曲线 SOP）、Measurement（检具校准）、Environment（车间温度）

---

## 八、关键约束

1. **纯本地生成**：不依赖任何外部 API 或付费服务
2. **中文为主**：所有文字用中文，技术术语保留英文（如 5Why、PFMEA、CPK）
3. **占位符替换**：`{defect_type}`、`{product_name}` 等必须替换为用户实际输入
4. **默认留空 + 按需自动填充**：默认情况下所有用户需填写的空白处用 `____` 标记；但当用户明确要求"你帮我填"、"给我示例"时，Agent 应启用 `auto_fill=True` 参数让脚本自动填充合理示例值（化名/示例日期/角色分配）
5. **信息不足时主动追问**：用 `AskUserQuestion` 工具引导用户补充
6. **不截断**：宁可文字精炼，也不省略号截断
7. **对话输出 D0-D8 完整内容**：必须在对话中输出 D0-D8 的完整内容（含 5Why 路径、6M 排查表、CA 措施等），让用户立即看到分析结果。然后再调用 `generate_8d_report_tool` 生成 xlsx/docx 文件并展示下载链接。对话内容与文件内容应保持一致。
8. **行业常识优先**：在向用户追问或在模板中填入示例值时，不良率与批次数量必须落在第十章「行业常识基准」规定的区间内，违反视为严重 bug。

---

## 九、文件清单

```
8d-report/
├── SKILL.md                              （本文档）
├── VERSION                               （版本号：2.0.3）
├── scripts/
│   └── generate_8d.py                    （核心生成脚本）
├── references/
│   ├── 8d_guide.md                       （8D 方法论详细指南）
│   ├── 5why_examples.md                  （5Why 范例库）
│   └── fishbone_guide.md                 （鱼骨图 6M 指南）
└── templates/
    ├── INDEX.md                          （模板索引）
    ├── paint-defect/                     （涂装缺陷模板）
    │   ├── template.json
    │   └── intro.md
    ├── assembly-defect/                  （装配缺陷模板）
    │   ├── template.json
    │   └── intro.md
    ├── welding-defect/                   （焊接缺陷模板）
    │   ├── template.json
    │   └── intro.md
    ├── dimensional-defect/               （尺寸超差模板）
    │   ├── template.json
    │   └── intro.md
    └── generic-defect/                   （通用兜底模板）
        ├── template.json
        └── intro.md
```

---

## 十、行业常识基准（🔴 严禁违反）

> 本章节是 8D Skill 的**硬约束**。当 Agent 向用户追问 `defect_rate` / `batch_size`、或在模板示例值/对话提示中引用任何数字时，必须落在以下区间内。违反将导致用户失去对报告专业性的信任，视为严重 bug。

### 10.1 不良率（defect_rate）

| 量级 | PPM | 百分比 | 适用场景 |
|---|---|---|---|
| 世界级标杆 | ≤ 25 PPM | ≤ 0.0025% | IATF 16949 优秀供应商日常水平 |
| 量产正常水平 | 25–100 PPM | 0.0025%–0.01% | 正常 SPC 受控状态 |
| 客户投诉触发 8D | 100–2000 PPM | 0.01%–0.2% | **本 Skill 的典型场景** |
| 严重质量事故 | 2000–5000 PPM | 0.2%–0.5% | 需启动 SCAR + 厂内停线 |
| 停发/召回级 | > 5000 PPM | > 0.5% | 必须高管介入、可能召回 |

**示例值使用规则**：
- 默认示例值用 **500 PPM (0.05%)**，这是最常见的客户投诉触发量级
- 安全件（焊接、制动、转向）示例值用 **50 PPM (0.005%)**，安全件门槛更低
- **严禁**用 3%、5%、8%、11.5% 这类「灾难级」数字作为常规示例——它们对应 30000/50000/80000/115000 PPM，已属于停发召回事故
- 用户实际提供的不良率若 >0.5%，Agent 应在 D0 阶段标注「严重度 = 高，建议升级处理」，但仍按用户实际数据生成报告
- 表达方式优先使用 PPM；若用户用百分比，换算后同时展示两种（如 0.05% / 500 PPM）

### 10.2 批次数量（batch_size）

> ⚠️ 本 Skill 的 `batch_size` 字段语义是**「本次 8D 分析的具体样本件数」**，即客户投诉/退货的具体件数，**不是生产批量、不是出货总量**。如需记录生产批量，请在 D2 问题描述中单独说明。

| 产品类型 | 典型 batch_size | 说明 |
|---|---|---|
| 线束、ECU、传感器 | 1–10 件 | 高价值件，客户投诉多为个位数 |
| 保险杠、仪表板、门板总成 | 3–20 件 | 中价值件，客户 Audit 退货常为个位数到十几件 |
| 紧固件、卡扣、冲压件 | 5–50 件 | 低价值大批量件，退货样本相对较多 |
| 涂装件颗粒投诉 | 5–30 件 | Audit 扣分对应样本 |

**示例值使用规则**：
- 默认示例值用 **12 件**（覆盖大多数场景）
- 线束类默认 **5 件**
- **严禁**用 500、2000、5000 这类数字作为 `batch_size` 示例——这混淆了「8D 分析样本」与「生产批量」
- 如用户混淆了概念（如回答「2000 件」），Agent 应在追问中澄清：「您说的 2000 件是本批次生产总量，还是客户本次投诉/退货的具体件数？8D 报告的 batch_size 字段记录的是后者」

### 10.3 AskUserQuestion 追问时的合规示例

✅ 合规示例：
```
不良率（推荐）：当前不良比例是多少？
  例：500 PPM、0.05%、8 件不良/16000 件出货

批次数量（推荐）：本次客户投诉/退货的具体件数？
  例：5 件、12 件（指本次 8D 分析的样本件数，不是生产批量）
```

❌ 违规示例（禁止出现）：
```
不良率：如 5.2%     ← 52000 PPM，属停发事故级，不应作常规示例
不良率：如 11.5%    ← 115000 PPM，召回级
批次数量：如 2000 根 ← 混淆生产批量与分析样本
批次数量：如 500     ← 同上
```

### 10.4 模板预填值的合规性

`templates/*/template.json` 中所有 `{defect_type} 不良率从 ____% 降至 ____%` 这类占位符保持 `____` 空白由用户填写即可；预填的具体数值示例（如 verification 字段、5why_examples.md 中的范例数据）必须符合 10.1 节的 PPM 区间。
