# CHANGELOG

## 2.0.3 (2026-06-23)

### 优化：版本号一致性 + PPM 互转 + README + 断点精简

**问题背景**：v2.0.2 审查发现 4 处可优化点：
1. SKILL.md 第 8 行版本号写死 `2.0.0`，与 VERSION 文件不一致
2. SKILL.md 第 175 行「template.json 数据结构（v2.0）」标签过时
3. generate_8d.py 不识别 PPM 格式的 `--defect-rate`，与 SKILL.md 示例用的 `500PPM` 不匹配
4. 缺 README.md，GitHub 仓库主页空白
5. 5why_examples.md 每个范例的「第 5 层达到根因」断点与开头通用表重复

**修复内容**：

1. **版本号一致性修正**：
   - `SKILL.md:8` `> **版本**：2.0.0` → `> **版本**：2.0.3`
   - `SKILL.md:175` `### template.json 数据结构（v2.0）` → `### template.json 数据结构（v2.0+）`
   - `SKILL.md:263` 文件清单 `（版本号：1.0.0）` → `（版本号：2.0.3）`

2. **generate_8d.py 增加 PPM ↔ % 互转**：
   - 新增 `normalize_defect_rate(value)` 函数，接受 `500PPM` / `0.05%` / `500` 等格式
   - 自动转换为双格式输出：`500 PPM (0.0500%)`，便于客户/内部双向理解
   - 纯数字（无 PPM/% 后缀）按行业常识判断：≤1 视为百分比，>1 视为 PPM
   - main 函数增加 >0.5%（5000 PPM）警告：标注「停发/召回级事故，建议 D0 阶段标注严重度=高」
   - 遵循 SKILL.md 第十章「行业常识基准」

3. **新增 README.md**：
   - 仓库主页 README，含快速开始、模板选择表、行业常识基准、文件结构、标准引用、版本历史
   - GitHub 仓库主页不再空白

4. **5why_examples.md 断点精简**（-13 行）：
   - 删除 5 个范例中重复的「- ✅ 第 5 层达到根因：...」断点（与开头通用表重复）
   - 「### 断点提示」重命名为「### 差异化断点（通用断点见开头表格）」
   - 188 行 → 175 行

**预期效果**：
- 用户看到的版本号与实际一致
- 用户输入 `500PPM` 或 `0.05%` 都能正确显示为双格式
- GitHub 仓库主页有专业 README
- 5why_examples.md 减少重复内容，token 占用降低

---

## 2.0.2 (2026-06-23)

### 真实化优化：行业标准引用 + IATF 条款修正

**问题背景**：v2.0.1 解决了"无常识示例值"问题，但审查发现：
1. `8d_guide.md` 中多处 IATF 16949 条款号引用错误（如把"问题解决"标为 10.2.3，实际是 10.2.2；引用了不存在的 8.5.2.2）
2. 模板和 references 缺少真实行业标准引用，导致 LLM 写报告时无权威依据
3. 部分内容与 template.json 重复，token 浪费

**修复内容**：

1. **`8d_guide.md` IATF 16949 条款号全面修正**（按 IATF 16949:2016 真实条款）：
   - `9.1.1.2（监视和测量资源）` → `9.1.1.1（制造过程监视与测量）` ✓
   - `10.2.3（问题解决）` → `10.2.2（问题解决）` ✓
   - `10.2.2（不合格和纠正措施）` → `10.2.1（不合格与纠正措施——总则）` ✓
   - `8.5.2.2（防错）` → `10.2.3（防错）` ✓（8.5.2.2 在 IATF 16949:2016 中不存在）
   - D7 新增 `8.5.1.1（控制计划）` 和 `8.5.6（更改控制）` 引用 ✓
   - D8 新增 `7.3（意识）` 引用 ✓

2. **`SKILL.md` 新增「行业标准参考」章节**（11 项真实标准）：
   - IATF 16949:2016、AIAG-VDA FMEA 1st Ed. (2019)、AIAG PPAP 4th Ed.、AIAG APQP 2nd Ed.、AIAG MSA 4th Ed.、AIAG SPC 2nd Ed.、VDA 6.3、Ford 8D (1986)、Toyota 5Why/Yokoten、Quality-One 8D、AIAG CQI-9/11/12/15/23
   - 每项标注发布机构、版本号、在 8D 中的引用场景
   - 增加约束："Agent 不应编造标准号或版本；如不确定，宁可不引用"

3. **5 套 `templates/*/intro.md` 新增「适用标准参考」章节**（按缺陷类型匹配真实 CQI 标准）：
   - `welding-defect` → CQI-15 (Welding, 2nd Ed. 2010) + AWS D17.2 + ISO 15614-1 + ISO 5817
   - `paint-defect` → CQI-12 (Coating, 2nd Ed. 2019) + ISO 11997-1 + ASTM B117/D3359/E1349 + VDA 6.3
   - `dimensional-defect` → AIAG PPAP/SPC/MSA + ISO 1101 (GPS) + ISO 2768 + VDA 6.3 P6.2
   - `assembly-defect` → CQI-14 (Warranty) + ISO 898-1/16048 + VDA 6.3 P6.4 + IATF §8.5.1.2
   - `generic-defect` → IATF §10.2.2 + AIAG-VDA FMEA + Toyota 5Why + Ishikawa 6M

4. **精简冗余内容（减少 token 占用 6.1%）**：
   - `8d_guide.md`：删除每个 D 阶段重复的「IATF 16949 关联」小节（保留末尾总览表），-35 行
   - `5why_examples.md`：合并底部「5Why 推演通用建议」和「5Why 与 6M 的关系」为精简版，-8 行
   - `fishbone_guide.md`：删除 ASCII 鱼骨图绘制规范和 Mermaid 示例（agent 输出表格即可），-67 行

5. **总大小**：始终加载到 system prompt 的 5 个文件从 65564 bytes → 61586 bytes（-6.1%），新增内容（行业标准参考）小于删除的冗余内容，净减少 token 占用。

**预期效果**：
- LLM 撰写报告时能引用真实权威标准（CQI-12/-15、AIAG-VDA FMEA、IATF 条款号正确）
- 不再出现编造的标准号或错误的条款引用
- system prompt 占用降低，响应速度不会因丰富内容而变慢

---

## 2.0.1 (2026-06-23)

### 修复：行业常识基准硬约束（解决"无常识示例"问题）

**问题背景**：v2.0.0 中所有不良率示例值（11.5% / 8% / 5% / 3%）与批次数量示例值（500）严重脱离汽车零部件行业实际，导致 Agent 在 AskUserQuestion 追问时模仿这些数字、生成"5.2% 不良率 + 2000 根批次"这类违反常识的提示。

**修复内容**：
1. **SKILL.md**：
   - defect_rate 默认示例 11.5% → **500 PPM (0.05%)**
   - batch_size 默认示例 500 → **12 件**（语义澄清为"8D 分析样本数"，非生产批量）
   - 调用示例 8% → 200 PPM (0.02%)
   - 新增第十章「行业常识基准（🔴 严禁违反）」，含 PPM 量级表、batch_size 分类表、合规/违规示例对照
   - 第八章新增第 8 条「行业常识优先」硬约束
2. **references/5why_examples.md**：5 个范例的不良率全部改为 PPM 级（50/100/200/500 PPM），并补充"退货件数/出货总量"的折算说明
3. **references/fishbone_guide.md**：漆面颗粒示例不良率 11.5% → 500 PPM
4. **references/8d_guide.md**：「不良率11.5%」→「不良率 500 PPM，本批次退货 15 件」
5. **scripts/generate_8d.py**：文档头示例参数 11.5% / 500 → 500PPM / 15，新增「⚠️ 行业常识基准」说明段
6. **app/agent/prompts.py**：8D 硬约束新增第 5 条「行业常识基准」，明确禁止 3%/5%/5.2%/8%/11.5% 与 500/2000/5000 作为示例值，并要求 batch_size 概念混淆时主动追问澄清

**预期效果**：Agent 在追问用户不良率/批次数量时，给出的示例值将落在 IATF 16949 行业常识区间内（PPM 级不良率 + 个位数到几十件的分析样本），不再出现"5.2% + 2000根"这类灾难级数字。

---

## 2.0.0 (2026-06-21)

### 重大升级：从模板框架到可用的 8D 报告生成器

**template.json v2.0**
- 5Why: hints(候选列表) -> steps(完整连贯预填路径，含答案和证据)
- 6M: 纯文本 -> 数组格式 + judgment字段(根本原因/贡献因子/排除/待确认)
- 新增 root_cause_summary: 预填RC1/RC2
- 新增 verification: 预填验证结论(含具体数据)
- 5套模板全部重写

**generate_8d.py 视觉升级**
- 章节标题: 深蓝底 -> 黄色底+深蓝字, 视觉分层
- Sheet: 单 Sheet，标签颜色 003366 深蓝
- D4生成器: 兼容新旧template格式
- Word同步更新

**SKILL.md**
- 新增第5条价值主张: 真实行业参考资料
- 新增v2.0 template.json数据结构文档
- 新增web_fetch知识增强搜索指引
- 新增YAML frontmatter(name+description)

**references/**
- 基于Quality-One 8D方法论补充真实行业实践内容

---

## 1.0.0 (2026-06-20)

### 初始版本
- 完整目录结构: SKILL.md + templates/ + references/ + scripts/
- 5套行业模板
- scripts/generate_8d.py (xlsx+docx双输出)
- references/: 8d_guide.md / 5why_examples.md / fishbone_guide.md
