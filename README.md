# 8D Skill — 汽车行业 8D 问题解决报告生成器

> **版本**：2.0.3
> **适用场景**：汽车零部件质量工程师、SQE、CQE 处理客户投诉、SCAR、根因分析
> **输出物**：8D 报告 .xlsx（单 Sheet）+ 8D 报告 .docx（标准文档）
> **依赖**：openpyxl、python-docx（纯本地生成，无外部 API）

---

## 一、这是什么

按 IATF 16949:2016 体系要求，自动生成规范的 8D（Eight Disciplines）问题解决报告。用户只需提供产品名、缺陷描述、客户名等基础信息，Skill 自动：

1. 按 5 套行业模板（涂装/装配/焊接/尺寸/通用）智能匹配
2. 预填完整 5Why 路径（6 步：问题 + Why1-5）
3. 预填 6M 鱼骨图排查项（含判定字段）
4. 预填 RC1/RC2/RC3 根因总结、验证结论、CA 措施、Yokoten 横向展开
5. 同时输出单 Sheet Excel 和标准 Word 文档

## 二、快速开始

```bash
python3 scripts/generate_8d.py \
  --product "前保险杠总成（涂装件）" \
  --defect "漆面颗粒/杂质" \
  --customer "XX汽车有限公司" \
  --defect-rate "500PPM" \
  --batch-size "12" \
  --template paint-defect \
  --output-dir ~/Desktop
```

**参数说明**：

| 参数 | 必填 | 说明 |
|---|---|---|
| `--product` | ✅ | 产品名称 |
| `--defect` | ✅ | 缺陷描述 |
| `--customer` | ✅ | 客户名称 |
| `--defect-rate` | 推荐 | 不良率，接受 `500PPM` / `0.05%` / `500` 等格式，自动互转 |
| `--batch-size` | 推荐 | 本次 8D 分析的客户投诉/退货样本件数（**不是生产批量**） |
| `--template` | ✅ | 模板 slug：`paint-defect` / `assembly-defect` / `welding-defect` / `dimensional-defect` / `generic-defect` |
| `--output-dir` | 可选 | 输出目录，默认 `~/Desktop` |

## 三、模板选择

| 用户描述关键字 | 模板 slug | 适用 CQI 标准 |
|---|---|---|
| 漆面、涂装、颗粒、流挂、色差、橘皮、缩孔 | `paint-defect` | CQI-12 (Coating) |
| 装配、间隙、面差、卡扣、异响、松动 | `assembly-defect` | CQI-14 (Warranty) |
| 焊接、虚焊、焊穿、焊渣、焊点、强度 | `welding-defect` | CQI-15 (Welding) |
| 尺寸、超差、CPK、公差、变形、收缩 | `dimensional-defect` | AIAG PPAP/SPC/MSA |
| 其他/无法明确分类 | `generic-defect` | AIAG-VDA FMEA |

## 四、行业常识基准（🔴 严禁违反）

- **不良率**：汽车零部件量产线 IATF 16949 目标 ≤50 PPM；客户投诉触发 8D 的典型量级 100–2000 PPM（0.01%–0.2%）；>0.5%（5000 PPM）属于停线停发事故。默认示例用 **500 PPM (0.05%)**，安全件用 **50 PPM (0.005%)**。
- **批次数量**：指本次 8D 分析的客户投诉/退货样本件数，不是生产批量。线束/ECU 类典型 1–10 件，保险杠/仪表板总成类典型 3–20 件。默认示例用 **12 件**（线束类用 **5 件**）。
- 完整规则参见 [`SKILL.md` 第十章](SKILL.md)。

## 五、文件结构

```
8d-skill/
├── SKILL.md                  # Skill 主文档（含工作流、行业常识基准、行业标准参考）
├── VERSION                   # 版本号
├── CHANGELOG.md              # 变更日志
├── README.md                 # 本文档
├── scripts/
│   └── generate_8d.py        # 核心生成脚本（xlsx + docx 双输出）
├── references/
│   ├── 8d_guide.md           # 8D 方法论详细指南（D0-D8 每步关键活动）
│   ├── 5why_examples.md      # 5 个完整 5Why 范例
│   └── fishbone_guide.md     # 6M 鱼骨图分析指南
└── templates/
    ├── INDEX.md              # 模板索引
    ├── paint-defect/         # 涂装缺陷模板
    ├── assembly-defect/      # 装配缺陷模板
    ├── welding-defect/       # 焊接缺陷模板
    ├── dimensional-defect/   # 尺寸超差模板
    └── generic-defect/       # 通用兜底模板
```

## 六、行业标准引用

本 Skill 在 D4/D5/D6/D7 阶段会引用以下真实标准（详见 `SKILL.md` 第五章「行业标准参考」）：

- **IATF 16949:2016**（国际汽车工作组）
- **AIAG-VDA FMEA 1st Ed. (2019)** / **PPAP 4th Ed.** / **APQP 2nd Ed.** / **MSA 4th Ed.** / **SPC 2nd Ed.**
- **VDA 6.3**（过程审核）
- **AIAG CQI-9/11/12/15/23**（特殊过程评估：热处理/电镀/涂装/焊接/成型）
- **Ford 8D (1986)**（8D 方法论起源）/ **Toyota 5Why-Yokoten**

## 七、版本历史

- **2.0.3** (2026-06-23)：版本号一致性修正 + PPM/% 互转 + README + 断点精简
- **2.0.2** (2026-06-23)：IATF 条款号全面修正 + 行业标准引用 + 冗余精简
- **2.0.1** (2026-06-23)：行业常识基准硬约束（解决"无常识示例值"问题）
- **2.0.0** (2026-06-21)：从模板框架升级为可用的 8D 报告生成器

完整变更日志见 [`CHANGELOG.md`](CHANGELOG.md)。

## 八、许可

本 Skill 由东风科技研发智能体平台团队维护，仅供内部使用。
