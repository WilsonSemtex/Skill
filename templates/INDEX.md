# 8D 报告模板索引

> 5 套行业模板，覆盖汽车零部件常见缺陷类型。根据用户输入的产品类型和缺陷描述智能匹配模板。

## 完整模板清单

| slug | 模板名称 | 适用场景 | 一句话特点 |
|---|---|---|---|
| `paint-defect` | 涂装/漆面缺陷 | 保险杠、车身面板涂装件 | 漆面颗粒、流挂、色差、橘皮 |
| `assembly-defect` | 装配缺陷 | 内饰件、外饰件装配 | 间隙面差、卡扣不到位、异响 |
| `welding-defect` | 焊接缺陷 | 金属结构件、支架 | 虚焊、焊穿、焊渣、强度不足 |
| `dimensional-defect` | 尺寸超差 | 注塑件、冲压件 | 关键尺寸超差、CPK 不足 |
| `generic-defect` | 通用缺陷 | 其他/未分类 | 兜底模板，根据描述自适应 |

## 按缺陷类型快速筛

| 缺陷类型 | 候选模板 |
|---|---|
| 漆面/涂装 | paint-defect |
| 装配/间隙 | assembly-defect |
| 焊接/连接 | welding-defect |
| 尺寸/公差 | dimensional-defect |
| 其他 | generic-defect |

## 按产品类别快速筛

| 产品类别 | 首选模板 |
|---|---|
| 保险杠、车身面板、外饰涂装件 | paint-defect |
| 内饰件、外饰装配件、座椅 | assembly-defect |
| 金属支架、排气管、车身焊接件 | welding-defect |
| 注塑件、冲压件、压铸件 | dimensional-defect |
| 紧固件、橡胶件、电子件等 | generic-defect |

## 模板匹配优先级

1. **第一优先级：缺陷描述关键字**——若用户明确说出"漆面颗粒""虚焊""尺寸超差"等行业术语，直接对应模板
2. **第二优先级：产品类别**——若缺陷描述模糊（如"质量有问题"），按产品类别匹配
3. **兜底：generic-defect**——以上两者均无法明确匹配时使用

## 模板数据结构

每个 `template.json` 包含以下字段：

- `slug`：模板唯一标识
- `name`：模板中文名
- `defect_types`：覆盖的缺陷类型清单
- `product_categories`：适用的产品类别
- `d0_d2`：D0-D2 阶段提示信息
- `d3_template`：D3 临时遏制措施预填列表（5 项）
- `d4_template`：D4 根本原因分析（含 5Why 路径 + 6M 方向）
- `d5_d6_template`：D5-D6 永久纠正措施预填列表
- `d7_template`：D7 横向展开措施清单

占位符约定：

- `{defect_type}`：替换为用户输入的缺陷描述
- `{product_name}`：替换为用户输入的产品名
- `{customer}`：替换为用户输入的客户名
