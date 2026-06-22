---
name: 8d-skill
description: "汽车行业IATF 16949 8D问题解决报告生成器。根据产品名和缺陷描述智能匹配5套行业模板（涂装/装配/焊接/尺寸/通用），自动生成单 Sheet .xlsx和标准.docx格式8D报告。覆盖D0-D8全流程：5Why完整预填路径、6M鱼骨图排查（含判定）、RC根因总结、CA纠正措施、D7 Yokoten横向展开。纯本地生成。Use when user needs 8D report for automotive quality issues, customer complaints, SCAR, or root cause analysis."
---

# 8D Skill — 汽车行业 8D 问题解决报告生成器

> **版本**：2.0.0
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
| defect_rate（不良率） | 推荐 | 11.5% | "目前的不良率/不良数量是多少？" |
| batch_size（批次数量） | 推荐 | 500 | "本批次数量多少？" |
| output_dir（输出目录） | 可选 | ~/Desktop | 默认 ~/Desktop |

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

### Step 3：读取模板并融合用户输入

读取 `templates/<slug>/template.json`，做以下处理：

1. **替换占位符**：将模板中所有 `{defect_type}`、`{product_name}`、`{customer}` 替换为用户实际输入
2. **填充 5W2H**：将 product、defect、customer、defect_rate、batch_size 填入 D0-D2 信息表
3. **生成 8D 编号**：格式 `8D-{YYYYMMDD}-{HHMMSS}`，例如 `8D-20260621-093015`

> **知识增强提示**：如果用户描述的缺陷类型在模板 `defect_types` 中匹配度较低，或涉及特殊工艺/材料，Agent 可使用 `web_fetch` 搜索以下关键词获取真实参考资料：
> - `"{defect_type} root cause analysis automotive"`
> - `"{product_type} manufacturing defects 8D"`
> - `"IATF 16949 8D problem solving"`
> - `"AIAG CQI process control"`
>
> 搜索结果应融入 template.json 的 5Why 路径和 6M 排查方向中，不要完全依赖预制模板。

### Step 4：调用生成脚本

```bash
python skills/8d-skill/scripts/generate_8d.py \
  --product "前保险杠总成（涂装件）" \
  --defect "漆面颗粒/杂质" \
  --customer "XX 汽车有限公司" \
  --defect-rate "11.5%" \
  --batch-size "500" \
  --template paint-defect \
  --output-dir data/export
```

脚本自动：

- 检查并尝试自动 pip install openpyxl、python-docx
- 加载对应模板 `template.json`
- 生成 `8D报告_{product}_{defect}_模板.xlsx`（单 Sheet）
- 生成 `8D报告_{product}_{defect}.docx`（标准 8D Word 文档）
- 将两个文件路径以 JSON 形式打印到 stdout 供 Agent 解析

### Step 5：向用户展示结果（仅返回摘要，禁止输出完整报告）

**⚠️ 关键约束：绝对禁止将 8D 报告的完整内容以 Markdown 文本形式输出到对话中。** 报告内容已在 .xlsx 和 .docx 文件中，重复输出会消耗大量输出 token 预算，导致文件生成链路被截断。

向用户返回（限 5 行以内）：

- xlsx 文件路径 + 一句话说明（"单 Sheet 完整 Excel，可继续编辑"）
- docx 文件路径 + 一句话说明（"标准 Word 文档，可直接提交客户"）
- 已匹配的模板 slug 和预填的 5Why 路径摘要（≤3 句）
- 提示用户："空白处已用 ____ 标记，请补充实际数据后提交客户"

**禁止项：**
- ❌ 禁止在对话中逐条列出 D1-D8 的完整表格内容
- ❌ 禁止在对话中输出鱼骨图、5Why 分析、遏制措施清单等完整文本
- ❌ 禁止以"先展示内容预览再生成文件"的方式操作

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

### template.json 数据结构（v2.0）

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

> 我是一家汽车外饰件厂，生产的后保险杠（注塑件）装配到车身后，客户反馈两侧间隙不一致，间隙超差约 2mm，不良率 8%，需要 8D 报告。

**Agent 处理**：

1. 提取信息：product="后保险杠（注塑件）"、defect="两侧间隙不一致，超差约 2mm"、customer="（缺失，需追问）"、defect_rate="8%"、batch_size="（缺失，需追问）"
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
4. **不留空**：所有用户需填写的空白处用 `____` 标记
5. **信息不足时主动追问**：用 `AskUserQuestion` 工具引导用户补充
6. **不截断**：宁可文字精炼，也不省略号截断
7. **禁止内联输出完整报告**：绝对不要将 D1-D8 的完整表格、鱼骨图、5Why、遏制措施等内容以 Markdown 文本输出到对话中。报告内容已在生成的 .xlsx/.docx 文件中，内联输出会消耗大量 token 预算导致文件生成被截断。Step 5 仅允许返回文件路径 + 简短摘要（≤5 行）

---

## 九、文件清单

```
8d-report/
├── SKILL.md                              （本文档）
├── VERSION                               （版本号：1.0.0）
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
