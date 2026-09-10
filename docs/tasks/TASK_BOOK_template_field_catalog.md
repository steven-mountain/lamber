# 任务书：字段目录泛化——写工具从"需求表专用"变成"目录驱动"

> **来源**：用户提出"模板层之下是数据库，模板填的内容其实都写进数据库"。
> 方向对，但当前存储恰好相反（模板才是存储单元）。
> 本文取**中间路线**：**不动存储，只把字段目录泛化**，
> 让"加一张模板 = 加数据、不加代码"。
>
> **位置**：接在 `TASK_BOOK_ai_write_template_fields.md`（④）之后。
> 前置：④ 的阶段 A / B 与 `read_template_fields` 已完成。

## 目标与判据

**判据只有一条：加第三张模板时，只需要改数据文件，代码零改动。**
达不到这条，就说明没泛化成功，只是把复制粘贴挪了个位置。

## 已确认的关键事实（写代码前不用重新调研）

### 1. 存储是按模板分片的，本次不动它

`project_template_states` 的唯一键是 `UNIQUE(project_id, template_id)`（`db.rs:246-260`），
每张模板存自己的一份 `filled_data_json`。**字段不是一等公民，是模板内部的键。**

用户设想的"字段作为一等公民、模板只是视图"是**更远期的目标架构**，不在本次范围。
理由：那要动保存链路、完成度规则、生成取值、常用资料四条链路，
而其中三条是 ① 那轮刚理顺的（生成取值刚统一到保存态、缺项规则刚抽成纯函数）。
**刚理顺就重构，风险大而收益小**——见关键事实 4，真正重复的字段只有 7 个。

### 2. 现有目录与生成机制（照它扩展，别另起炉灶）

- 目录：`src-ui/src/lib/templateCompletion/demand-fields.json`
  条目形状 `{ key, label, kind, defaultValue?, requiredWhen? }`，
  例如 `{"key":"gen_demand_branch_name","label":"项目需求单位","kind":"text","defaultValue":"XXX分公司"}`。
  **它已经带 `kind`**，且已含非文本条目（如 `techItems` / 技术方案可行性清单）。
- 生成：`agent-bridge/scripts/build-contract.mjs`
  - 过滤 `kind === 'text'` → `templateFields.generated.ts`（写白名单）
  - **直接编译 UI 的纯函数** → `demandCompletion.generated.ts`，
    脚本注释原文：*"Compile the production pure function itself; do not maintain a second completion algorithm."*
    **这条原则要保住**——泛化后仍然只能有一份缺项算法。

### 3. ⚠ 已存在一个"平行"的字段注册表，**但 key 空间不通，不要合并**

`src-ui/src/lib/presetFieldKeys.ts` 的 `PRESET_FIELD_REGISTRY`（53 条）形状是：

```ts
{ fieldKey: "project_basic.background", label: "项目背景",
  templates: ["ICT生命周期测算", "立项签批表", "会审纪要"],
  fieldType: "long_text", kind: "short_value", ... }
```

**它正是"字段 → 属于哪些模板"的形状**，可以借鉴。**但**：

- 它的 key 是**业务语义 key**（`project_basic.background`），
  而表单与 `filled_data_json` 用的是**表单控件名**（`gen_demand_branch_name`）。
- 两套 key **没有任何映射**（在 `presetFieldKeys.ts` 里搜 `gen_` 无结果）。
- 它服务的是"常用资料"功能，不是写入目标。

→ **不要把 `presetFieldKeys` 直接当成新目录，也不要在本次统一两套 key 空间。**
借鉴它的**结构**，新目录仍以表单 key 为写入键。
若将来要做"字段一等公民"，两套 key 的合并是那次的事，不是这次。

### 4. 八张模板的实际情况（已扫描，不用重新调研）

| 模板 | 占位符 | 人工填写字段 | 本次处理 |
| --- | ---: | --- | --- |
| 需求导入表 `.docx` | 22（含 2 图） | 8 个 `gen_demand_*` | ✅ 已做，本次作回归基线 |
| **立项签批表** `.docx` | 25 | 2 个 `gen_sign_*` + 共用 | **本次实现（验证泛化）** |
| 甄选结果签批表 `.docx` | 59 | 8 个 `gen_zx_*` + 共用 | 机制验证后再加 |
| 会审纪要 `.docx` | 35（含 1 图） | 3 个 `gen_meet_*` + 十余个散落 | 最后，且只做文本部分 |
| 立项决策汇报 `.pptx` | 50 | 4 个 `gen_ppt_*` | 单独评估，本次不做 |
| 售前预算表 `.xlsx` | **0** | 无 | **永久排除** |
| 效益分析表 `.xlsx` | **0** | 无 | **永久排除** |
| 立项决策纪要 `.docx` | **0** | 无 | **永久排除** |

最后三张扫描到**零占位符**，且 `docfill.rs` 里没有任何专门处理逻辑：
两张 Excel 走 `internal_generate_xlsx`，数字直接来自测算与科目数据；决策纪要基本是固定文本。
**它们没有人工填写的文本字段，给它们做写工具是白费**——在目录里显式标注排除，
免得以后有人再来做一遍。

全项目人工输入字段共 **52 个**（`gen_*`），已覆盖 8 个，**剩 44 个**。

### 5. 头号陷阱：完成度清单 ≠ 字段目录

四张完成度清单里混着**校验结论**，它们不是字段、写不得。例如甄选结果签批表：

```js
{ label: "公共字段一致",     filled: selectionBlockingConflicts.length === 0 },
{ label: "立项金额低于50万元", filled: selectionApprovalAmountPreview.lt(500000) },
{ label: "续签成本归类已确认", filled: selectionRenewalDecisionsComplete },
{ label: "甄选后方案",       filled: currentSchemeStage === "post_selection" },
```

需求导入表那 11 项恰好几乎全是可填字段，所以 ① 做得顺——**这是特例，不是通例**。
**每张表都必须人工过一遍，把可填字段与校验结论分开**，不能拿完成度清单当目录用。

## 要做的事

1. **扩展目录格式**，在现有 `{key,label,kind,defaultValue,requiredWhen}` 基础上增加：
   - `template`：归属模板（或把目录改为按模板分文件，二选一，在 README 写明理由）
   - `kind` 的取值必须能区分至少这几类，且**只有第一类可写**：
     `text`（可写文本）/ `derived`（测算或项目数据派生）/ `check`（校验结论，关键事实 5）/
     `list`（动态行）/ `image`（图片）
2. **`build-contract.mjs` 改成遍历目录**，而不是硬编码 demand 一份。
   生成的白名单按模板分组；**缺项算法仍然只有一份**（保住关键事实 2 那条原则）。
3. **`fill_template_fields` 与 `read_template_fields` 按 `templateId` 从目录取白名单**，
   不再依赖某张表专用的常量。
4. **本次只实现立项签批表**，用它验证泛化后的机制能跑通。
   甄选结果签批表与会审纪要**不在本次范围**，但目录格式要能容纳它们
   （尤其是会审纪要的 `list` / `image` 条目要能被标注出来并被写工具拒绝）。
5. **三张零字段模板在目录中显式标注排除**并写明原因。

## 不要做的事

- **不动 `project_template_states` 的结构与保存链路**（关键事实 1）。
  本次是目录层泛化，不是数据模型重构。
- **不合并 `presetFieldKeys` 的 key 空间**（关键事实 3）。借鉴结构，不动它。
- **不要维护第二份缺项算法**——`build-contract.mjs` 那条注释是硬要求。
- **不做动态表格与图片的写入**。目录要能*标注*它们，但写工具一律拒绝。
- **不写金额、税率、年限、折现率、测算结果**，服务端拒绝，与 ④ 的边界一致。
- **不给会审纪要开工**——它字段最散、还带表格和截图，等机制在两张简单表上验稳。
- 不为三张零字段模板做任何写入能力。
- 不改 `calculator.rs` / `docfill.rs` / 测算引擎 / 0 容差校验。

## 验证要求

1. **泛化判据（最重要）**：在只改数据文件、不改任何 `.ts` / `.rs` 代码的前提下，
   把甄选结果签批表的字段加进目录，确认写工具立刻能处理它。
   **需要改代码就说明没泛化成功**，本次不算通过。
2. **校验结论不可写**：构造把"立项金额低于50万元""公共字段一致"这类 `check` 条目
   塞进 `fields` 的调用，确认**服务端**拒绝。
3. **立项签批表全链路真人走一遍**（沿用 ④ 的 Gate B 形式）：
   留空一个文本字段 → 聊天让 AI 填 → 审批读新旧对照并改一处 → 批准 →
   回模板页完成度 +1 → 生成 docx 核对改后文本。
4. **需求导入表回归**：泛化后行为与泛化前完全一致，8 个字段一个不少、一个不多。
5. **排除项**：确认三张零字段模板不出现在任何可写白名单里。
6. `cargo test`、`cargo test agent_bridge -- --ignored`（真实 key）、
   前端 lint/build、插件 typecheck、`npm run test:packaging` 全过。
7. 结果记入 `docs/verification/`，照实区分"已验证"与"未验证 / 已知限制"。

## 顺带记下：更远期的"字段一等公民"

用户的原始设想是把字段提升为一等公民、模板只作视图。本次没做，理由见关键事实 1。
但本次产出的**统一字段目录正是那次重构的起点**——届时要解决的是
`presetFieldKeys` 与表单 key 两套空间的合并，以及 `filled_data_json` 从
按模板分片改为按字段存储。**那是一份独立任务书，不要在本次顺手开始。**
