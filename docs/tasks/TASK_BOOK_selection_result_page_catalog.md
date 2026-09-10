# 任务书：甄选结果签批表页面接入字段目录（最后一张两套口径的表）

> **触发**：2026-09-08 会审纪要收口时复核发现，2026-09-10 用户指定处理。
> 这是**唯一一张页面完成度与目录完成度不一致**的模板。
>
> 影响是实的：AI 能写 `gen_zx_industry` / `gen_zx_std_plan` / `gen_zx_content_desc`，
> **写完页面完成度不动**；用户在页面看到的缺项与在聊天问到的**不是同一份**。

## 现状：14 vs 16，且条目不对应

- 页面：`TemplateForms.tsx:2223-2243` 的 `selectionResultCompletionItems`，**手写 14 项数组**
- 目录：`catalog.json` 的 `selection`，**16 项**
- 另外三张表（需求导入表 / 立项签批表 / 会审纪要）都已走 `getCatalogCompletion`

### 逐项映射（已核对）

| 页面 14 项 | 目录字段 | 说明 |
| --- | --- | --- |
| 甄选后方案 | `post_selection_scheme` (check) | ✔ |
| **合并项目名称** | **无** | ❌ **目录缺这一项** |
| 项目背景 | `gen_proj_bg` (derived) | ✔ |
| 中选合作伙伴 | `gen_zx_winner_name` | ✔ |
| 甄选范围 | `gen_zx_scope` | ✔ |
| 甄选方式 | `gen_zx_method` | ✔ |
| 甄选规则 | `gen_zx_rule` | ✔ |
| 供应商是否中小企业 | `gen_zx_is_sme` | ✔ |
| 收入侧收款方式 | `gen_rev_collection` | ✔ |
| 支出侧付款方式 | `gen_exp_payment` | ✔ |
| 公共字段一致 | `public_fields_consistent` (check) | ✔ |
| 批次字段差异已确认 | `batch_overrides_acknowledged` (check) | ✔ |
| 续签成本归类已确认 | `renewal_costs_confirmed` (check) | ✔ |
| 立项金额低于50万元 | `approval_amount_below_500k` (check) | ✔ |
| — | **`gen_zx_content_desc` 合作内容描述** | ❌ **页面未计数** |
| — | **`gen_zx_industry` 行业** | ❌ **页面未计数** |
| — | **`gen_zx_std_plan` 标准方案** | ❌ **页面未计数** |

**那三个"页面未计数"的字段是真实渲染在页面上的**（各 5 处引用，与
`gen_zx_winner_name` 同一模式），只是没进完成度清单。**页面现在是少算的。**

## 已确认的关键事实

### 1. 这张表拖到最后，不是因为被忘了，是因为一半的项算不出来

页面 6 个 check 项的判定式依赖的状态，**分成两类**：

**（a）在保存态里**（`TemplateForms.tsx:947-952` 的 save payload，与
`projectScale` / `hasMidThree` 同级的根键）：

`selectionResultMode`、`selectionBatchProjectIds`、`selectionBatchName`、
`selectionBatchNameCustomized`、`selectionRenewalDecisions`、`selectionConflictAcknowledged`

**（b）运行期派生，不落库**：

`selectionSharedConflicts`（`:573-574` 拆成 blocking / override）、
`selectionBatchProjects`（批次项目实际加载结果）、
`currentSchemeStage`、`selectionApprovalAmountPreview`（`:2207`，来自测算）、
`selectionRenewalProjectsPreview`

**逐项判定能力：**

| 页面项 | 判定式 | 保存态能否算 |
| --- | --- | --- |
| 甄选后方案 | batch：ids≥2 且 projects 已全部加载／single：`currentSchemeStage === "post_selection"` | ❌ 两支都要派生数据 |
| **合并项目名称** | `mode === "single" \|\| hasText(selectionBatchName)` | ✅ **全在保存态** |
| 公共字段一致 | `mode === "single" \|\| blockingConflicts.length === 0` | ❌ conflicts 派生 |
| 批次字段差异已确认 | `mode === "single" \|\| overrideConflicts.length === 0 \|\| acknowledged` | ❌ conflicts 派生 |
| 续签成本归类已确认 | `renewalPreview.every(p => decisions[p.projectId])` | ❌ preview 派生 |
| 立项金额低于50万元 | `approvalAmountPreview.lt(500000)` | ❌ 测算派生 |

→ **只有"合并项目名称"能变成真正的目录字段，其余 6 项仍是外部事实。**

> 🔴 **更正二（2026-09-10 开工核查）**：**外部事实是 6 项，不是 5 项。**
> 漏掉的是**项目背景**（`gen_proj_bg`）——本表那份是无 `completionSources` 的
> 光秃秃 `derived`，页面靠 `hasText(projectBackground)` 判定。
>
> 而 `projectBackground` **不在模板保存态的根键里**（save payload 只有
> `itContent` / `ctContent` / `projectScale` / `selectionResultMode` 等，
> 见 `TemplateForms.tsx` 的 `const payload = {...}`），
> 它是项目级数据在页面运行期取到的。→ **聊天侧同样只能是 unknown。**
>
> ⚠ 顺带记一个不一致（**本轮不要动**）：同一个 `gen_proj_bg`，
> meeting 用 `completionSources`、approval/selection 用 `completionValues`。
> 按本书第 2 条的判据（"事实不在保存态里就用 `completionValues`"），
> **meeting 那份才是异类**。但它是回归基线，改它风险大于收益，**留待日后**。

### 2. 🔴 本次 `completionValues` 是**合法机制**，不是绕过——与会审纪要那本的禁令不冲突

[会审纪要任务书](./TASK_BOOK_meeting_review_and_list_fields.md) 写过
"**不得用 `completionValues` 绕过**"。**那条禁令在这里不适用，必须说清区别：**

| | 会审纪要（禁止） | 本次（允许） |
| --- | --- | --- |
| 场景 | 目录**有能力**表达（条件必填），但契约有缺口 | 目录**根本无法知道**的外部事实 |
| 用它的后果 | 掩盖缺口，页面通过而聊天仍错 | 页面用它、聊天报 `unknown`，**两边都正确** |
| 正确做法 | 补通用契约 | **就该用它** |

**判据：如果这个事实存在于保存态里，就用 `completionSources` / `requiredWhen`；
只有页面运行期才知道的，才用 `completionValues`。**

> 🔴 **更正一（2026-09-10 开工核查，任务书作者的错）**：
> 本节原写"`completionValues` 目前是死代码，请一并清理 `TemplateForms.tsx:2218`"。
> **完全错了，那处是活的，删了立项签批表会从 7/7 变成 6/7**（执行方实测）。
>
> 错因：`gen_proj_bg` **只有 `meeting` 那份带 `completionSources`**；
> `approval` 与 `selection` 的都是光秃秃的 `derived`。我查了 meeting 的定义，
> 就把结论套到了 `:2218`——**而 `:2218` 是 `approvalCompletionItems`，不是 meeting。**
>
> → **`completionValues` 不是死代码，它现在就在支撑立项签批表的项目背景。**
> **不要删除 `:2218`。原判定保持不变。**

### 3. `check` 类字段缺 `completionValues` 时按 `unknown` 处理，这是对的

`catalog.ts:55`：`kind === 'derived' || kind === 'check'` → `evaluated = false, filled = false`。
→ 聊天侧问"甄选结果签批表还缺什么"，那 6 项会报**未知**而不是**未完成**。

**这是正确行为，不是缺陷。** 聊天确实不知道批次冲突和测算金额。
**工具描述与模型口径必须如实说明"这几项需在页面查看"**，
不得把 `unknown` 说成"已完成"或"未完成"。

### 4. 通用目录能力已足够，本次**不需要**再补契约

会审纪要那轮补齐的能力正好覆盖本次所需：

- `requiredWhen: { field, equals }`（等值条件）→ 合并项目名称按 `selectionResultMode` 判定
- `completionSources` → 从根键读取存在性（**本次用它，不用 `stateKey`**）
- ⚠ 构建期校验：**`stateKey` 只允许 `kind: "text"`**（`build-contract.mjs:21`）——
  这正是更正三的由来，**不要试图给 `derived` 加 `stateKey`**

→ **本件应当是纯数据活 + 一处页面替换**，与会审纪要第一件同形。
**若发现必须改通用契约代码，先停下来汇报**（这条规矩已连续四轮生效）。

### 5. ⚠ 页面完成度数字会从 14 变成 17，用户可见

- +3：合作内容描述 / 行业 / 标准方案（原本渲染了但没计数）
- 合并项目名称从手写移入目录，不增减

其中**行业**（默认 `/`）与**标准方案**（默认 `竞价法`）有默认值，多数项目会直接算已填；
**合作内容描述无默认值**，未填的项目会新增一个缺项。

→ **原本显示 14/14 的项目，修复后可能显示 16/17。**
**这不是回归，是页面原先少算了。** 变更说明里必须写清，
否则用户会当成新 bug——与甄选费定额档"数字会变"同一类。

## 要做的事

1. **`catalog.json` 的 `selection` 增加"合并项目名称"**：

```json
{ "key": "selection_batch_name", "label": "合并项目名称", "kind": "derived",
  "completionSources": [{ "field": "selectionBatchName" }],
  "requiredWhen": { "field": "selectionResultMode", "equals": "batch" },
  "reason": "批次合并名称由页面维护（含自动命名），不开放 AI 写入" }
```

> 🔴 **更正三（2026-09-10 开工核查）**：本条原写
> "`stateKey: \"selectionBatchName\"`" **且**"建议 `kind: derived`"——**两者互斥**。
> `build-contract.mjs:21` 规定 `stateKey` 仅允许 `kind === "text"`，
> 所以 `derived + stateKey` **会被构建校验直接拒绝**。
>
> → **正确写法是 `derived` + `completionSources`，不带 `stateKey`**
> （`completionSources` 无 `source` 时即读根键，与 meeting 的
> `gen_mid_three` 同形，且它已有 `requiredWhen` 共存的先例）。
> 执行方的判断正确。

2. **保持 `derived`（不对 AI 可写）**：保存态里有 `selectionBatchNameCustomized` 标志，
   **说明存在自动命名逻辑**，AI 写入可能与之互相覆盖。
   `kind: "text"` 会让写工具自动暴露该字段（`build-contract.mjs:26-27`），本次不要。

3. **`selectionResultCompletionItems` 改为调用 `getCatalogCompletion`**，
   与另外三张表一致；**保持原有条目顺序**，新增的三项排在合理位置并说明。

4. **页面通过 `completionValues` 提供 6 个外部事实**（更正二）：
   `gen_proj_bg` / `post_selection_scheme` / `public_fields_consistent` /
   `batch_overrides_acknowledged` / `renewal_costs_confirmed` /
   `approval_amount_below_500k`，**判定式逐字沿用现有表达式，不得重写**。

5. **不要动 `TemplateForms.tsx:2218`**（更正一）。立项签批表的原判定保持不变，
   它现在就靠这处 `completionValues` 维持 7/7。

6. **模型口径**：工具描述说明这 6 项在聊天侧为 `unknown`，需在页面查看；
   **不得把 unknown 表述成已完成或未完成**（关键事实 3）。

## 不要做的事

- **不要为了让聊天也能算这 6 项而把派生数据写进保存态**。
  它们会立刻过期（批次冲突、测算金额随时变），**存下来的旧结论比 unknown 更危险**。
- **不要重写那 6 个判定式**。逐字沿用页面现有表达式，只是换个承载方式。
- **不要因为"目录里有 16 项"就假设页面也该正好 16 项**——正确答案是 17。
- **不要动另外三张表**的完成度数字、缺项清单与顺序（回归基线，已连续两轮如此要求）。
- **不要在本件里补通用契约**（关键事实 4）；真需要就停下来汇报。
- 不改甄选批量模式逻辑、冲突判定、续签归类、`handleGenerate`。

## 验证要求

1. **三表回归**：需求导入表 / 立项签批表 / 会审纪要的完成度数字、缺项、顺序
   **逐项不变**，且页面与 `read_template_fields` 两条路径一致。
2. **本表逐项对照**：改造前后，甄选结果签批表的
   **14 项判定结果全部一致**（新增 3 项另行核对）。
   分别覆盖 **single 与 batch 两种模式**、有/无冲突、已/未确认。
3. **合并项目名称**：batch 模式下留空 → 缺项；填写 → 完成；
   **single 模式下不要求**（`requiredWhen` 生效）。
4. **两条路径口径**：聊天问缺项时，6 个外部事实报 **unknown**，
   其余项与页面一致；**模型没有把 unknown 说成已完成**。
5. **数字变化如实**：确认页面从 14 项变为 17 项，
   且三项新增的判定与实际填写状态相符（关键事实 5）。
6. **零代码判据**：除 `selectionResultCompletionItems` 那处替换外，
   `catalog.ts` / `template_catalog.rs` / `template_read.rs`
   与构建脚本**源码哈希不变**。
   ⚠ 原文这里还写着"与死代码清理"，**已随更正一取消**——`:2218` 不动。
6b. **立项签批表仍为 7/7**：这是更正一的直接回归项，单独验一次。
7. **AI 写入联动**：让 AI 写 `gen_zx_industry`，确认**页面完成度随之变化**
   （这正是本件要修的原始症状）。
8. 前端 lint/build、插件 typecheck、`cargo test`、相关回归通过；
   结果记入 `docs/verification/`。

## 附：这张表是四张表统一的最后一块

统一之后，四张模板的完成度**全部由同一个纯函数产出**
（`getCatalogCompletion`，且插件侧由 `build-contract.mjs:37-39` 从该源文件直接编译）。

到那时"页面与聊天口径不一致"这一类问题**在结构上不再可能**——
这与 D 轮 `prepareIctTaxItemsInclBatch` 把"校验与写入"合成一个纯函数是同一种收束。
