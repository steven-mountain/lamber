# 任务书：科目金额清零会永久删除收款计划（恢复机制被一句 delete 架空）

> **触发**：2026-09-10 执行 AI 试算（B 件）桌面对照时撞到——
> 同样把集成收入改成 800000，两条路径给出**不同的 NPV**：
>
> | 路径 | 收款计划 | NPV |
> | --- | --- | --- |
> | AI 试算 | 四年各 200000（原计划按比例缩放） | 591259.19 |
> | 桌面清空再输入 | 首年收完 | 645281.66 |
>
> 利润率两边都是 88.89%（只依赖合计，不受计划影响），**差的是 54022.47 元 NPV**。
>
> 🔴 **结论：AI 那条是对的，桌面那条是坏的。**
> **不要为了让两边一致而去改 AI。**

## 根因：恢复机制写好了，但被上游一句 `delete` 变成死代码

`src-ui/src/lib/ictSubjectFundingPlan.ts` 的 `syncSubjectFundingPlanToAmount`：

```ts
// :771-776   金额 ≤ 0
if (newCents <= 0) {
  if (!existing) return plans;
  const nextPlans = { ...plans };
  delete nextPlans[id];        // ← 计划被整个删掉
  return nextPlans;
}

// :779-783   金额 > 0 且计划不存在 → 建默认 upfront（首年收完）
if (!existing) return { ...plans, [id]: createDefaultSubjectFundingPlan(...) };

// :802-806   金额 > 0 且计划存在 → 从零恢复
const existingTotalCents = existing.annualInclValues.reduce(...);
const isRecoveringFromZero = existingTotalCents === 0;
const baseValues = (isRecoveringFromZero && existing.lastValidAnnualInclValues)
  ? normalizeAnnualInclValues(existing.lastValidAnnualInclValues)
  : existing.annualInclValues;
```

**`:802-806` 这段是专门为"金额归零又回来"写的**，配套还有
`lastValidAnnualInclValues` 字段和 `lastChangeReason: "restored_after_zero"`（`:820`）。

**但它永远跑不到。** 因为 `:774` 已经把计划删了，`existing` 从此为 `null`，
下一次正数金额必然落到 `:779` 的"建默认 upfront"。

> `:803` 计算 `existingTotalCents` 并判 `=== 0`——
> **这行代码的存在本身就说明：设计意图是"把计划清零并保留"，不是"删除"。**
> 删除那句让整套恢复逻辑成了死代码。

## 触发路径：逐键提交 + `Number("") === 0`

`IctLifecycle.tsx:1966`：

```tsx
onChange={e => { ... updateTaxItem(groupId, item.key, 'incl', Number(e.target.value)); }}
onBlur={() =>  { ... commitTaxItemIncl(groupId, item.key); }}
```

- **`onChange` 每敲一个键就提交一次**，清空输入框时 `Number("")` = **0**
- → 走到 `:771` 的删除分支，**原计划当场消失**
- → 再键入 800000 时计划已不存在，建成"首年收完"

⚠ **注意这里已经有 `onBlur → commitTaxItemIncl`**，
说明"编辑中"与"正式提交"的区分**在 UI 层已经存在**，
只是资金计划同步走的是 `onChange` 那条（`useIctState.ts:496-499`）。

## 影响范围：这是今天就在丢真实数据的缺陷，与 AI 无关

**用户只要"全选改数字"就会触发**——这是最常见的编辑手势。
改完保存，原来配好的多年收款计划就**永久变成首年收完**，
本例中 NPV 凭空多出 **54022.47 元（约 8%）**，而且**界面不提示、不报错**。

**已保存的真实项目可能已经丢过计划。** 本书**不做数据迁移**（无从判断哪些是有意改的），
但修复说明里必须写清这一点。

### 另外两个同样调用该同步的入口，需一并回归

`useIctCalculations.ts:432`（甄选限价回填）与 `:990`（智能结构反算）
都走 `updateTaxItemsInclBatch`。**若它们在中间步骤写过 0，同样会删计划。**
本次必须verify这两条路径在修复前后的行为。

## 要做的事

1. **停止删除，改为清零保留**：`:771-776` 的分支不再 `delete`，
   而是把 `annualInclValues` 清零、**写入 `lastValidAnnualInclValues`**、
   置 `enabled: false`，使 `:802-806` 的既有恢复逻辑能够真正生效。
   > **优先采用这个方案**，因为它复用已经写好并测试过的恢复路径，
   > 不需要引入新的"编辑态 / 提交态"语义。

2. **执行方提出的另一方案**（"编辑中暂时清空不删除，正式提交 0 才处理"）
   **作为备选**：它需要把资金计划同步从 `onChange` 挪到 `commitTaxItemIncl`，
   改动面更大，且与第 1 条部分重复。
   **两者选其一即可，不要同时做。** 若选第 2 条，须说明为何第 1 条不够。

3. **确认"真正的 0"仍然正确**：用户确实要把某科目清零并保存时，
   该科目不应再产生现金流。清零保留不得让已归零的科目重新参与测算。

4. **`initializeMissingSubjectFundingPlans` 的配合**（`useIctState.ts:498`）：
   它按 `activePositiveSubjects` 补建缺失计划，已排除刚归零的科目。
   修复后要确认它**不会**把保留下来的零值计划再覆盖成 upfront。

## 不要做的事

- 🔴 **不要为了让桌面与 AI 一致而修改 AI 试算的联动逻辑。**
  AI 那条走的正是 `syncSubjectFundingPlanToAmount` 的比例缩放分支，**是对的**。
- **不要删除 `lastValidAnnualInclValues` / `isRecoveringFromZero` / `restored_after_zero`**
  这套机制——本次是让它生效，不是替换它。
- **不要自动迁移历史项目**已经丢失的计划（无从判断哪些是用户有意改的）。
- 不改折现、NPV、利润率公式，不改 `initializeMissingSubjectFundingPlans` 的补建策略本身。
- 不改甄选费、模板、AI 工具面。

## 验证要求

1. **原始场景复现**：多年计划的科目，全选清空再输入新金额，
   **计划形状保持原比例**，不再变成首年收完。
2. **NPV 一致**：修复后 AI 试算与桌面手工输入同一覆盖，
   **NPV / NPV 率 / 动态回收期 / 利润率逐项一致**（本书触发场景的两个数字应当合一）。
3. **真正清零仍然有效**：把科目改成 0 并保存，确认该科目不产生现金流，
   且重新给正数时按 `restored_after_zero` 恢复原比例。
4. **`restored_after_zero` 确实被走到**：断言 `lastChangeReason`，
   证明恢复分支不再是死代码。
5. **两个批量入口回归**：甄选限价回填（`useIctCalculations.ts:432`）与
   智能结构反算（`:990`）修复前后行为一致，且不会误删计划。
6. **模式覆盖**：`proportional` / `upfront` / `custom` 三种 mode 各验一次。
7. 前端 lint/build、既有资金计划测试、`test:selection-fee`、相关回归通过；
   结果记入 `docs/verification/`。

## 附：这次是怎么被发现的

**不是靠读代码，是靠两条路径算出不同的 NPV。**

`:802-806` 那段恢复逻辑写得很完整，单看没有任何问题；
`:771-776` 的删除单看也很合理。**只有把它们放在一起，才知道后者让前者永远跑不到。**

这与 AI 测算任务书的主线是同一件事：
**错误不在某一行，而在两行之间；通读发现不了，交叉验算才能发现。**
