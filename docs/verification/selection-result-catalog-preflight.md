# 甄选结果签批表目录接入：开工核查暂停

日期：2026-09-09。对应 `docs/tasks/TASK_BOOK_selection_result_page_catalog.md`。

用户要求任务书与代码不一致时先停报。本轮仅增加只读见证及记录，未修改业务源码、目录、构建契约、模板保存态或任何金额。此前 D 已完成的工程验证结论不变。

## 1. 要求删除的 completionValues 仍在生效

任务书关键事实2称 `TemplateForms.tsx:2218` 的 `gen_proj_bg` 已被 `completionSources` 覆盖，是死代码，要求清理。

实际：`catalog.json:88` 的 **approval.gen_proj_bg 没有 completionSources**；只有会审的同名字段（`:300`）有。`catalog.ts:53` 的 supplied 分支因此仍在判定立项背景。

运行原 `getCatalogCompletion`，同一已填状态：

| 立项签批表 | 总项数 | 已填 | unknown |
| --- | ---: | ---: | --- |
| 保留当前 supplied 背景 | 7 | 7 | 无 |
| 按任务书删除 supplied 背景 | 7 | 6 | gen_proj_bg |

删除会违反“不动另外三表”的要求。不能将会审同名 key 的声明当作立项目录声明。

## 2. 甄选背景也是聊天未知项，实际是6项而非5项

`catalog.json:141` 的 **selection.gen_proj_bg 同样没有 completionSources**。页面原14项使用 `hasText(projectBackground)`，此值从外部项目参数提供；模板保存 payload（`TemplateForms.tsx:932–960`）没有权威的根 `projectBackground`。

原读工具只投影模板状态；原纯函数即便收到根 `projectBackground`，没有相应来源声明也不会使用。因此本表聊天未知项目前是：项目背景 + 任务书列出的5个check，共 **6项**。若页面只传指定5项，项目背景会由原已填变为unknown，无法保持原14项判定。

## 3. derived + stateKey 与现有构建约束冲突

任务书要求 `stateKey: selectionBatchName`，同时默认建议 `kind: derived`。

`agent-bridge/scripts/build-contract.mjs:21` 明确只允许 text 字段声明 stateKey。直接提取并执行该生产 guard，建议组合抛出 `Invalid root text state key`。

这里不需要扩大契约：`completionSources: [{field: "selectionBatchName"}]` 已经能够引用根值，derived 字段可省略 stateKey。自动命名由 `TemplateForms.tsx:551` 的 customized 标志控制，保持只读更安全。

## 建议修订（待任务书明确后继续）

1. 保留立项签批表现有 `completionValues.gen_proj_bg`，撤回本轮“死代码清理”要求，另外三表保持不动。
2. 甄选页面传入 **6项**外部事实：原5个check + `gen_proj_bg: hasText(projectBackground)`；聊天如实报告 **6项unknown**，不把派生结论写入模板保存态。若要求聊天必须知道项目背景，则需另行明确跨域读取范围，不能靠本次目录数据修改假装已知。
3. 合并项目名称采用 **derived + requiredWhen + completionSources**，省略 stateKey；保留不可写边界，无需修改构建脚本或通用契约。

本轮暂停实现，未自行采用上述修订。原任务书未改。

## 可复跑证据

`node src-ui/scripts/witness_selection_catalog_preflight.cjs`

见 [机器证据](./selection-result-catalog-preflight-evidence.json)。见证直接执行生产完成度函数及提取的生产构建guard，仅在内存比较“保留/删除 supplied”；记录7份相关源码SHA-256作为后续零代码判据基线。未运行完整build、Rust或真实桌面，因为尚未实施修改。
