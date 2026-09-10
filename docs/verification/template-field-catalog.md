# 字段目录泛化验证（2026-09-07）

## 结果

已实现立项签批表目录接入，并完成用户要求的第三模板数据扩展实验。需求表 8 个文本字段保持原范围；立项 4 个（IT/CT 服务内容、收入收款、支出付款）；仅增加目录数据后甄选 10 个（8 个 gen_zx_* 与两条收付款条款）由同一个读写工具处理。甄选页面、批次业务规则及会审纪要未扩展。

公共项目背景来自项目基本信息，列为 derived 只读；不跨存储边界写入。收付款条款声明原有根状态位置，避免写入 formData 影子值后在生成时被旧状态覆盖。没有修改数据库结构、保存事务、测算引擎、docfill、0 容差门禁或 presetFieldKeys。

## 严格数据扩展实验

1. 两模板基线运行同一通用 Rust 事务测试和插件测试通过。
2. 冻结 240 个源码文件哈希，覆盖 `.ts/.tsx/.rs` 与测试/构建脚本，包含生成的 TypeScript。
3. **唯一业务修改是 `src-ui/src/lib/templateCompletion/catalog.json` 新增甄选记录**。正常构建仅更新 JSON/编译产物，生成的 TypeScript 内容也未变化。
4. 同一个预先编写的通用测试遍历第三模板：读取空字段 → 审批展示旧值 → 自动修改后批准 → 原保存事务提交 → 再读取并比对批准值 → 共用完成度增加。10 个字段全部通过，根状态无 gen_* 影子副本。
5. 源码哈希比较通过，无任何代码变更。不是修改测试/业务代码后宣称数据驱动成功。

证据：[源码冻结与两版目录](./template-catalog-extension-evidence.json)、[实际 Rust 保存/读取结果](./template-catalog-transaction-evidence.json)。可复现命令见目录 [README](../../src-ui/src/lib/templateCompletion/README.md)。

## 自动检查

- `cargo test`：89 passed，20 ignored，0 failed。
- 使用应用已配置凭据运行 `cargo test agent_bridge -- --ignored`：20 passed，0 failed、0 skipped。新增通用真实模型用例自动遍历立项与甄选扩展：主动读字段 → 调用写工具 → 自动改后批准 → 保存结果是修改后的正文。测试工作区为合成数据，不写用户项目；凭据不进入仓库或证据。
- 前端 lint/build；dsh、template-state、session-scope、approval-review 回归通过。模板测试覆盖四类生产生成函数、立项完成度 +1、批准正文进入 IT/CT 变量、收付款进入原生成变量，以及根状态和表单字段的同步、同字段草稿冲突阻断、无关草稿保留。
- 插件 typecheck/build、scope、contract、template-read、template-catalog 通过。Rust 投影经生产插件 execute 计算的完成度与 UI 纯函数一致；需求表维持 11 项、8 个可写文本及原截断行为。
- 服务端拒绝所有目录非文本项（含真实 check key `public_fields_consistent`、`approval_amount_below_500k`）、中文校验标签、未知字段、财务字段、跨项目/通用会话、未审批和已消费授权。原版本冲突及审批期间变更回归通过。
- 三张排除模板无任何可写字段；路径模板名和歧义模板名拒绝。插件和 Rust 共享 v5 契约，开发/staging 插件均已构建同步。
- `npm run test:packaging`：6 passed。已有 Rust warning 和前端 bundle 体积提示仍存在，与本次业务结果无关。

[真实模型与批准保存证据](./template-catalog-real-evidence.json)。真人可读性判断不能由自动审批替代。

## 尚未验证及语义边界

- **本轮未完成真人 Gate B，也未执行新立项案例的桌面 DOCX 生成核对。** 已验证到实际保存态、生产生成变量与完成度；不能把这些称为真人全链路验收。
- 未保存的 IT/CT 动态默认值返回 `dynamic_default_unavailable`。项目背景和需外部方案/批次计算的 check 项返回 `evaluated=false`，计入 `unknownCount`，不冒充缺项或通过。有未知项时不能宣称整张表已经完成；甄选原业务生成校验保留。
- 数据目录随构建嵌入/打包，不支持已运行进程热更新。当前桌面应用与正式安装包须使用同代构建；本轮未更新用户正式安装应用。Windows NSIS 与既有发版真人 gate 仍待完成。

## 真人立项验收步骤

1. 在测试项目的立项签批表中把 IT 服务内容显式清空并保存，记录当前完成度。
2. 在绑定该项目的聊天中让 AI 先读取签批表，再仅补 IT 服务内容。
3. 审批中核对旧值确为空、目标项目和模板正确；修改一处正文后批准。
4. 回到模板页面确认完成度 +1，IT 内容为修改后的正文；切换标签或重新进入仍一致。
5. 通过原业务校验生成 DOCX，核对 IT 服务内容与人工修改后的文字一致。
6. 可另用收入侧收款条款验证根状态同步；生成后核对收款条款。项目背景应提示到项目基本信息修改，不能在模板写工具中直接修改。
