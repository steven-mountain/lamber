# 会审纪要目录与两张清单卡片验证

2026-09-08。按更新后的[任务书](../tasks/TASK_BOOK_meeting_review_and_list_fields.md)执行。此前按“必须零代码”暂停的核查已获用户批准修订，本轮完成通用契约及两张卡片开发，锁屏后桌面复测已补齐；整体验收仍有下述真人和旧版留样待项。

## 实现与边界

- 通用契约：`requiredWhen`兼容旧字符串开关并支持`{field, equals}`，Rust保留条件原始类型；`completionGroup`联合多个文本key计一个完成项；`completionSources`声明既有只读文本来源；`validRow`判有效列表行；`listType`与reason区分editable/generated。保留五种kind、未知事实和原三表语义。没有按meeting id分叉的目录完成度代码，没有用completionValues覆盖修补。
- 会审人工分类：18个可写文本key占17项，10个只读/复合项与2个清单项，共29项，顺序不变。新增数据阶段的[191份源码对比](./meeting-catalog-extension-evidence.json)排除了唯一的页面完成度替换；后续两张卡片改动单独实施。
- 技术清单：仅本轮明确技术清单编辑/建议意图挂卡片；用户采用AI表格后增删改、保存。界面共享状态与按模板独立持久化曾不一致，故卡片明确展示来源及目标，用户保存时通过原保存主体在一个事务更新两表；所有期望列表先校验，任何冲突整体回滚。不增加AI写工具。
- 询价：用户确认后调用原`autoGenerateInquiry`；只添加可选错误报告出口及成功/失败返回值，未复制金额推导。用户金额编辑仍调用原`handleInquiryAmountChange`，截图复用资产保存。实际结果经React状态提交和原串行保存后返回卡片，不提供随机值预览，不允许卡片增删询价行。重新生成时按原规则保留截图并明示核对证据。
- 请求沿聊天生成文档的主窗口加载协议，校验绑定、工作区、项目、模板和期望快照。失败保留错误；保存失败停止原目标自动保存，异步失败不污染切换后的目标。
- 本地卡片回执原未进入dsh持久会话，模型下一轮会看不到实际错误。本轮添加appReceipt来源标记、持久化和本会话最近8条按序只读上下文，覆盖所有现有卡片。普通用户/模型文本不按关键词冒充回执。

## 自动化结果

| 验证 | 结果与证据 |
| --- | --- |
| 原三表完成度回归 | 12组改动前快照，完成数字、缺项、顺序和evaluated逐项一致；fixture在`src-ui/scripts/fixtures/meeting-catalog-baseline.json` |
| 会审原29项 | 24组大型/非大型、采购方式、中台能力及空白/有效列表状态，与原页面29项逐项一致 |
| 页面与读工具 | 实际Rust投影37组送入同一catalog函数，逐项一致；含第105行才有效而显示只返回100行的场景 |
| unknown/条件/分组 | 缺失等值来源、缺少复合来源保留unknown；分公司两key任何一个空都未完成；旧布尔式不改 |
| 构建期校验 | 未知kind、非法条件、来源、清单子类和有效行配置被拒绝 |
| 双模板保存 | 实际事务成功；另一表旧快照冲突或损坏状态时整体回滚，版本不前进，其他字段/询价保留 |
| 清单交互逻辑 | 意图、建议解析、陈旧快照、数量校验、原生成器三条前置条件、实际报价范围及原封顶通过 |
| 请求/回执 | 生产保存回调失败隔离、应用回执进入下一轮渲染prompt、历史事件顺序与非回执排除通过 |
| Rust常规 | 91 passed，20 ignored，0 failed；ignored为需真实key的独立用例 |
| 真实key集成 | 全20项执行通过，0 skipped；包括跨项目拒绝/聚合、原文本写入及新增会审的真实读→审批修改→写入。审批修改由测试驱动，不冒充真人 |
| 前端 | lint、build、meeting-lists、template-state、chat-document、demand-upload、session-scope、approval-review、dsh通过 |
| 插件与打包 | typecheck、build、catalog-validation、template-catalog、template-read、scope、contract及6项打包测试通过；独立macOS测试app构建成功 |

主要复跑命令（仓库根目录）：

```sh
npm run test:meeting-lists --prefix src-ui
LAMBER_MEETING_PROJECTION=/tmp/lamber-meeting-projection.json cargo test --manifest-path src-tauri/Cargo.toml meeting_catalog_projection_matches_saved_predicate_inputs
LAMBER_MEETING_PROJECTION=/tmp/lamber-meeting-projection.json npm run test:meeting-lists --prefix src-ui
node agent-bridge/scripts/test-catalog-validation.mjs
python3 scripts/verify-catalog-delivery-witness.py
cargo test --manifest-path src-tauri/Cargo.toml
npm run lint --prefix src-ui
npm run build --prefix src-ui
npm run typecheck --prefix agent-bridge/dsh-tool-lamber
npm run test:packaging
```

真实集成沿现有`cargo test -- --ignored --test-threads=1`执行，已有配置密钥只注入测试进程环境，不写入仓库或日志。会审结果在[真实目录证据](./template-catalog-real-evidence.json)，原审批写入链路在[目录事务证据](./template-catalog-transaction-evidence.json)。本轮日志在`/tmp/lamber-meeting-real.log`、`/tmp/lamber-meeting-final-cargo.log`等临时文件。

## 独立第五张表

以新合成模板“交付交接记录（契约复验）”验证通用能力，使用独立字段、异地交付等值条件、责任人共同项、布尔旧写法、复合来源和交付任务有效行。临时仅增加catalog JSON条目，运行现有Rust与插件的读取→修改审批→保存→完成度链路；源码哈希均不变，随后恢复正式目录并重建插件。最终哈希数量以[证据JSON](./catalog-delivery-witness-evidence.json)为准（首次232份，回执模块加入后复跑233份）。

这是独立第五张合成表的契约验证，不是售前预算/效益表/决策纪要的实际业务接入，也不是带真实Word模板的交付功能。三个显式排除模板仍保持排除。

## 桌面与实际Word

环境为独立`Lamber Scope Test.app`（`com.cmcc.benefitcalc.scope-test`），合成项目“3A流式验收项目”，工作区`/private/tmp/lamber-template-20260906-workspace`。未覆盖正式安装app，也未通过脚本写真实项目库。

1. 从未打开模板页的聊天中，请真实模型提议两行技术方案可行性清单；卡片出现四列及“同时用于需求导入表与会审纪要”。采用建议、修改第一行名称为“卡片验收-网络接入”、增加“卡片验收-联调”行，再删除原第二行；保存后两张模板均见相同两行。需求11/11，会审29/29。
2. 需求表通过卡片保存和页面编辑到相同状态分别生成实际DOCX，逐文本节点、规范化XML表格/正文及按文档位置解析的图片顺序相同。两行技术清单也出现在实际会审Word中。
3. 会审五个Tab分别生成实际DOCX，108个文本节点、规范化正文/表格结构、图片顺序完全相同。
4. 询价真实模型只引导点击确认，没有预报厂商或报价。确认后实际产出厂商A/B/C三行，报价67800、74595、71543，税率6，均在卡片展示且每行有上传入口。用户改第一家为“合成验收厂商甲”、输入999999并上传合成截图，保存后金额为106000（收入上限），其余两行未改变。
5. 实际会审Word包含上述三家、封顶金额及同一截图SHA256；[DOCX证据](./meeting-list-docx-evidence.json)保留产物路径、文本/正文和图片哈希。临时产物在`/tmp/lamber-meeting-docx-evidence/`。
6. 将合成项目的临时收入改为63600，小于IT成本67800，卡片确认拒绝生成；原三行/截图保持不变，原错误完整回到聊天：“当前 IT 投入含税总成本为 67800.00，已超过含税总收入 63600.00，无法生成合规三家报价。”模型未给手工补三行等绕过方案，但当时下一轮看不到本地回执，要求粘贴错误。已按上文修复上下文，不将修复前回答算作理想通过。
7. 临时收入已在退出旧测试应用前通过UI恢复106000；仅用于拒绝测试，未保存为新的正式测算方案。修复版重建后曾被锁屏打断，随后用户要求继续，已完成下面两项复测。
8. **回执修复后的真实模型复测通过**：再次在UI把收入暂改为63600，点击卡片确认，原成本67800校验拒绝，原三行及截图仍在。仅发送“询价报错了，接下来怎么办？”，没有手工附上错误原文；真实模型直接引用最近回执的完整业务错误，并说明成本/收入前置条件及不得手工加行绕过。没有索要已有错误、虚构本次成功或提议具体询价厂商/金额/税率。测试后UI收入恢复106000，税率6，不含税100000。
9. **未绑定/通用会话桌面验证通过**：新会话的“未建立项目权限”选择页没有两张卡片；选择通用聊天后发送“请帮我提议技术方案可行性清单，并生成三家询价。”，生成中及完成后均无清单写卡片。模型明确当前无绑定项目，不能触发询价，没有调用项目读写工具或提议具体报价。明示通用聊天仅聚合只读。

本次接续证据：[meeting-list-desktop-followup.json](./meeting-list-desktop-followup.json)。上述操作由Computer Use执行，不是用户真人验收。

[生成源码证据](./meeting-generation-source-evidence.json)确认本轮前后`handleGenerate`、`toDocImagePayload`和原金额编辑函数提取后的代码哈希相同。没有改动docfill/表格克隆/财务公式/0容差门禁。

## 独立复核（2026-09-08，任务书作者，非执行方）

按任务书 0a–0f 逐条抽查源码，不以执行方报告为准。**结论：第〇件通过。**

已实际核对：

- **无模板分叉**：`catalog.ts` / `template_catalog.rs` / `template_read.rs` 检索无 `meeting`。
- **缺口 2 已关**：`as_bool` 仅余 `hasPublicUrl` / `hasSecurity` 两个真布尔根键；
  条件字段按原始类型投影（`template_read.rs:71-76`）。
- **契约白名单**：`kind` 仍五种；`validate-template-catalog.mjs` 拒非法等值条件、
  非法 `completionSources`、非法 `listType` / `validRow`，且 `completionGroup`
  成员不足 2 个即报错。
- **计数自洽**：meeting 30 字段 − 1 个共享组 = 29 项，与 18/17/10/2 的分类一致。

两处**优于任务书要求**、值得后来者知道的做法：

1. **插件的完成度函数是从 UI 源文件编译出来的**（`build-contract.mjs:37-39`），
   不是照着重写的第二份实现，且有哈希漂移测试。
   任务书只要求"共用同一个函数"，这是结构上不可能不一致。
2. **`validRow` 的谓词输入按全部行去重后投影**（`template_read.rs:102-121`），
   不是只投可见的 100 行 → **可见行截断改变不了完成度**。任务书未要求，执行方自行识别。

**第五张表判据：认定通过。** 用的是合成表而非 presales/benefit/decision，
但该判据的目的是"不得用逼出缺口的那张表自证"，合成表满足；
且它专门覆盖等值条件/共同项/旧布尔/复合来源/有效行五种新能力，针对性强于真实业务表。

## 复核发现的两个问题（执行方未提及）

### A. 🔴 甄选结果签批表仍是两套完成度口径（非本轮引入，但本轮后更易误读）

`TemplateForms.tsx:2223-2243` 的 `selectionResultCompletionItems` **仍是手写 14 项数组**，
未走 `getCatalogCompletion`；而 `catalog.json` 的 `selection` 有 **16 项**。
条目并不对应：页面有"合并项目名称"，目录有"合作内容描述 / 行业 / 标准方案"。

**实际后果**：AI 可写 `gen_zx_industry` / `gen_zx_std_plan` / `gen_zx_content_desc`，
**写入后页面完成度不变**；用户在页面看到的缺项与在聊天里问到的不是同一份。

这是字段目录泛化那轮就存在的已知限制（`CURRENT_TASK.md:101`
"甄选业务页面和会审纪要未扩展"），本轮只收了会审纪要。
但本轮之后"四张表都已统一"的错觉显著增强，**故升级为显式未完项**。

### B. `TemplateForms.tsx:2218` 的 `completionValues` 已成死代码

`gen_proj_bg` 已改为 `completionSources`，而 `catalog.ts` 中 completionSources
分支排在 `supplied` 之前 → 该覆盖永不生效。全仓仅此一处使用 `completionValues`。
留着会让后来者误以为它仍是机制。**建议清除，或注释说明其已失效。**

## 待完成与接续

- **真人会审文本验收**：绑定项目→让AI补一个空白会审文本→审批改一处→批准→页面完成度+1→生成Word核对。真实key自动测试覆盖会审审批修改及保存，不能代替用户真人。
- **真人共享提示确认**：已由CUA看见并验证两表完成度，仍需任务书要求的真人确认。
- **旧版五Tab产物基线**：本轮有当前五Tab实际Word一致、生成函数前后源码不变及原29项基线，没有旧二进制五Tab实际DOCX留样。因此不把这一组合证据写成“旧版与新版五Tab实际产物逐一比对完成”。
- **发版**：沿既有Gate A/B和Windows NSIS验收清单，macOS开发构建不代表Windows验收通过。
- ✅ **甄选结果签批表页面接入目录**（复核发现 A）：已于 2026-09-10 立项，
  见[任务书](../tasks/TASK_BOOK_selection_result_page_catalog.md)。
  调研补充：该表 6 个 check 项中 5 项依赖运行期派生数据，
  **聊天侧只能报 unknown**——这是设计上的正确行为，不是遗留缺陷。
- **清除 `TemplateForms.tsx:2218` 的死 `completionValues`**（复核发现 B）。
