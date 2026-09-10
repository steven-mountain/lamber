# 会审纪要目录接入：开工核查

2026-09-08。状态：完成只读核查，触发任务书关键事实 6 的停工汇报条件；业务代码和目录尚未修改，未执行产品或 DOCX 验收。

## 发现的契约缺口

任务书要求第一件除页面完成度替换外零代码改动。但当前目录不能完整表达会审纪要的既有完成度语义：

1. `src-ui/src/lib/templateCompletion/catalog.ts` 的 `requiredWhen` 是根状态键名，仅执行 `Boolean(state[key])`。`projectScale` 的 `large` 和其他非空值都会成为 true；`procurementMethod` 的“其他”和“短名单甄选”同样如此。
2. `src-tauri/src/agent_bridge/template_read.rs` 仅用 `Value::as_bool` 投影条件，因此以上字符串经读工具均成为 false。单在页面替换处补临时布尔值不能修复聊天读取，会产生两个完成度口径。
3. 原“分公司参会人员”是 `gen_branch_name` 与 `gen_branch_attendees` 的合取；两个输入共占一个完成度项。当前算法一个字段产生一个计数项，无法同时开放两个文本字段、保持总数 29 并完整检查这两个值。
4. 原询价项检查“至少一行有厂商名或正金额”；目录 list 只检查数组非空。空白行在两条路径上会产生不同结果。

`completionValues` 可供页面传递外部事实，但读工具不会生成这些值。用它覆盖上述所有问题只会让页面看起来通过，不能作为目录泛化成功的证据。

清单子类是第四件单独授权的契约扩展：Rust `Field` 启用了 `deny_unknown_fields`，不能直接加任意 `subtype` 属性；也不能直接把 kind 改为 `list.editable`，构建脚本只接受五种现有 kind。建议保留 `kind:list`，增设经过校验的子类元数据，不改变文本工具的可写判断。

## 原 29 项人工分类

以下保留原计数顺序。只读指本次不开放 AI 文本写入，不意味着原页面禁止用户编辑。

| # | 完成度项 | 字段或状态 | 本次分类与原判断 |
| --- | --- | --- | --- |
| 1 | 会议开始日期 | gen_meet_start | 文本；默认当天，动态默认不能在读工具里伪造 |
| 2 | 会议结束日期 | gen_meet_end | 文本；默认当天，同上 |
| 3 | 会议方式 | gen_meet_mode | 文本；默认“线上” |
| 4 | 项目规模 | projectScale | 选项，只读；非空 |
| 5 | 市公司参会人员 | gen_city_attendees | 文本；projectScale 等于 large 时必填 |
| 6 | 分公司参会人员 | gen_branch_name + gen_branch_attendees | 两个文本，共一个完成度项；两者均非空，名称默认 XXXX |
| 7 | 驻点支撑人员 | gen_onsite_support | 文本；非空 |
| 8 | 项目背景 | projectBackground | 项目来源，只读；非空 |
| 9 | IT建设内容 | itContent | 既有根状态/业务来源，只读；非空 |
| 10 | CT建设内容 | ctContent | 既有根状态/业务来源，只读；非空 |
| 11 | 技术方案 | gen_tech_solution | 文本；沿用原静态默认值 |
| 12 | 技术方案可行性清单 | techItems | 用户编辑清单；行数大于零，amount 是数量 |
| 13 | 涉及中台能力调用 | hasMidThree + midThreeCode + midThreeName | 条件与复合检查，只读；未涉及，或编号和名称均非空 |
| 14 | 自主三问 | selfThreeValue | 固定选项，只读；非空 |
| 15 | 三化方案 | gen_threeization | 文本；沿用原静态默认值 |
| 16 | 战略价值 | gen_strategic_value | 文本；非空 |
| 17 | 综论 | gen_tech_conclusion | 文本；沿用原静态默认值 |
| 18 | 收入付款方式 | revCollection | 按本任务书只读；原根状态非空 |
| 19 | 支出付款方式 | expPayment | 按本任务书只读；原根状态非空 |
| 20 | IT服务模式/商务模式 | itBusMode | 既有根状态选项，只读；非空 |
| 21 | 资金来源 | itFundSrc | 既有根状态选项，只读；非空 |
| 22 | 询价情况/询价过程 | inqVendors | 既有生成器清单；存在厂商名非空或金额大于零的行 |
| 23 | 时间要求 | gen_construction_time_req | 文本；沿用原静态默认值 |
| 24 | 风险点及其他责任人 | gen_risk_owner | 文本；沿用原静态默认值 |
| 25 | 是否联合体投标 | gen_is_joint | 文本；默认“否” |
| 26 | 项目评审清单准确完整 | gen_review_acc | 人工表述文本；不是财务自动校验结论，沿用原默认值 |
| 27 | 是否涉及单一来源 | gen_single_source | 文本；hasSingleSource 为 true 时必填，沿用原默认值 |
| 28 | 采购方式 | procurementMethod + gen_procurement_method_other | 选择项只读，“其他”时要求说明文本非空 |
| 29 | 售中建设及施工界面 | gen_construction_interface | 文本；沿用原静态默认值 |

实际为 18 个可写文本 key（“分公司参会人员”包含两个），占 17 个完成度项；另有 10 个只读/复合项、2 个清单项。任务书的“约 17”不能直接作为 schema 字段数。

## 建议的受控修订（待用户决定）

允许第一件补齐通用目录契约：声明式等值条件、多个字段共用完成度项、清单有效行检查；Rust 只投影规则需要的安全状态，UI/插件继续复用唯一完成度函数。缺失外部事实保持 unknown，不制造默认通过。字段目录与完成度计数可有不同数量，保持原页面 29 项及顺序。

预计涉及 `catalog.ts`、`template_catalog.rs`、`template_read.rs`、构建校验及对应测试；不以 meeting 专用分支实现。清单子类元数据随第四件扩展；两张卡片继续各自实现，询价复用原生成器和金额编辑路径。

获准后先完成第一件的独立数据/源码差异及 29 项回归证据，再实施两张卡片。旧“零代码判据”应改记为“契约补齐 + 数据接入”，不得继续宣称纯数据接入；本次尚未执行任何测试，真人与发版 gate 均未改变。
