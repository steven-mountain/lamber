# 模板字段目录

`catalog.json` 是唯一数据源，按模板分组放在一个文件中。选择单文件是为了 UI、Rust 的 `include_str!` 和插件构建都直接消费同一份目录；新增模板只增加一个对象，不需要维护 Rust 文件列表或插件 import 清单。它使用实际表单 key，与 `presetFieldKeys.ts` 的业务语义 key 不合并。

## 目录契约

- 模板：`id`（稳定标识）、`name`（模板族名称/读取别名）、`suffix`、`fields`。完整名称按包含族名且后缀一致匹配；路径和歧义拒绝。别名读取只在绑定项目恰有一个保存候选时解析，写入必须完整名称。
- 字段：`key`、`label`、`kind`。只有 `text` 可写；`derived` 是外部项目/测算派生数据，`check` 是业务校验结论，`list` 是动态行，`image` 是资产槽位。
- 文本默认存储于 `formData[key]`；`stateKey` 显式声明已有模板根状态位置。收付款条款分别使用 `revCollection` / `expPayment`，不写无效的 `gen_*` 影子值。不得将项目基本信息误标为模板文本。
- `defaultValue` 是静态界面默认值；显式空字符串优先。`dynamicDefault` 表示默认值需外部业务状态才能计算，读工具不伪造该值，返回 `dynamic_default_unavailable`。
- `requiredWhen` 兼容原字符串布尔开关，并支持 `{field, equals}` 显式比较字符串/布尔/数字；等值条件来源缺失记未知。`list.columns` 只列可读取的行字段；图片仅返回安全资产存在标志。动态行、图片始终不可由 AI 工具写入。
- `completionAlways` 用于既有完成度中不因选择值改变而缺失的选项：原布尔开关及甄选“供应商是否中小企业”（原页面始终完成，否/是均有效）。不得用它绕过财务或批次校验。
- `excludedReason` 与空 `fields` 显式排除售前预算表、效益分析表和立项决策纪要：零占位符，无人工文本字段。会审纪要已接入18个文本key；PPT尚未接入。

`gen_proj_bg` 来自项目基本信息，不属于模板保存态，本轮明确标为派生只读。甄选目录中的校验项已人工分类，不能把完成度清单整体当成可写字段。

## 生成与完成度

`catalog.ts` 是唯一目录完成度算法；需求表兼容入口 `demand.ts` 与立项、会审、甄选页面都调用它。插件构建直接编译这个源文件，绝不重写算法。外部派生值/业务校验无法从模板保存态判定时返回 `evaluated=false`；这些项不进入 `missingFields`，单独计入 `unknownCount`。目录检查不替代原业务生成门禁，也不能把有 unknown 的结果称为整表完成。

立项页面向纯函数传入其既有项目背景判定和动态 IT/CT 默认值。主动读取仅使用保存态，因此未保存的动态默认值与项目背景可能未知；不会为了读取重算财务或扩大被动注入范围。甄选页面已接入目录，共17项；页面提供背景与5个业务校验事实，聊天侧这6项保持unknown。原14项判定与相对顺序不变，新增合作内容描述、行业、标准方案。合并名称以derived + completionSources + requiredWhen读取根状态，batch必填，single不要求，不开放AI写入。

`agent-bridge/scripts/build-contract.mjs` 遍历全部模板，生成按模板分组的文本 schema JSON 和目录 JSON，复制到插件 `lib/`；生成的 TypeScript 适配器保持不变。插件 schema 使用文本字段并集，执行时再次按目标模板校验；Rust 在审批前及实际写入前独立按目标目录校验，非文本、其他模板字段、非字符串和财务字段调用均拒绝。

新增目录后须正常构建插件、前端和 Rust，并同步分发资源；这里的“只改数据”不等于已运行二进制热加载。桥接契约 v5 防止旧需求表插件与新读写接口混用。

## 扩展验证

先保留需求表与立项两组，运行 `node scripts/verify-template-catalog-extension.mjs snapshot`；然后仅加入甄选目录并构建，运行 Rust `template_catalog_all` 测试、插件 `test:template-catalog` 与上述脚本 `verify`。脚本比较全部业务源码和测试/构建脚本哈希，包含生成的 `.ts`，扩展时任何源码变化都会失败。

实际验证与真人待办见 [template-field-catalog.md](../../../../docs/verification/template-field-catalog.md)。

## 通用条件、共同完成项与清单子类（2026-09-08）

`completionGroup` 让多个text key共同占一个完成度项，要求各字段均完成，第一项决定显示名称与顺序。`completionSources` 声明根状态或formData中的既有文本来源及静态默认值，全部非空才完成；不执行派生业务计算。来源缺失保持未知。会审30个目录条目对应29个完成项，不应按可写key数量计算完成度。

`validRow: {nonEmpty, positive}` 声明至少一行满足任一非空文本或正数判据。无该配置仍沿用数组非空语义。Rust仅投影存在/正负证据；显示行截断到100行，但完成度证据取全部行的判据组合并去重，不因第101行以后才出现有效行而漏判。

`kind` 仍只接受既有五种。`listType: editable | generated` 配合必需的reason说明边界；techItems可由AI提议、用户保存，inqVendors只能由用户确认触发既有生成器，AI不参与行内容。构建校验和Rust严格类型同步，不新增meeting特例；不得借completionValues修补两端语义。

真实项目外部背景无法从模板保存态取得时，读工具保留unknown；页面掌握该事实时可以完成，两端是在同一证据输入下逐项相同，不伪造数据来源相同。

独立第五张交付交接记录用临时JSON条目验证读→修改审批→保存→同函数完成度，232份源码哈希不变，验证后恢复正式目录。[本轮证据与限制](../../../../docs/verification/meeting-review-and-list-fields.md)。
