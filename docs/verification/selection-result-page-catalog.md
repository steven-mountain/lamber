# 甄选结果签批表页面接入目录验收（2026-09-10）

按修订后的任务书完成本轮工程验证。四张模板页面现共用目录完成度函数；甄选页面原14项移入目录，增加合作内容描述、行业、标准方案，共17项。原14/14可能因合作内容描述未填显示16/17，属于补齐原漏计数；生成门禁、默认生成内容、批次规则及财务计算未改。本轮独立于D。

## 修复范围

根因是甄选页面保留手写清单，与AI读写所用目录分离，三个可写且已渲染字段未参与页面计数。只在目录修补字段而保留手写清单仍会分叉，因此改为同一个纯函数。原14项相对顺序不变；三个新项分别接在中选合作伙伴、甄选范围、甄选规则后。

页面逐字保留6个外部事实判定式：项目背景、甄选后方案、公共字段一致、批次差异确认、续签归类、立项金额门槛。聊天保存态无法判定时为unknown，不能伪造已完成或未完成。合并名称为只读derived，使用completionSources与requiredWhen，不带stateKey；batch非空才完成，single不要求。供应商是否中小企业使用已有completionAlways保留原页面始终完成语义（否/是均有效，旧异常空值也不新增门禁）。

立项现有背景completionValues保留。其他三张模板目录记录和页面调用块逐字不变。除清单替换外，TemplateForms仅移除不再使用的类型导入。未改通用算法、Rust目录/读取生产代码、构建校验、批次逻辑或handleGenerate。

## 自动对照

- [冻结基准对照](./selection-page-catalog-evidence.json)：直接执行改前14项页面代码与改后页面调用，共2048组single/batch及状态组合，逐项判定和相对顺序全部相同；新增三项逐个空/非空影响计数；立项明确7/7。
- [真实保存/读取投影](./selection-page-projection-evidence.json)：Rust生产保存及读取路径，共四表×两模式×空/已填16组；没有将运行期判定写入保存态。生产写入能力仍不含合并名称。
- [实际读工具执行](./selection-page-read-tool-evidence.json)：上述16组投影送入构建后的read_template_fields，仅网络返回用桩；同证据输入下与UI共同函数逐项一致，其他三表与旧目录逐项一致，selection的6个unknown不进入missingFields。
- 5个受保护源文件SHA-256与改前一致：catalog.ts、template_catalog.rs、template_read.rs、build-contract.mjs、validate-template-catalog.mjs；值详见第一份证据。
- 既有会审清单测试仅调整其selection旧16项基准到新17项映射，其他三表原数组基准保留。改前暂停见证改为读取冻结目录，不覆盖历史证据。

通过：前端lint/build；插件build/typecheck；Rust常规101通过、23个显式ignored未运行；甄选批次、模板状态、会审清单、聊天文档、会话绑定回归；原暂停见证复跑。独立macOS测试应用已重建。

复跑入口：

```sh
npm run test:selection-page-catalog --prefix src-ui
cargo test --manifest-path src-tauri/Cargo.toml selection_page_catalog_saved_projection_regression -- --nocapture
node agent-bridge/scripts/test-selection-page-catalog.mjs
```

## 真实桌面与模型

[结构化桌面证据](./selection-page-desktop-evidence.json)。在重建的独立Scope Test中打开合成项目“桥接测试项目-28科目桌面对照”的实际模板页面：

| 操作 | 行业值 | 页面完成度 |
| --- | --- | --- |
| 原状态 | / | 13/17 |
| 页面清空并自动保存 | 空 | 12/17 |
| 真实模型读取后发起审批；核对仅行业一个字段并批准 | 软件和信息技术服务业 | 13/17 |
| 页面恢复原值 | / | 13/17 |

AI写入通过原审核对话框和真实fill_template_fields，数据库审计确认仅gen_zx_industry。模板内容刷新后，未重启页面即显示新值与13/17。再次真实读取，模型逐项列出2个缺项（合作伙伴、内容描述）、6个unknown，并明确：“上述6项为unknown，不代表已完成或未完成，需在甄选结果签批表页面查看判定结果。”行业显示已保存。聊天9个完成与页面13个完成的差值来自6个运行期事实，未声称两侧掌握相同数据。

验收使用原生无障碍树读取实际UI及真实模型响应，未声称像素截图验收或真实业务项目验收。进入模板时合成旧基准有三处既有0.01元税额尾差，通过原“忽略微小尾差”确认进入，未保存财务编辑器。退出后只读核对甄选前快照仍v11（11条）、甄选后仍v1（1条）；模板行业恢复为/。未生成文件，正式项目未写。测试应用退出，临时AI配置恢复并核对。未提交、未发布；既有真人及Windows发版gate不在本轮结论内。
