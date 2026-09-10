# 模板状态生成与聊天附件补齐验证（2026-09-06）

## 结论

路线图①代码及核心产品交互已完成。四模板跨页生成一致，聊天图片入库、双窗口同步、实际路径回执、docx 嵌图及数据库写入失败回滚均完成验证。
**真实模型口述缺项未验收**：本轮独立测试配置没有可用 API Key，模型请求明确失败；不能将前端确定性缺项卡片展示冒充模型回答成功。

## 环境与证据

- 独立 macOS debug 应用：`com.cmcc.benefitcalc.template-test`，未覆盖正式应用。
- 合成工作区：`/private/tmp/lamber-template-20260906-workspace`，由之前的3A合成工作区复制；模板复制自用户指定的最新 `../workspace/templates`。
- 仅合成数据库注入非默认表单和故障触发器；测试结束移除触发器并恢复测试项目文件夹绑定。未修改正式工作区数据。
- 本地 Word 留样：`/private/tmp/lamber-template-evidence/`；哈希及实际握手回包保留于 [evidence.json](./template-state-and-chat-assets-evidence.json)。临时文件可能被系统清理。

## 四模板跨标签生成

通过 Computer Use 点击真实产品生成按钮，读取产物 `word/document.xml` 比较：

| 模板 | 对照路径 | 结果 |
| --- | --- | --- |
| 需求导入表 | 需求信息 / 生成确认 | XML完全一致；自定义需求单位、服务内容等保留 |
| 立项签批表 | 签批信息 / 生成确认 | XML完全一致；自定义IT/CT内容及复选框保留 |
| 会审纪要 | basic / content / business / risk / confirm | 五份XML完全一致；跨页文本、技术明细、询价行保留 |
| 甄选结果签批表 | 甄选信息 / 生成确认 | XML完全一致；自定义中选人、甄选规则保留；使用真实甄选前/后合成方案通过原财务校验 |

独立自动化调用生产 `handleGenerate`（由 TypeScript AST 提取，不复制生成逻辑），覆盖四模板两次状态序列化往返；formRef访问直接抛错，以防重新依赖DOM。
另外注入生成输入清空、退回默认值两种故障，均在 invoke 前阻断。

## 图片闭环

1. 产品需求表初始9/11。打开聊天后，按正式模板上下文显示附件1/2两个确定性缺项卡片。
2. 经聊天卡片选择本地PNG存入附件1；核对SQLite的project_id/template_name/usage及磁盘文件：
   `3A/3A流式验收项目-图片/assets/asset_a72aea4193cb0681.png`。
3. 模板主窗即时从9/11到10/11，出现对应图片操作；聊天出现带完整实际路径的系统回执，附件1卡片消失。
4. 合成项目设为未绑定文件夹，附件2同入口上传，实际落于
   `.projects/id_79962ba708eb402dbb198e9d770d6ab3/assets/asset_35523ec9602ed6d.png`。
   回执准确报告此路径，完成度到11/11。
5. 从确认页生成需求表：ZIP含 `word/media/attach1_image_0.png` 和 `attach2_image_0.png`，正文有两个drawing节点。证明聊天与文档使用同一资产。
6. 在合成SQLite设置BEFORE INSERT拒绝触发器后，经聊天真实上传PNG：界面显示 `synthetic asset insert failure`，无新增成功回执；前后记录数与物理图片数均为1，证明文件回滚。随后移除触发器。
7. 自动化验证 >20MB、GIF/BMP、工作区切换、写库拒绝、固定目标/usage及fallback路径回执。界面也观察到不支持格式的拒绝提示。

## 模型能力

独立运行实际dsh ACP initialize，`deepseek-v4-flash` 返回 `promptCapabilities.image=false`；回包见evidence.json。
上传入库不依赖视觉模型。未更改用户模型/profile，未增加新审批或写工具。

## 检查

- 前端 lint/build 通过；保留原有bundle大小提示。
- `test:template-state` 通过：生产生成处理器、11项规则、图片缺失、复选框历史值、输入防呆、上传边界。
- `test:selection-batch`、`test:tax-split`、`test:ai-compute-quote`、`test:dsh` 通过。
- Rust常规测试：79 passed / 13 ignored / 0 failed（含新增需求表资产清理保护测试）。ignored模型集成测试本轮未执行，不冒充通过。
- 本轮产品GUI验证覆盖上传与生成主链；最后新增的孤儿清理保护及模板上传事件发布由自动化/构建验证，未重复整套GUI。

## 尚未覆盖

- 使用有效key的模型真实口述缺项及识图；本轮仅证实模型失败不会阻断确定性补图。
- 聊天卡片的系统剪贴板图片粘贴端到端；选择文件路径已实测。
- 原dsh发版gate、Windows安装验收及用户真人验收不在本次完成范围。
