# dsh 阶段 3B：默认切换与临时回退

日期：2026-09-05。第 4 步代码完成；第 5 步“实际使用稳定后下线”尚未执行。
本记录不代表 3B 整体验收、Gate 1A 或发版通过。

## 工作区核对

- 代码：`.`，原 `master` 工作目录；保留前几阶段未提交改动。
- 用户本轮明确指定的最新模板工作区：`../workspace`。
- 只读检查应用配置：`modulePaths.ict_lifecycle` 已指向该目录，无需更改。
- 目录内已确认需求导入表、会审纪要、立项签批表、甄选结果签批表、决策纪要、预算/效益 Excel 及决策汇报 PPT 模板。
- 该路径是模板模块目录，和承载 `.lamber.sqlite` 的业务数据工作区是两个配置；没有将模板目录自动初始化为业务数据库。

## 实现及接替

- `AiChatPanel` 通过 `chatTransition.usesDsh` 默认选择 dsh；旧 `AiRuntime` 只在显式回退发送时创建。
- 模型设置提供“临时回退并新建会话”和“返回新版并新建会话”。生成中禁止切换；当前输入和附件保留。
- 回退会话 id 只存在当前窗口内存中。其他新建会话及重开窗口使用 dsh，不存在持久化的全局旧链路默认值。
- 切回已有 dsh 会话依旧通过持久化映射恢复。无映射且含用户消息的旧历史显示过渡提示，下一次发送创建独立 dsh 会话；旧记录不删除、不自动上传或重放。
- 设置依旧由 `AiAgentSettingsCard` 管理。服务地址、模型和密钥不从旧链路自动迁移。
- 文本/快捷动作、业务上下文、上传图片/模板附件、流式、取消和恢复仍沿用 [3A 接替表](./dsh-stage3a-streaming-and-handover.md)。未修改财务公式、文档引擎、审批策略或工具集合。

## 自动化验证

| 检查 | 实际结果 |
| --- | --- |
| `npm run test:dsh --prefix src-ui` | 通过；含默认 dsh、回退会话隔离、重新打开、旧历史保留、ACP 映射恢复，以及原适配器回归 |
| `npm run lint --prefix src-ui` | 通过 |
| `npm run build --prefix src-ui` | 通过；既有 bundle 大小告警仍在 |
| `cargo test --manifest-path src-tauri/Cargo.toml` | 78 passed，13 ignored，0 failed；既有编译告警仍在 |
| `cargo test … agent_bridge -- --ignored --test-threads=1` | 13 passed，0 ignored，0 failed；真实 key 从已有本机设置读入测试进程环境，无缺 key 跳过 |
| `npm run test:dsh-stream` | 真实 dsh + 受控 SSE：6 次提交前更新，最终 ACP 一致 |
| `verify-dsh-stage3a.py --real --app-config …` | 640 个增量、637 次提交前显示；首增量 1.693 秒、首 ACP 提交 7.711 秒，最终正文/思考一致 |
| `npm run typecheck --prefix agent-bridge/dsh-tool-lamber` | 通过 |
| `npm run test:packaging` | 5 passed |

真实模型回归覆盖取消续聊、恢复记忆及 cwd 拒绝、视觉识图、测算工具和审批。
自动化测试使用合成数据；不以它冒充指定模板目录的真实界面生成验证。

## 产品界面验收（Computer Use）

构建：独立 macOS debug `.app`，bundle id `com.cmcc.benefitcalc.stage3b`，未覆盖安装包。
合成数据库沿用 `/private/tmp/lamber-stage3a-ui-workspace`；模板模块目录使用用户指定的真实目录。
测试包补入原生 macOS Node 并刷新本轮插件，重新 ad-hoc 签名；这不代表 Windows 安装验收。

| 操作 | 实际结果 |
| --- | --- |
| 首次打开 AI / 模型设置 | 显示“已启用新版 AI 对话”，直接使用 dsh 设置卡，密钥仅显示已保存 |
| 默认发送中文合成提示 | 收到“3B默认链路正常。”，不需要手动打开试用开关 |
| 生成中停止再续聊 | 输入恢复，收到“停止后续聊正常。” |
| 临时回退 | 会话数量增加，旧会话保留，旧设置独立；请求命中本地受控 SSE 服务，未携带 dsh 助手历史 |
| 回退快速完整响应 | 初测暴露旧解析器异步末块丢失；修复后显示“临时回退测试正常。” |
| 重开窗口 | 自动恢复 dsh；当前旧会话显示独立过渡提示 |
| 旧历史过渡发送 | 从 3 个会话变为 4 个，原记录仍可见，真实收到“历史过渡正常。”；生成中不出现过渡提示 |
| 最新模板目录及页面 | macOS 授权后，生命周期工作区按钮 Help 为 `../workspace`；列表包含 8 个顶层模板，需求导入表页面与附件控件正常加载 |
| 实际运行文件 / cwd | Node PID 73173 为测试 `.app/Contents/Resources/agent-runtime/node`；2 条 ACP 映射 cwd 都是合成数据工作区，未回退代码目录 |

回退末块问题的原因是 `finalize()` 只调度 React 状态，旧的零延时清理可能先清掉目标会话 id。
现在解析器同步返回最终快照，成功/取消时先写入发起会话，再释放轮次；移除零延时清理。
新增不提交任何 React render 的回归，验证瞬时完整响应和取消前缓冲仍可保存。
默认首轮绑定未完成时还会短暂满足旧历史判定，已让过渡提示在该会话生成中隐藏。

界面工具第一次误定位到旧安装版，其旧配置 schema 在保存时移除了主配置的 `aiAgent`。
已从本机既有 `com.cmcc.benefitcalc.gate` 配置恢复 AI 凭据及 Flash/官方端点；
随后独立测试包成功使用该本机配置完成真实问答。模板模块路径未更改。

## 尚未完成

- 模板目录权限等待已解除，指定目录和模板页面显示核对通过。此次没有再次生成文件或上传模板图片；既有跨 Tab 字段回退缺陷仍见独立任务书。
- Gate 1A 保存 key 全链路、安装缺件文案与无系统 Node 的完整安装证据仍未补齐。已有自动化对应测试通过不等于产品 gate 全部通过。
- 用户真实稳定使用结论尚未取得，故保留 `AiRuntime.ts`、旧 endpoint/model/apiKey UI 和 `useStreamingParser` 供临时回退，不宣称完成下线。
- Windows NSIS Gate 1B、真人审批四路径及用户真人验收仍需发版前完成。
- 模板跨 Tab 生成的既有字段回退缺陷仍按独立任务书待修。

后续删除旧实现时必须同时移除临时回退入口，并按 3A 接替表逐条核验；不能留下可点击但无实现的回退选项。

## 环境问题补记

首次访问指定模板目录时，独立测试 App 主线程卡在同步 `get_available_templates → read_dir/open`。
TCC 日志确认为 `kTCCServiceSystemPolicyDocumentsFolder` 的 `AUTHREQ_PROMPTING`；
系统授权完成后页面正常显示，无需改模板路径或文档引擎。
