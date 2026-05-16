# Lamber 效益分析工具集 - 项目状态报告 (PROJECT_STATE)

## 1. 项目概览
本项是一个基于 **Tauri + React + Rust** 构建的售前工具集，旨在通过自动化手段提升项目效益测算、申报材料制作以及 5G/ICT 项目全生命周期管理的效率。

- **当前分支**: `feature/ai-integration`
- **当前版本**: v1.1.0 (AI Copilot 集成版)
- **核心架构**: 
  - 前端: Vite + React 18 + TailwindCSS + Shadcn/UI
  - 后端: Rust + Tauri (使用 native HTTP 插件规避 CORS)
  - AI 层: 支持 SSE 流式传输，兼容 Ollama 及 OpenAI 标准接口

---

## 2. 已完成功能 (Milestones)

### A. 效益测算模块 (Benefit Tool)
- [x] **单项目测算**: 支持 IT 集成与 CT 产品分项输入，自动计算毛利率、NPV 等。
- [x] **批量测算**: 支持 Excel 模板导入，一键生成批量计算报告。
- [x] **可视化看板**: 采用玻璃拟态设计的计算结果展示区域。

### B. 材料制作模块 (Docfill Tool)
- [x] **模板引擎**: 实现基于 Rust 的 Word 文档变量填充。
- [x] **表单系统**: 结构化的项目基本信息输入界面。

### C. Lamber 智能顾问 (AI Copilot) - NEW!
- [x] **流式输出控制**: 解决 SSE 重复拼接 Bug，实现平滑打字机效果。
- [x] **上下文感知**: 自动抓取当前页面的测算数据（如毛利率、NPV）进行分析。
- [x] **混合知识体系**: 建立内置产品库，实现 `[系统内置]` 与 `【系统外扩展】` 标签的差异化渲染。
- [x] **配置持久化**: AI 设置（端点、模型、API Key）自动保存至 `localStorage`。

---

## 3. 待开发计划 (Roadmap)

### 第一阶段：AI 深度集成 (Phase 2 补全)
- [ ] **多轮对话上下文优化**: 优化历史消息清理逻辑，防止 Context Window 溢出。
- [ ] **本地知识库扩展**: 允许用户上传本地 PDF/Markdown 文件作为 AI 的私有知识增强 (RAG)。

### 第二阶段：业务功能扩展 (Business Logic)
- [ ] **项目协同功能**: 实现项目测算结果的导出与分享（JSON/PDF 格式）。
- [ ] **ICT 生命周期管理**: 完善 `IctLifecycle` 视图，支持项目关键节点的时间轴追踪。

### 第三阶段：工程优化 (Engineering)
- [ ] **安装包体积优化**: 移除开发期冗余依赖。
- [ ] **多语言支持**: 引入 i18next 实现中英文切换。

---

## 4. 关键技术细节
- **AI 接口地址**: 默认 `http://localhost:11434/v1/chat/completions` (Ollama)。
- **持久化配置**: AI 配置存放在浏览器 `localStorage` 中。
- **构建指令**: `npm run tauri dev` (开发) / `npm run tauri build` (打包)。

---
*最后更新时间: 2026-05-16*
