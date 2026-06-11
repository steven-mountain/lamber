# PROJECT_INDEX.md

## Common Materials & Project Presets

- **Module doc:** [docs/modules/common-presets.md](./modules/common-presets.md)
- **Frontend page:** [PresetCenterView.tsx](../src-ui/src/views/PresetCenterView.tsx)
- **FieldKey catalog:** [presetFieldKeys.ts](../src-ui/src/lib/presetFieldKeys.ts)
- **Quick-fill component:** [CommonPresetQuickFill.tsx](../src-ui/src/components/common-presets/CommonPresetQuickFill.tsx)
- **Backend commands:** [common_presets.rs](../src-tauri/src/common_presets.rs)
- **Persistence:** workspace SQLite table `common_presets`

## 项目定位
Lamber 是一个基于 **Tauri + React + Rust** 架构的销售支撑桌面工具，旨在规范 5G/ICT 项目全生命周期管理，提供自动化经济效益测算（NPV、利润率、现金流），并支持使用结构化模板生成并填充招投标及项目评审文档（Word/Excel）。

## 核心业务模块与主要入口
1. **项目看板 (Project Board)**
   * 前端：[ProjectBoard.tsx](file:///Users/hermesjang/Documents/CMCC/tools/lamber/src-ui/src/views/ProjectBoard.tsx)
   * 后端：[project_files/service.rs](file:///Users/hermesjang/Documents/CMCC/tools/lamber/src-tauri/src/project_files/service.rs)
2. **ICT 生命周期测算 (ICT Lifecycle Calculator)**
   * 前端：[IctLifecycle.tsx](file:///Users/hermesjang/Documents/CMCC/tools/lamber/src-ui/src/views/IctLifecycle.tsx)
   * 测算引擎：[calculator.rs](file:///Users/hermesjang/Documents/CMCC/tools/lamber/src-tauri/src/benefit/calculator.rs)
   * 业务管理：[service.rs](file:///Users/hermesjang/Documents/CMCC/tools/lamber/src-tauri/src/benefit/service.rs)
3. **文档与模板填充 (Template Word/Excel Engine)**
   * 前端：[TemplateForms.tsx](file:///Users/hermesjang/Documents/CMCC/tools/lamber/src-ui/src/views/TemplateForms.tsx)
   * 后端：[docfill.rs](file:///Users/hermesjang/Documents/CMCC/tools/lamber/src-tauri/src/docfill.rs)
4. **数据管理与文件扫描 (Data Management & File Scan)**
   * 前端：[DataManagement.tsx](file:///Users/hermesjang/Documents/CMCC/tools/lamber/src-ui/src/views/DataManagement.tsx)
   * 后端扫描：[scanner.rs](file:///Users/hermesjang/Documents/CMCC/tools/lamber/src-tauri/src/project_files/scanner.rs)
   * 工作空间维护：[workspace_maintenance.rs](file:///Users/hermesjang/Documents/CMCC/tools/lamber/src-tauri/src/workspace_maintenance.rs)
5. **AI 顾问 (AI Assistant)**
   * 前端：[AiChatPanel.tsx](file:///Users/hermesjang/Documents/CMCC/tools/lamber/src-ui/src/components/ai/AiChatPanel.tsx)
   * 上下文构建：[buildAiChatContext.ts](file:///Users/hermesjang/Documents/CMCC/tools/lamber/src-ui/src/ai/context/buildAiChatContext.ts)
6. **智算报价测算 (AI Compute Quote)**
   * 前端：[AiComputeQuoteView.tsx](../src-ui/src/features/ai-compute-quote/AiComputeQuoteView.tsx)
   * 计算核心：[calculations.ts](../src-ui/src/features/ai-compute-quote/calculations.ts)
   * 模块文档：[ai-compute-quote.md](./modules/ai-compute-quote.md)

## 模块详细设计上下文目录 (On-Demand)
* **前端外观与设置体系 (Appearance)**: [appearance.md](file:///Users/hermesjang/Documents/CMCC/tools/lamber/docs/modules/appearance.md)
* **采购甄选费测算 (Selection Fee)**: [selection-fee.md](file:///Users/hermesjang/Documents/CMCC/tools/lamber/docs/modules/selection-fee.md)
* **智算报价测算 (AI Compute Quote)**: [ai-compute-quote.md](./modules/ai-compute-quote.md)

## AI 上下文读取规则
在每轮开发任务开始前，AI 必须严格执行按需读取，以节省上下文空间：
1. **每次任务必读（轻量入口）：**
   * [PROJECT_INDEX.md](file:///Users/hermesjang/Documents/CMCC/tools/lamber/docs/PROJECT_INDEX.md) （本文件）
   * [CURRENT_TASK.md](file:///Users/hermesjang/Documents/CMCC/tools/lamber/docs/CURRENT_TASK.md) （当前/下一步任务记录）
2. **涉及具体模块开发时按需读取：**
   * 外观/主题/色彩/样式修改：读取 [appearance.md](file:///Users/hermesjang/Documents/CMCC/tools/lamber/docs/modules/appearance.md) 与 [DESIGN.md](file:///Users/hermesjang/Documents/CMCC/tools/lamber/DESIGN.md)
3. **涉及全局架构/启动流程/核心交互时按需读取：**
   * 读取 [ARCHITECTURE_MAP.md](file:///Users/hermesjang/Documents/CMCC/tools/lamber/docs/ARCHITECTURE_MAP.md)
4. **追溯历史或回归分析时按需读取：**
   * 读取 [CHANGELOG_AI.md](file:///Users/hermesjang/Documents/CMCC/tools/lamber/docs/CHANGELOG_AI.md) （默认不读取）
   * 读取 [PROJECT_STATUS.md](file:///Users/hermesjang/Documents/CMCC/tools/lamber/docs/PROJECT_STATUS.md) （历史/综合兼容，默认不读取）
   * 读取 [AI_CONTEXT.md](file:///Users/hermesjang/Documents/CMCC/tools/lamber/docs/AI_CONTEXT.md) （历史/综合兼容，默认不读取）
