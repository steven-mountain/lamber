# 常用资料与项目预设独立模块第一阶段

- **Status:** Done
- **Objective:** 新增 workspace 作用域的常用资料管理模块，并在典型表单字段中建立主动选择填充与主动保存为常用内容的底座。

## Completed

1. [x] **独立模块入口**
   - Hub 新增“常用资料与项目预设”入口。
   - 新增 `preset_center` 路由与 `PresetCenterView`。
   - 未打开工作区时复用 `WorkspaceGate`，不创建 localStorage 临时事实来源。

2. [x] **SQLite 持久化**
   - 新增 `common_presets` 表，schema version 升至 6。
   - 当前支持 `workspace` scope，预留 `user` scope 但不实现跨工作区同步。
   - 支持新增、编辑、启用/停用、软删除、使用次数与最近使用时间更新。

3. [x] **fieldKey 绑定机制**
   - 新增 `presetFieldKeys.ts`，使用稳定键而非中文标签绑定。
   - 初始键包括 `project_basic.customer_name`、`project_basic.background`、`project_basic.solution`、`approval.reviewers`、`approval.department`、`approval.project_manager`。

4. [x] **表单快捷填充**
   - 新增复用组件 `CommonPresetQuickFill`。
   - 支持选择常用内容填充、长文本替换/追加、保存当前字段为常用内容。
   - 使用常用内容后更新 `usageCount` 与 `lastUsedAt`。

5. [x] **首批字段接入**
   - ICT 基础信息：客户单位、项目背景。
   - 模板表单：项目背景、技术方案、参会人员、分公司/部门名称、风险/项目负责人。

6. [x] **表单字段覆盖扩展**
   - 项目需求导入表：项目需求单位、服务内容、客户确认、部署环境要求。
   - 会审纪要：驻点支撑人员、IT建设内容、CT建设内容、收入侧收款方式、支出侧付款方式、时间要求。
   - 立项签批表：IT服务内容、CT服务内容、收入侧收款方式、支出侧付款方式。
   - 收入侧收款方式和支出侧付款方式使用共享 fieldKey，便于会审纪要与立项签批表复用同一类常用资料。

7. [x] **Common preset action alignment repair**
   - Added `CommonPresetFieldHeader` so form labels and preset actions render in one responsive header row.
   - Replaced split `label` + separate `justify-end` preset action rows in ICT basic info and template forms.
   - The layout uses normal flex wrapping rather than fixed offsets or hardcoded positioning.
8. [x] **Preset and non-preset field header normalization**
   - Added `CommonPresetLabelHeader` for plain fields that need to align with preset-enabled fields.
   - Compact preset action buttons now avoid increasing normal label-row height.
   - Removed the detached sign-off payment preset band and attached payment presets to the actual payment input labels.

## Validation

- `npx tsc --noEmit` in `src-ui`: passed.
- `node scripts/test_common_presets.cjs` in `src-ui`: passed.
- Existing `scripts/test_subject_funding*.cjs` in `src-ui`: passed.
- `npm run build` in `src-ui`: passed; Vite reported the existing large chunk warning.
- `cargo fmt -- --check` in `src-tauri`: passed.
- `cargo test common_presets::tests` in `src-tauri`: passed.
- `cargo test benefit::calculator::tests` in `src-tauri`: passed.
- Full `cargo test` in `src-tauri`: blocked by existing missing docfill test template `项目全生命周期文件模版/效益分析表 .xlsx` (10 passed, 2 failed).
- Local Vite service at `http://localhost:5173`: HTTP 200.

## Remaining

- Desktop Tauri runtime manual click-through should be performed with a real workspace to verify CRUD, quick fill, persistence after reload, and usage-count updates end to end.
- Phase 2 remains out of scope: full project preset templates, project creation from presets, one-click multi-field application, intelligent recommendation, and AI collaboration.
