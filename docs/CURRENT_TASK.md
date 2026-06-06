# 常用资料与项目预设阶段 1.5 收尾修复

- **Status:** Done
- **Objective:** 提升字段预设关闭入口的可发现性，并修正“选择常用”的图标语义。

## Completed

1. [x] **字段预设关闭入口**
   - 已启用字段的 label 工具区新增紧凑“更多”菜单，可直接选择“关闭预设”。
   - 常用内容面板底部保留同一关闭入口。
   - 关闭前明确提示：字段当前内容和资料库记录都会保留，仅隐藏该字段的预设操作。

2. [x] **关闭状态持久化**
   - 继续复用 schema v7 `preset_field_settings`，写入该 fieldKey 的 `enabled = false`。
   - 不修改 `common_presets.applicable_field_keys_json`，不删除任何预设项，不调用字段 `onApply`。
   - 同一 fieldKey 的多个表单位置同步关闭，工作区重载或应用重启后仍读取关闭状态。

3. [x] **图标语义修复**
   - “选择常用”从 `quickAction` 闪电图标改为 `presetLibrary` 书签图标。
   - 图标继续使用现有 Lucide / `AppIcon` 体系，没有新增依赖。
   - 其他真正表示快速执行的 `quickAction` 场景不受影响。

## Validation

- `npx tsc --noEmit`: passed.
- Targeted ESLint: passed.
- `node scripts/test_common_presets.cjs`: passed.
- All `scripts/test_subject_funding*.cjs`: passed.
- `npm run build`: passed; existing Vite large-chunk warning remains.
- `cargo fmt -- --check`: passed.
- `cargo test common_presets::tests`: passed; covers enable → disable persistence and preset data retention.
- `cargo test db::tests::schema_v6_migrates_preset_field_settings_to_v7`: passed.
- Full `cargo test`: 12 passed, 2 existing docfill tests failed because `项目全生命周期文件模版/效益分析表 .xlsx` is missing.
- Local Vite service returned HTTP 200. Browser visual automation remained unavailable in this session.

## Remaining

- Desktop Tauri click-through should verify the more menu and confirmation dialog under light/dark, font scale, and density combinations.
- Phase 2 has not started; full project preset templates and one-click multi-field application remain future work.

# 常用资料与项目预设阶段 1.5

- **Status:** Done
- **Objective:** 将第一阶段快捷填充升级为可复用的字段级预设机制，并在所有用户可见位置使用业务化字段信息。

## Completed

1. [x] **字段元信息 registry**
   - `presetFieldKeys.ts` 集中维护 `fieldKey`、业务名称、说明、所属模板、分组、字段类型、资格、推荐分类、别名和默认启用状态。
   - 未登记字段默认不可使用预设；UI 回退为“未命名字段 / 暂未配置”，不显示内部 key。

2. [x] **按需启用与持久化**
   - schema v7 新增 workspace SQLite 表 `preset_field_settings`。
   - 普通文本字段未启用时显示轻量“+ 预设”；启用后显示选择常用、保存当前、替换/追加和停用入口。
   - 同一 fieldKey 的多个渲染位置共享启用状态，重载工作区后从 SQLite 恢复。

3. [x] **业务化字段展示**
   - 预设中心卡片、适用字段选择器、快捷填充面板和保存区域展示字段名称、适用模板与所属分组。
   - “收入条款”“支出条款”均显示适用于“立项签批表、会审纪要”，所属分组为“商务条款”。

4. [x] **财务字段排除**
   - registry 将金额、税率、比例、现金流、NPV、利润率、智能反算和差额承接定义为不可用字段。
   - Rust 允许列表同时校验字段启用和 preset 绑定，绕过 UI 传入财务 fieldKey 也会被拒绝。

5. [x] **首批验证字段**
   - 保持第一阶段字段默认启用并补全业务元信息。
   - 新增“产权归属”作为默认关闭、用户主动启用的普通输入框代表。
   - 未改动任何财务公式、项目保存域或文档生成变量链路。

## Validation

- `npx tsc --noEmit`: passed.
- Targeted ESLint for changed frontend files: passed.
- `node scripts/test_common_presets.cjs`: passed.
- All `scripts/test_subject_funding*.cjs`: passed.
- `npm run build`: passed; existing Vite large-chunk warning remains.
- `cargo fmt -- --check`: passed after formatting.
- `cargo test common_presets::tests`: passed, including forbidden financial binding tests.
- `cargo test db::tests::schema_v6_migrates_preset_field_settings_to_v7`: passed.
- Full `cargo test`: 12 passed, 2 existing docfill tests failed because local template `项目全生命周期文件模版/效益分析表 .xlsx` is missing.
- Browser visual automation was unavailable because the configured Browser plugin did not expose its execution tool in this session.

## Remaining

- Desktop Tauri click-through with a real workspace should verify enable/save/fill/reload behavior and light/dark/font/density presentation.
- Phase 2 remains out of scope: full project preset sets, project-creation application, one-click multi-field application, AI recommendation, and AI collaboration.

# Tauri 启动空白/关闭卡死修复

- **Status:** Done
- **Objective:** 修复启动时上次工作区自动恢复阻塞 Tauri 窗口/WebView 初始化，导致页面空白且关闭响应异常的问题。

## Completed

1. [x] **启动恢复后台化**
   - `main.rs` 不再在 `setup` 阶段同步调用工作区打开逻辑。
   - `workspace.rs` 新增后台恢复入口，把 `lastOpenedWorkspacePath` 打开、SQLite 初始化、根目录注册和每日备份放入 `spawn_blocking`。
   - 恢复失败仍写入 `WorkspaceRuntime.startup_error`，不改变前端错误展示路径。

2. [x] **前端状态同步**
   - `App.tsx` 监听 `lamber-workspace-state-changed`。
   - 主窗口挂载后先订阅事件再刷新工作区状态，后台恢复完成后自动刷新 `useWorkspaceStore`。

## Validation

- `npm run build --prefix src-ui`: passed; Vite reported the existing large chunk warning.
- `cargo fmt -- --check` in `src-tauri`: passed.
- `cargo check` in `src-tauri`: passed with existing warnings unrelated to this change.
- Tauri dev startup smoke test with the configured last workspace: WebView2 is created immediately with the app process instead of showing the previous ~65 second delayed initialization.

## Remaining

- Manual visible-window click-through should confirm the Hub renders normally and the window closes from the title-bar X on the user's desktop session.

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

9. [x] **Preset center list-first layout**
   - `PresetCenterView` no longer keeps the full create form visible by default.
   - Added a list toolbar search box that filters by name, category, content, and tags while preserving kind/category/sort filters.
   - Added an on-demand right-side create/edit panel with close/cancel discard confirmation for unsaved draft changes.
   - Preset cards now show denser business metadata: category, enabled state, content摘要, tags, applicable-field summary, usage count, and last-used time.
   - Fixed the create/edit panel viewport height so bottom actions are not clipped, and desktop/mobile outside clicks now request panel close.

## Validation

- `npx tsc --noEmit` in `src-ui`: passed.
- `npx eslint src/views/PresetCenterView.tsx` in `src-ui`: passed.
- `node scripts/test_common_presets.cjs` in `src-ui`: passed.
- Existing `scripts/test_subject_funding*.cjs` in `src-ui`: passed.
- `npm run build` in `src-ui`: passed; Vite reported the existing large chunk warning.
- Full `npm run lint` in `src-ui`: blocked by existing lint errors in `useAiContextStore.ts` and `useAppearanceStore.ts`; the changed `PresetCenterView.tsx` passes targeted lint.
- `cargo fmt -- --check` in `src-tauri`: passed.
- `cargo test common_presets::tests` in `src-tauri`: passed.
- `cargo test benefit::calculator::tests` in `src-tauri`: passed.
- Full `cargo test` in `src-tauri`: blocked by existing missing docfill test template `项目全生命周期文件模版/效益分析表 .xlsx` (10 passed, 2 failed).
- Local Vite service at `http://localhost:5173`: HTTP 200.

## Remaining

- Desktop Tauri runtime manual click-through should be performed with a real workspace to verify CRUD, quick fill, persistence after reload, and usage-count updates end to end.
- Phase 2 remains out of scope: full project preset templates, project creation from presets, one-click multi-field application, intelligent recommendation, and AI collaboration.
