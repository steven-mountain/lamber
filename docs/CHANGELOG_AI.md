# CHANGELOG_AI.md

> [!NOTE]
> **历史兼容性说明**：本文件记录 AI 代理所做的结构修改、业务规则和上下文更改，不再作为 AI 每次任务的默认必读文件。
> 只有在追溯历史回归或用户明确要求时，才需要加载和查看此变更日志。

This changelog records structural modifications, business rules, and context changes made by AI agents to maintain a reliable project state mapping.

## 2026-06-07

### Windows Release Version 1.1.0

- Unified the Tauri bundle, Rust crate, root npm package, and frontend npm package versions as `1.1.0`.
- The Windows NSIS installer filename now carries `1.1.0` instead of repeatedly producing `1.0.0`.
- Version `1.1.0` is a minor release for the completed common-material, business-dictionary, and full project-preset capabilities.

### Full Project Preset Templates Phase 2

Created/modified:
- Added `project_presets.rs` and schema v9 tables `project_preset_templates` / `project_preset_template_entries`.
- Added typed frontend service, field safety/value helpers, management UI, current-project save/application preview, and three application strategies.
- Added optional project preset selection to project creation with compensating rollback.
- Extended `TemplateForms` with controlled project-preset bindings and seed consumption through the existing save path.
- Added Rust and frontend regression tests and updated project documentation.

Decision:
- Existing-project application never uses a backend direct-write command; it updates mounted formal form state and invokes unified save.
- Dictionary fields store business values. Financial and computed fields are forbidden by frontend registry checks and a Rust allowlist.
- New-project template values use a seed only until the matching controlled form persists them through its existing save handler.

Tests:
- TypeScript, targeted ESLint, related frontend scripts, production build, Rust formatting, schema v9 migration, and full Rust tests passed.
- Browser automation was unavailable because its JavaScript execution tool was not exposed.

### Unified Field Preset Panel Views

Modified:
- `CommonPresetQuickFill.tsx`: Removed portal-based save/edit modals and introduced one anchored panel view state for selection, save-current, and edit.
- `test_common_presets.cjs`: Replaced modal assertions with regression checks for the three panel views, absence of portal/overlay rendering, and return-to-selection after save/update.
- Common preset module and current-task documentation were updated.

Decision:
- All field-preset operations stay in the field-local panel context. Selection, save, and edit are mutually exclusive views rather than simultaneously rendered surfaces.
- Existing create/update/delete services, preset data shape, usage tracking, field replacement/append behavior, and field-preset persistence remain unchanged.

### Preset Replace Semantics and Editing

Modified:
- `CommonPresetQuickFill.tsx`: Renamed the primary row action from “插入” to “替换” and added a dedicated edit modal for preset name, content, category, and tags.
- `test_common_presets.cjs`: Added regression checks for replace wording, edit fields, update identity/binding preservation, list refresh, and removal of the old visible insert label.
- Common preset module and current-task documentation were updated.

Decision:
- Whole-row activation and the “替换” button share the existing replacement and usage-tracking path.
- Editing reuses `save_common_preset` with the original record id and non-editable metadata. It never calls the owning form setter and therefore cannot change the current project field.
- The picker remains open behind the edit modal so a successful update returns the user to the refreshed list.

### Common Preset Picker and Save Dialog Split

Modified:
- `CommonPresetQuickFill.tsx`: Separated list selection from preset creation. The picker now contains only list operations plus a secondary save entry, while “保存当前” opens a dedicated portal modal with field context, read-only content preview, and labeled metadata inputs.
- `test_common_presets.cjs`: Added regression checks for the modal, semantic labels, default-name generation, and removal of the inline expandable save form.
- Common preset module and current-task documentation were updated.

Decision:
- Selection and creation are distinct user intents and must not share the same expanded panel.
- Save still writes the existing `CommonPresetInput` shape and refreshes through the existing list service. No backend command, database table, or business-state path changed.

### Preset Row Click and Direct Action Repair

Modified:
- `CommonPresetQuickFill.tsx`: Restored whole-row insertion with keyboard support, strengthened insert/append button visibility, exposed delete directly, and removed the redundant per-item more menu.
- `test_common_presets.cjs`: Added regression checks for row insertion, child-action propagation guards, weak-destructive delete styling, and removal of menu state.
- Common preset module and current-task documentation were updated.

Decision:
- The row and insert button share the existing replacement path, including usage metadata updates.
- Insert, append, and delete buttons stop propagation to prevent accidental replacement.
- Delete remains a confirmed soft delete followed by list reload and never mutates current project field content.

### Compact Field Preset Rows and Item Actions

Modified:
- `CommonPresetQuickFill.tsx`: Replaced tall preset cards with compact list rows, added explicit insert/append actions, and added confirmed soft deletion in a per-item more menu.
- `test_common_presets.cjs`: Added regression checks for field-type-aware append separators, soft delete, list refresh, and item action visibility.
- Common preset module and current-task documentation were updated.

Decision:
- `text_snippet` append uses a newline; `short_value` append uses a space. Empty-field append remains equivalent to insert.
- Insert and append continue through the owning form setter and update preset usage metadata.
- Item deletion uses the existing soft-delete service, never clears current project content, and refreshes the matching list after success.

## 2026-06-06

### Field Preset Picker Visual Hierarchy

Modified:
- `CommonPresetQuickFill.tsx`: Reframed the field panel as a list-first common-content picker, added clickable preset cards with business metadata, introduced a compact first-save empty state, and downgraded save/disable controls to secondary bottom regions.
- `test_common_presets.cjs`: Added source-level regression checks for the picker heading, empty-state action, usage metadata, tags, and removal of duplicated bound-field details.
- Common preset module and current-task documentation were updated.

Decision:
- Field context is supporting information; matching reusable content is the panel's primary task and visual region.
- The existing save, apply, usage tracking, and disable persistence paths remain unchanged. Disabling a field preset continues to preserve current field content and common materials.

### Field Preset Interaction Simplification

Modified:
- `CommonPresetQuickFill.tsx`: Added outside-pointer closing for the field preset picker/save panel and replaced the one-item more menu with a direct close-preset button.
- `test_common_presets.cjs`: Added regression checks for outside-close handling and removal of the redundant menu state.
- Common preset module and current-task documentation were updated.

Decision:
- A one-item overflow menu adds interaction cost without adding organization value. Closing the field capability is now directly available while retaining the existing confirmation and data-retention semantics.

### Common Preset Phase 1.5 Final Coverage and Business Dictionaries

Created/modified:
- Added `business_dictionaries.rs`, schema v8 tables `business_dictionaries` / `business_dictionary_items`, default dictionaries, workspace-scoped item CRUD, enable/disable, soft delete, and ordering commands.
- Added `businessDictionaryService.ts`, `BusinessDictionarySelect.tsx`, and `BusinessDictionaryManager.tsx`.
- Extended `presetFieldKeys.ts` with `dictionaryKey`, radio/checkbox/date field types, controlled-field metadata, and additional reusable free-text field definitions.
- Updated `TemplateForms.tsx` so representative controlled fields read dictionary options while formal values remain in existing template state.
- Added the business-dictionary tab to `PresetCenterView.tsx` and expanded free-text preset headers.
- Added frontend and Rust regression tests plus schema v7-to-v8 migration coverage.

Decision:
- Common presets apply only to reusable free text. Select/radio/checkbox-like business options use dictionaries and cannot be bound as common presets.
- Dictionaries are option sources only. They do not overwrite project values, participate in document generation, or become an AI write path.
- Disabled/deleted options disappear from new choices, while an old saved value remains visible as the current inactive value.
- Phase 2 full-project templates and one-click multi-field application remain out of scope.

Tests:
- TypeScript, targeted ESLint, frontend production build, common-preset tests, business-dictionary tests, subject-funding tests, Rust formatting, dictionary CRUD tests, preset tests, and schema v8 migration test passed.
- Automated Browser visual verification was unavailable because the execution tool was not exposed.

### Template Path Verification and Branch Attendee Layout Repair

Modified:
- `TemplateForms.tsx`: Replaced the fixed-width branch-name/attendee flex row with a responsive proportional grid so both labels and preset actions align without wrapping into the input row.
- `docfill.rs`: Added test-template discovery through `LAMBER_TEMPLATE_ROOT`, the legacy repository directory, and the current module workspace template directory.
- `CURRENT_TASK.md`: Replaced the stale missing-template test result with the verified template location and passing docfill tests.

Tests:
- TypeScript passed.
- Targeted ESLint produced only the existing `TemplateForms` hook warnings.
- Both docfill Excel tests passed using `D:\HermesJang\CMCC\tools\workspace\templates\效益分析表 .xlsx`.

### Common Preset Phase 1.5 Close and Icon Polish

Modified:
- [CommonPresetQuickFill.tsx](../src-ui/src/components/common-presets/CommonPresetQuickFill.tsx): Added a compact more-actions menu with “关闭预设”, retained the panel-level close action, and added a confirmation that field content and reusable materials remain intact.
- [iconMap.ts](../src-ui/src/components/icons/iconMap.ts): Added `presetLibrary` using the existing Lucide `Bookmark` icon and `more` using `MoreHorizontal`.
- [common_presets.rs](../src-tauri/src/common_presets.rs): Extended tests to verify enable → disable persistence while preserving the bound common preset record.
- [test_common_presets.cjs](../src-ui/scripts/test_common_presets.cjs): Added regression assertions that preset selection no longer uses `quickAction` and that close/retention semantics remain present.
- Phase 1.5 project and module documentation was updated.

Decision:
- Closing a field preset is not an unbind/delete operation. It updates only workspace capability state in `preset_field_settings`; current form values, reusable materials, and shared applicable-field bindings remain unchanged.
- The close action must be discoverable without opening the full common-content picker, so it lives in a compact more menu aligned with the field label actions.
- Reusable-material selection uses bookmark/library semantics. Lightning remains available only for genuine quick-execution actions elsewhere.
- Phase 2 project preset templates and one-click multi-field application remain out of scope.

Tests:
- Frontend TypeScript, targeted ESLint, common-preset tests, subject-funding regression scripts, and production build passed.
- Rust formatting, common-preset persistence tests, and schema v6-to-v7 migration test passed.
- Full Rust suite: 12 passed; two existing docfill tests remain blocked by the missing local Excel template.
- Local Vite returned HTTP 200; browser visual automation was unavailable in this session.

### Common Materials & Project Presets Phase 1.5

Created/modified:
- [presetFieldKeys.ts](../src-ui/src/lib/presetFieldKeys.ts): Expanded the stable key catalog into a field metadata registry with business labels, templates, groups, field types, eligibility, recommended categories, aliases, defaults, and explicit financial exclusions.
- [CommonPresetQuickFill.tsx](../src-ui/src/components/common-presets/CommonPresetQuickFill.tsx): Added workspace-persisted opt-in field activation, synchronized duplicate field controls, business metadata display, and explicit choose/save/replace/append/disable actions.
- [PresetCenterView.tsx](../src-ui/src/views/PresetCenterView.tsx): Replaced raw fieldKey display in cards and binding selection with field name, applicable templates, and business groups.
- [IctBasicInfo.tsx](../src-ui/src/components/IctBasicInfo.tsx): Added “产权归属” as the first default-off representative field.
- [common_presets.rs](../src-tauri/src/common_presets.rs), [db.rs](../src-tauri/src/db.rs), and [main.rs](../src-tauri/src/main.rs): Added schema v7 `preset_field_settings`, list/set commands, and backend eligible-key validation for field activation and preset binding.
- [commonPresetService.ts](../src-ui/src/services/commonPresetService.ts) and [test_common_presets.cjs](../src-ui/scripts/test_common_presets.cjs): Added typed IPC access and registry/exclusion tests.
- Project/module context documents were updated for the Phase 1.5 architecture and Phase 2 exclusions.

Decision:
- Field activation is workspace UI configuration, not project form data. Formal values continue through existing field setters and save domains.
- Existing Phase 1 fields remain default-enabled for compatibility; newly connected ordinary text fields can default off and require explicit activation.
- Financial safety is enforced in both layers. The UI registry hides ineligible fields, while Rust rejects unknown, amount, percent, and computed field bindings even through direct IPC.
- Raw fieldKey values are internal identifiers. User-facing surfaces must show business metadata and neutral fallback text when metadata is missing.

Tests:
- Frontend TypeScript, targeted ESLint, common-preset tests, subject-funding regression scripts, and production build passed.
- Rust formatting, common-preset tests, and schema v6-to-v7 migration test passed.
- Full Rust suite ran 12 successful tests; the two existing docfill tests remain blocked by the missing local Excel template.
- Browser visual automation was unavailable in this session because its execution tool was not exposed.

## 2026-06-05

### Tauri Startup Workspace Restore Hang Repair

Modified:
- [main.rs](../src-tauri/src/main.rs): Stopped synchronously restoring `lastOpenedWorkspacePath` during Tauri `setup`; startup now creates `WorkspaceRuntime` and schedules restore work in the background.
- [workspace.rs](../src-tauri/src/workspace.rs): Added `spawn_restore_last_workspace`, kept the existing workspace open path for restore, and emits `lamber-workspace-state-changed` after restore success or failure.
- [App.tsx](../src-ui/src/App.tsx): Subscribes to `lamber-workspace-state-changed` before the initial workspace-state refresh so the frontend updates after background restore completes.
- [ARCHITECTURE_MAP.md](./ARCHITECTURE_MAP.md) and [CURRENT_TASK.md](./CURRENT_TASK.md): Documented the startup restore lifecycle.

Decision:
- Startup must not synchronously wait for workspace SQLite initialization, daily backup, hidden-attribute updates, or filesystem checks. These operations can touch user-controlled disks and should not block WebView creation, rendering, or title-bar close handling.
- Restore failures continue to flow through `WorkspaceRuntime.startup_error`; this keeps the existing `WorkspaceGate`/workspace error UX without adding a parallel frontend source of truth.

Tests:
- Ran `npm run build --prefix src-ui`: passed with the existing Vite chunk-size warning.
- Ran `cargo fmt -- --check` in `src-tauri`: passed.
- Ran `cargo check` in `src-tauri`: passed with existing warnings unrelated to this change.
- Tauri dev startup smoke test with the configured last workspace showed WebView2 creation immediately alongside the app process, replacing the previous delayed initialization behavior.

### Common Preset Center List-First Layout

Modified:
- [PresetCenterView.tsx](../src-ui/src/views/PresetCenterView.tsx): Reworked the management page into a list-first layout with an on-demand create/edit side panel, search toolbar, denser preset cards, collapsible applicable-field selection, fixed panel actions, and discard confirmation for unsaved drafts.
- [PresetCenterView.tsx](../src-ui/src/views/PresetCenterView.tsx): Constrained the side panel to viewport height so bottom actions are not clipped, and added desktop/mobile outside-click close behavior that respects unsaved-change confirmation.
- [common-presets.md](../docs/modules/common-presets.md) and [CURRENT_TASK.md](../docs/CURRENT_TASK.md): Recorded the management-page layout rule and validation status.

Decision:
- The preset center should not reserve permanent space for an empty create form. Create and edit are explicit user actions that open an auxiliary panel; the list remains the primary scanning surface.
- Search stays as a frontend filter over already loaded workspace presets, so category filtering, sort order, CRUD commands, enabled/disabled state, and the SQLite `common_presets` schema remain unchanged.

Tests:
- Ran `npx tsc --noEmit` in `src-ui`: passed.
- Ran `npx eslint src/views/PresetCenterView.tsx` in `src-ui`: passed.
- Ran `node scripts/test_common_presets.cjs` in `src-ui`: passed.
- Ran `npm run build` in `src-ui`: passed with the existing Vite chunk-size warning.
- Full `npm run lint` remains blocked by pre-existing lint errors in `useAiContextStore.ts` and `useAppearanceStore.ts`.
- `npm run typecheck` is not available in `src-ui`; the project uses `npx tsc --noEmit` / `npm run build` for TypeScript validation.

### Common Preset Quick-Fill Field Header Alignment

Modified:
- [CommonPresetQuickFill.tsx](../src-ui/src/components/common-presets/CommonPresetQuickFill.tsx): Added `CommonPresetFieldHeader`, a reusable form header that keeps the visible label and preset actions in one responsive row.
- [IctBasicInfo.tsx](../src-ui/src/components/IctBasicInfo.tsx) and [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Replaced split label/action rows with `CommonPresetFieldHeader` across the first-phase preset-connected fields.
- [common-presets.md](../docs/modules/common-presets.md) and [CURRENT_TASK.md](../docs/CURRENT_TASK.md): Recorded the form-side layout rule.
- [CommonPresetQuickFill.tsx](../src-ui/src/components/common-presets/CommonPresetQuickFill.tsx): Added `CommonPresetLabelHeader` for plain fields that need to align with preset-enabled fields, and compacted the quick-fill action buttons.
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Removed the detached sign-off payment preset band and attached the revenue/expenditure preset actions to the actual payment input labels.

Decision:
- Preset buttons should not be aligned with fixed offsets or separate `justify-end` rows. They now use a shared flex header so normal desktop fields keep the label and actions on the same line, while narrow containers can wrap naturally.
- Forms that mix preset and non-preset fields should use the shared header components for both variants so neighboring inputs align vertically.

## 2026-06-04

### Common Materials Quick-Fill Field Coverage Expansion

Modified:
- [presetFieldKeys.ts](../src-ui/src/lib/presetFieldKeys.ts): Added stable keys for demand-import fields, meeting-review fields, sign-off IT/CT service fields, and shared revenue/expenditure payment method fields.
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Added `CommonPresetQuickFill` controls for project demand unit, demand service content, customer confirmation, deployment environment requirement, onsite support staff, IT/CT construction content, revenue collection method, expenditure payment method, time requirement, and sign-off IT/CT service content.
- [common-presets.md](../docs/modules/common-presets.md), [CURRENT_TASK.md](../docs/CURRENT_TASK.md), and [test_common_presets.cjs](../src-ui/scripts/test_common_presets.cjs): Recorded and tested the expanded field coverage.

Decision:
- Revenue collection and expenditure payment methods use shared `payment.*` field keys because meeting-review and sign-off sections write the same formal `revCollection` / `expPayment` state. This lets one saved common material appear in both places without duplicating records.

### Common Materials & Project Presets Phase 1

Created:
- [common_presets.rs](../src-tauri/src/common_presets.rs): Added workspace-scoped reusable material commands for listing, saving, enabling/disabling, soft deletion, and usage tracking.
- [PresetCenterView.tsx](../src-ui/src/views/PresetCenterView.tsx): Added the independent "常用资料与项目预设" management page with short-field and long-text tabs.
- [presetFieldKeys.ts](../src-ui/src/lib/presetFieldKeys.ts): Added stable field-key definitions for reusable content and future project preset binding.
- [CommonPresetQuickFill.tsx](../src-ui/src/components/common-presets/CommonPresetQuickFill.tsx): Added reusable field-side picker/save control.
- [commonPresetService.ts](../src-ui/src/services/commonPresetService.ts): Added frontend IPC wrapper for common preset commands.
- [test_common_presets.cjs](../src-ui/scripts/test_common_presets.cjs): Added field-key catalog tests.

Modified:
- [db.rs](../src-tauri/src/db.rs): Added `common_presets` table initialization and schema version 6 migration.
- [main.rs](../src-tauri/src/main.rs): Registered common preset Tauri commands.
- [App.tsx](../src-ui/src/App.tsx), [useNavigationStore.ts](../src-ui/src/store/useNavigationStore.ts), and [iconMap.ts](../src-ui/src/components/icons/iconMap.ts): Added the first-level Hub entry and route for `preset_center`.
- [IctBasicInfo.tsx](../src-ui/src/components/IctBasicInfo.tsx): Connected quick fill to customer name and project background.
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Connected quick fill to project background, technical solution, meeting reviewers, branch/department name, and risk/project owner fields.
- Project context docs were updated to record the module boundary, SQLite structure, fieldKey mechanism, first connected fields, and later-phase exclusions.

Tests:
- Ran `npx tsc --noEmit` in `src-ui`: passed.
- Ran `node scripts/test_common_presets.cjs` in `src-ui`: passed.
- Ran all `scripts/test_subject_funding*.cjs` in `src-ui`: passed.
- Ran `npm run build` in `src-ui`: passed with the existing Vite chunk-size warning.
- Ran `cargo fmt -- --check` in `src-tauri`: passed after formatting.
- Ran `cargo test common_presets::tests` in `src-tauri`: passed.
- Ran `cargo test benefit::calculator::tests` in `src-tauri`: passed.
- Ran full `cargo test` in `src-tauri`: 10 passed, 2 failed because the local docfill test template `项目全生命周期文件模版/效益分析表 .xlsx` is missing.

Decision:
- Presets are only reusable fill sources. After a user applies content, the owning form state is updated through existing setters and save domains; document generation and AI context continue to read official project/template state.
- Phase 1 deliberately does not implement full project preset templates, project-creation preset selection, one-click multi-field application, automatic history extraction, AI recommendation, or AI auto-fill.

### ICT Selection Fee: Fixed Quote/Limit Anchor

Modified:
- [useIctCalculations.ts](../src-ui/src/hooks/useIctCalculations.ts): Added `selectionFeeAnchor` for mutually exclusive quote/limit anchoring. Markup changes now call the forward selection-fee command when quote is fixed and the reverse command when limit is fixed. Added request-sequence protection so stale async invoke results cannot overwrite newer user input.
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Added compact fixed-state dot buttons beside supplier quote and selection limit labels, wired with `aria-pressed`, and added accessible labels to the selection-fee inputs.
- [selection-fee.md](../docs/modules/selection-fee.md), [PROJECT_INDEX.md](../docs/PROJECT_INDEX.md), and [CURRENT_TASK.md](../docs/CURRENT_TASK.md): Recorded the anchor model, UI behavior, and validation status.

Tests:
- Ran `npx tsc --noEmit` in `src-ui`: passed.
- Ran `npm run build` in `src-ui`: passed.
- Ran `cargo test` in `src-tauri`: passed (11/11; existing warnings only).
- Browser smoke-tested the local Vite page: fixed-dot default, mutual exclusion, and edit-to-anchor state passed. Numeric invoke calculation requires the Tauri runtime and was not executed in the plain browser page.

Decision:
- The root cause was that markup edits always used supplier quote as the implicit source of truth. Adding another local conditional would preserve the hidden coupling, so the bounded fix makes the quote/limit source of truth explicit and reuses the existing Rust forward/reverse commands.

### ICT Subject Funding Plans Final: Migration and Single Official Source

Modified:
- [ictSubjectFundingPlan.ts](../src-ui/src/lib/ictSubjectFundingPlan.ts): Added migration version `SUBJECT_FUNDING_PLAN_MIGRATION_VERSION = 1`, `legacy_migration` audit reason, migration-plan creation, and `migrateLegacySubjectFundingPlans()` to fill only missing non-zero subject plans while preserving existing valid/custom/equal plans and surfacing invalid plans through coverage validation.
- [useIctState.ts](../src-ui/src/hooks/useIctState.ts) and [useIctCalculations.ts](../src-ui/src/hooks/useIctCalculations.ts): Defaulted state and payloads to `subject_funding_plans`; official annual cashflow overrides now always come from subject funding plans when coverage is valid, and `CashflowSegment` no longer contributes formal annual cashflow.
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Migrates old project loads after tax-item hydration, writes `subjectFundingPlanMigrationVersion`, removes the calculation-source switch, and keeps coverage summary, batch init, locate, clear-all, balance allocation, CT linkage, and smart reverse flows on the new source.
- [IctBasicInfo.tsx](../src-ui/src/components/IctBasicInfo.tsx), [IctCashflowTable.tsx](../src-ui/src/components/IctCashflowTable.tsx), and [IctSubjectFundingPlanEditor.tsx](../src-ui/src/components/IctSubjectFundingPlanEditor.tsx): Removed old model A-E / segment funding UI and old-source explanatory text; cashflow preview and drill-down now present only the subject-plan source.
- [ProjectFilesTab.tsx](../src-ui/src/components/project/ProjectFilesTab.tsx), [projectService.ts](../src-ui/src/utils/projectService.ts), [models.rs](../src-tauri/src/benefit/models.rs), [excel.rs](../src-tauri/src/benefit/excel.rs), and [calculator.rs](../src-tauri/src/benefit/calculator.rs): Added migration-version compatibility and made Excel import payloads use the subject funding source.
- Added [test_subject_funding_migration.cjs](../src-ui/scripts/test_subject_funding_migration.cjs) and [test_subject_funding_final.cjs](../src-ui/scripts/test_subject_funding_final.cjs).

Tests:
- Ran `for f in scripts/test_subject_funding*.cjs; do node "$f"; done` in `src-ui`: passed.
- Ran `npx tsc --noEmit` in `src-ui`: passed.
- Ran `npm run build` in `src-ui`: passed.
- Ran `cargo test` and `cargo fmt -- --check` in `src-tauri`: passed; existing warnings remain.

Decision:
- The root issue was not just a visible source selector. Keeping `legacy_model` as a runnable branch meant old segment schedules could still compete with subject plans as annual cashflow sources. The bounded fix removes the user switch and formal legacy branch while retaining old fields for read/migration compatibility.
- Existing abnormal subject plans are intentionally not overwritten by migration. The app fills missing non-zero subjects, then lets the canonical coverage validator block formal calculation and saving until the user resolves the invalid rows.

### ICT Funding Plans: Default Activation and Cashflow Source Repair

Modified:
- [ictSubjectFundingPlan.ts](../src-ui/src/lib/ictSubjectFundingPlan.ts): Changed zero-amount sync from "disable and keep record" to removing the subject plan so the UI returns to "未维护"; plan mode switches and custom annual edits now force `enabled: true`; normalization preserves sync audit fields.
- [useIctState.ts](../src-ui/src/hooks/useIctState.ts): Removed the old gate that only synced plans when `cashflowCalculationSource === "subject_funding_plans"`. Amount edits now always sync subject plans, create missing upfront plans, backfill other positive missing subjects, and switch the formal cashflow source to `subject_funding_plans` on positive amount input.
- [IctSubjectFundingPlanEditor.tsx](../src-ui/src/components/IctSubjectFundingPlanEditor.tsx): Opening a zero-amount subject no longer creates a persisted 0 yuan plan; missing plans show an unchecked enable state until explicitly created.
- [subject-funding-plan.md](../docs/modules/subject-funding-plan.md) and [CURRENT_TASK.md](../docs/CURRENT_TASK.md): Recorded the new default activation and zero-reset behavior.

Tests:
- Ran `npm run build` in `src-ui`: passed.
- Ran `for f in scripts/test_subject_funding*.cjs; do node "$f"; done` in `src-ui`: passed.

Decision:
- The root cause was split ownership between subject amount input and cashflow source selection. Users could maintain a subject annual plan while the official cashflow table still read the legacy global model, so custom year-2 amounts appeared in the editor but not in the 10-year result.
- The stable bounded fix makes subject amount input the convergence point: positive amounts create/enable default subject plans and activate the subject-plan source; clearing an amount removes that subject's plan and returns it to "未维护".

## 2026-06-03

### ICT Lifecycle Phase 3.5/4 Regression Repair: Legacy Restore, Clear-All, Coverage Locate, Excel Cashflow

Modified:
- [models.rs](../src-tauri/src/benefit/models.rs): Added backward-compatible serde defaults for IT cashflow fields and `IctResult.cashflow` alias/default support so legacy snapshots without Phase 4 fields can still deserialize.
- [calculator.rs](../src-tauri/src/benefit/calculator.rs): Updated test `IctInput` builders for the optional `subject_funding_plans` field.
- [useIctState.ts](../src-ui/src/hooks/useIctState.ts): Added a unified `clearFinancialSubjects()` state action that clears all revenue/cost amounts, subject names, subject funding plans, Model E segment amount schedules, balance allocation, tail-difference state, and reconciliation prompts.
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Restored tax items from both `incl_tax/tax_rate` and `incl/tax/excl` shapes; added confirmed "一键清空全部收入和支出"; added coverage issue drill-down that reuses `subjectFundingCoverage.issues[0]` and auto-opens the target funding plan editor.
- [IctSubjectFundingPlanEditor.tsx](../src-ui/src/components/IctSubjectFundingPlanEditor.tsx): Added `forceOpenToken` so coverage locate can expand a collapsed plan editor.
- [IctCashflowTable.tsx](../src-ui/src/components/IctCashflowTable.tsx): Completed the project-wide vs IT-only 10-year view with IT cumulative net and PV values, avoiding invalid `NaN` display.
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Fixed Excel cashflow variable source from the incorrect `metrics.cashflows` to the formal `metrics.cashflow`; only emits `CASH_IN_Y1..10` and `CASH_OUT_Y1..10` when official cashflow rows exist so template fallback formulas remain intact.
- [subject-funding-plan.md](../docs/modules/subject-funding-plan.md) and [CURRENT_TASK.md](../docs/CURRENT_TASK.md): Recorded repair rules and current validation status.

Tests:
- Ran `npm run build` in `src-ui`: passed.
- Ran `node scripts/test_subject_funding_plan.cjs`, `node scripts/test_subject_funding_cashflow.cjs`, and `node scripts/test_subject_funding_sync.cjs` in `src-ui`: passed.
- Ran `cargo test` in `src-tauri`: 11/11 passed; existing warnings remain.
- Ran `cargo fmt -- --check` in `src-tauri`: passed after formatting.

Decision:
- The old-project zero-amount symptom was addressed at the deserialization boundary rather than masking it in the UI. Legacy output metrics missing new IT fields now load with zero IT defaults, preserving existing project-wide values.
- Clear-all treats custom subject naming as data on fixed standard subjects because the current catalog has no separate dynamic subject-instance structure. The safe behavior is to restore fixed subject rows to blank initial state and remove all user-entered financial schedules.
- Coverage locate reuses the canonical validation result so issue ordering and classification cannot drift from the blocking calculation-source logic.
- Investment-benefit Excel multi-year cashflow uses the official calculator result, not a parallel reconstruction, preserving old-model and subject-plan compatibility.

### ICT Subject-Level Funding Plans Phase 3.5 & 4: Status Semantics, Batch Init, & Drill-Down

Created:
- [test_subject_funding_phase4.cjs](../src-ui/scripts/test_subject_funding_phase4.cjs): Comprehensive tests for zero-recovery (proportional preservation), exact reason tracking, batch initialization skipping existing plans, and annual cashflow drill-down generation.
- [docs/modules/subject-funding-plan.md](../docs/modules/subject-funding-plan.md): Modular design doc covering the entire subject funding plan system, rules, and integration boundaries.

Modified:
- [ictSubjectFundingPlan.ts](../src-ui/src/lib/ictSubjectFundingPlan.ts): Added `lastValidAnnualInclValues`, `lastChangeReason`, and `lastChangedAt` to `SubjectFundingPlan` state for zero-recovery and transparency. Added `initializeMissingSubjectFundingPlans` helper and `buildAnnualCashflowSubjectContributions` drill-down generator. Appended `revenueSubjects` and `costSubjects` arrays to `FundingPlanCoverageResult`.
- [useIctState.ts](../src-ui/src/hooks/useIctState.ts): Updated `updateTaxItem` and `updateTaxItemsInclBatch` to accept optional `reason` arguments from caller layers, propagating them to the sync algorithm to replace "manual_amount_sync" with accurate business contexts like "reverse_calculation_sync" or "ct_linkage_sync".
- [useIctCalculations.ts](../src-ui/src/hooks/useIctCalculations.ts): Injected specific update reasons (`reverse_calculation_sync` and `balance_allocation_sync`) into amount update dispatches.
- [IctSubjectFundingPlanEditor.tsx](../src-ui/src/components/IctSubjectFundingPlanEditor.tsx): Added a subtle, friendly UI status pill displaying the `lastChangeReason` (in Chinese) for auto-adjusted plans.
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Added batch initialization buttons ("一键生成...") visible when `subject_funding_plans` calculation source is active. Re-enabled the smart reverse button while in subject plans mode (removed accidental phase 3 blocking logic).
- [IctCashflowTable.tsx](../src-ui/src/components/IctCashflowTable.tsx): Implemented an interactive drill-down state `expandedYear` that renders nested per-subject revenue/cost contribution breakdowns natively below each cashflow year row when expanded.

Tests:
- Ran `node scripts/test_subject_funding_phase4.cjs` in `src-ui`: 4/4 passed.
- Ran `cargo test` in `src-tauri`: 10/10 passed.
- Ran `npm run build` in `src-ui`: Build succeeded.

Decision:
- Zero Recovery: The UI requires subjects to seamlessly toggle on and off. Instead of deleting zeroed plans, we preserve their scale in `lastValidAnnualInclValues`. Re-adding an amount restores the prior timeline proportions precisely.
- Explicit Sync Reason: Differentiating manual edits from reverse calcs or linkages directly aids user comprehension. Storing `lastChangeReason` at the domain layer ensures the UI displays an accurate origin narrative, increasing trust.
- Drill-Down Readability: Aggregated NPV values hide composition. Exposing the exact `buildAnnualCashflowSubjectContributions` mappings within the cashflow table removes the "black box" feeling, letting users trace final outputs directly back to input subjects.

### ICT Subject-Level Funding Plans Phase 3: Amount-Change Synchronization

Created:
- [test_subject_funding_sync.cjs](../src-ui/scripts/test_subject_funding_sync.cjs): 12 pure-function tests covering proportional scaling, tail-difference correction, auto-create upfront, clear-to-zero, zero-total fallback, batch sync, CT linkage, mode preservation, scale-down, negative amounts, and immutability.

Modified:
- [ictSubjectFundingPlan.ts](../src-ui/src/lib/ictSubjectFundingPlan.ts): Added `syncSubjectFundingPlanToAmount` (single-subject proportional scale with tail correction, auto-create, and zero-clear) and `syncSubjectFundingPlansToAmounts` (batch variant). Uses integer-cents arithmetic internally to avoid floating-point drift.
- [useIctState.ts](../src-ui/src/hooks/useIctState.ts): Injected funding plan sync into `updateTaxItem` and `updateTaxItemsInclBatch`. When `cashflowCalculationSource === "subject_funding_plans"`, any `incl` or `excl` field change (including CT linkage side-effects) triggers `syncSubjectFundingPlansToAmounts` in the same React state batch. Tax-rate-only changes (`field === "tax"`) do not trigger sync.
- [useIctCalculations.ts](../src-ui/src/hooks/useIctCalculations.ts): Added `fundingPlansOverride` option to `buildInputDataPayload` so candidate evaluations use simulated-synced plans. Added `buildCandidateSyncUpdates` helper for CT-aware sync update construction. Modified `buildReverseCandidate` and `buildLockedTotalStructureCandidate` to compute and inject synced plans during candidate evaluation.
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Removed the `subject_funding_plans` reverse calculation block — smart reverse now works under both calculation sources.

Tests:
- Ran `node scripts/test_subject_funding_sync.cjs` in `src-ui`: 12/12 passed.
- Ran `node scripts/test_subject_funding_plan.cjs` in `src-ui`: passed (existing tests unaffected).
- Ran `node scripts/test_subject_funding_cashflow.cjs` in `src-ui`: passed (existing tests unaffected).
- Ran `npx tsc --noEmit` in `src-ui`: zero errors.

Decision:
- Sync logic centralized in `updateTaxItem` / `updateTaxItemsInclBatch` (the convergence points for all amount writes), rather than scattered across each caller. This ensures all 5 write paths (manual edit, normal reverse, locked-total reverse, balance allocation, CT linkage) are automatically covered.
- Proportional scaling preserves the user's custom year-by-year distribution shape. Mode (`upfront`, `equal`, `custom`) is preserved across syncs.
- `legacy_model` mode is completely unaffected; sync is gated by `cashflowCalculationSource`.

### ICT Subject-Level Funding Plans Phase 2

Created:
- [test_subject_funding_cashflow.cjs](../src-ui/scripts/test_subject_funding_cashflow.cjs): Added pure Node tests for subject-plan coverage validation, cashflow source normalization, and annual cashflow aggregation with per-subject tax conversion.

Modified:
- [ictSubjectFundingPlan.ts](../src-ui/src/lib/ictSubjectFundingPlan.ts): Added `CashflowCalculationSource`, coverage issue/result types, full coverage validation, and subject-plan annual cashflow generation helpers. The annual cashflow helper converts each subject's annual tax-inclusive plan values to tax-exclusive values using that subject's tax rate before summing.
- [useIctState.ts](../src-ui/src/hooks/useIctState.ts): Added `cashflowCalculationSource`, normalized setter, and legacy default behavior.
- [useIctCalculations.ts](../src-ui/src/hooks/useIctCalculations.ts): Added coverage and annual cashflow derivation, serialized `cashflow_calculation_source`, and routed valid subject-plan annual arrays through existing Rust direct cashflow override fields. Invalid active subject-plan coverage blocks recalculation and keeps the previous result.
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Restores, saves, and clears the calculation source; renders the cashflow source selector and coverage summary; blocks switching into subject-plan mode when coverage is incomplete; blocks official benefit saves and document generation when active subject-plan coverage is invalid; disables smart reverse while subject-plan mode is active.
- [IctSubjectFundingPlanEditor.tsx](../src-ui/src/components/IctSubjectFundingPlanEditor.tsx) and [IctCashflowTable.tsx](../src-ui/src/components/IctCashflowTable.tsx): Updated UI text and cashflow previews so the active source is visible and stale subject-plan states are explicit.
- [projectService.ts](../src-ui/src/utils/projectService.ts), [ProjectFilesTab.tsx](../src-ui/src/components/project/ProjectFilesTab.tsx), [models.rs](../src-tauri/src/benefit/models.rs), [calculator.rs](../src-tauri/src/benefit/calculator.rs), and [excel.rs](../src-tauri/src/benefit/excel.rs): Added payload/source compatibility fields and legacy defaults.

Tests:
- Ran `node scripts/test_subject_funding_plan.cjs` in `src-ui`: passed.
- Ran `node scripts/test_subject_funding_cashflow.cjs` in `src-ui`: passed.
- Ran `npm run build` in `src-ui`: TypeScript and Vite build passed.
- Ran `cargo test` in `src-tauri`: 10 tests passed; only existing compiler warnings were emitted.

Decision:
- Subject-level funding plans affect official cashflow only after an explicit user switch to `subject_funding_plans`; old projects and projects with plans remain on `legacy_model` unless switched.
- New-source coverage failures do not fall back to legacy. The app keeps prior cashflow/metrics visible with an invalid/stale warning and blocks benefit saves/doc generation until coverage is repaired or legacy mode is selected.
- Smart reverse remains a legacy-source feature for now because the current reverse solvers do not update every affected subject-level annual plan.

### ICT Subject-Level Funding Plans Phase 1

Created:
- [ictSubjectFundingPlan.ts](../src-ui/src/lib/ictSubjectFundingPlan.ts): Added subject-level funding plan types and pure helpers for `side + groupId + key` IDs, default upfront plans, 1-10 year equal splits, custom annual value updates, normalization, advisory validation, upsert, and deleted-subject cleanup.
- [IctSubjectFundingPlanEditor.tsx](../src-ui/src/components/IctSubjectFundingPlanEditor.tsx): Added inline revenue/cost subject funding plan editor with "收款计划" / "付款计划" wording, three modes (`upfront`, `equal`, `custom`), 10-year tax-inclusive annual inputs, enabled toggling, and difference/status messaging.
- [test_subject_funding_plan.cjs](../src-ui/scripts/test_subject_funding_plan.cjs): Added a lightweight Node logic test for funding-plan helpers.

Modified:
- [useIctState.ts](../src-ui/src/hooks/useIctState.ts): Added `subjectFundingPlans` current state plus normalized setter, upsert, and subject-ref cleanup methods.
- [useIctCalculations.ts](../src-ui/src/hooks/useIctCalculations.ts): Serializes `subject_funding_plans` into the existing lifecycle/AI payload for persisted context. Calculation formulas and reverse solvers are unchanged.
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Restores subject funding plans from lifecycle payloads or cashflow assumptions, persists them under cashflow assumptions, marks cashflow dirty when they change, clears stale plans when entering free/empty contexts, and renders the inline editor below each subject row.
- [PROJECT_STATUS.md](../docs/PROJECT_STATUS.md), [ARCHITECTURE_MAP.md](../docs/ARCHITECTURE_MAP.md), [AI_CONTEXT.md](../docs/AI_CONTEXT.md): Updated long-term project status and architecture context.

Tests:
- Ran `node scripts/test_subject_funding_plan.cjs` in `src-ui`: passed.
- Ran `npm run build` in `src-ui`: TypeScript and Vite build passed.
- Ran `cargo test` in `src-tauri`: 10 tests passed; only existing compiler warnings were emitted.

Decision:
- Subject-level funding plans are current business state, not a calculation source in Phase 1.
- Plans bind to concrete subject instances using `side + groupId + key`, not just `subjectCode`, so future duplicate standard subjects can maintain independent schedules.
- Existing `cashflowModel`, `CashflowSegment`, NPV/IRR/payback, balance allocation, and smart reverse behavior remain unchanged.
- Existing projects without `subject_funding_plans` / `subjectFundingPlans` normalize to `{}`. No migration from old funding models or segment schedules is performed in this phase.

### Custom Accent Color Real-Time Preview Fix

Modified:
- [SettingsView.tsx](../src-ui/src/components/settings/SettingsView.tsx): Resolved a bug where custom accent colors (especially low-contrast inputs) did not update the real-time preview or warning banners correctly. Implemented a `useEffect` that continuously translates user-typed hex values (including HSL self-adjusted safe versions) and applies them to the DOM for instant preview. Supported flexible hex formats (3/6 digits, with or without leading '#').

### Front-End Advanced Appearance Customization & Accessibility Safeguards (Phase 3)

Created:
- [colorUtils.ts](../src-ui/src/theme/colorUtils.ts): Custom RGB/HSL conversions, relative luminance calculations, and WCAG contrast ratio checkers.
- [deriveAccentTokens.ts](../src-ui/src/theme/deriveAccentTokens.ts): Generates WCAG-compliant derived HSL accent tokens (`primary`, `primary-foreground`, `primary-soft`, `ring`, `accent`, `accent-foreground`) from custom colors using HSL lightness shifting.
- [test_color_math.js](../scratch/test_color_math.js): Verification script to test HSL conversion, WCAG contrast calculations, and automatic lightness adjustments for custom accent colors.

Modified:
- [appearance.ts](../src-ui/src/theme/appearance.ts): Extended `AppearanceSettings` type to support `contrastPreference`, `customAccent` settings, and version fields. Upgraded standard default settings to version 3.
- [presets.ts](../src-ui/src/theme/presets.ts): Refined HSL dark theme preset mappings (`DARK_THEMES`) for all 5 presets to establish distinct visual styles in dark modes. Added HSL contrast override variables.
- [applyAppearance.ts](../src-ui/src/theme/applyAppearance.ts): Implemented sequence-based color tokens resolution: standard presets are modified by custom accent calculations and then overlaid with high-contrast rules. Applied `data-contrast` attribute on document element.
- [useAppearanceStore.ts](../src-ui/src/store/useAppearanceStore.ts): Added preference setters and state for Custom Accent and Contrast. Added Version 3 localStorage settings parsing and configuration migration paths.
- [SettingsView.tsx](../src-ui/src/components/settings/SettingsView.tsx): Extended layouts with Accent color palette recommended selections, Custom HTML5 Color Picker, Standard vs High Contrast selectors, and warning banners for auto-adjusted low contrast accent inputs. Enhanced real-time preview panel with focused input ring and simulated AI chatbot bubble elements.
- [index.ts](../src-ui/src/theme/index.ts): Exported `colorUtils` and `deriveAccentTokens`.
- [DESIGN.md](../DESIGN.md): Appended Phase 3 visual specs, contrast boundaries, custom accent derivation rules, and dark theme refinement notes.
- [PROJECT_STATUS.md](../docs/PROJECT_STATUS.md), [ARCHITECTURE_MAP.md](../docs/ARCHITECTURE_MAP.md): Updated project status history, directory maps, and theme components definitions.

Decision:
- Accent color adjustments are validated dynamically against WCAG standards (4.5:1 standard minimum, 7.0:1 high contrast minimum) against light/dark backgrounds. If low contrast is detected, lightness is shifted in HSL, warning the user and applying a safe version.
- High Contrast preference applies pure white/black canvas background blocks, highly visible text, and thick distinct border lines.
- SQLite project database structures, business cashflow calculations, NPV metrics, and AI context composer schemas remain completely untouched.

### Front-End Appearance Settings Center & Theme Runtime Switching (Phase 2)


Created:
- [appearance.ts](../src-ui/src/theme/appearance.ts): Formulates brightness modes, color themes, font scaling levels, interface density configurations, default presets, and font scaling ratio factors.
- [presets.ts](../src-ui/src/theme/presets.ts): Formulates full HSL variables mapping for 5 corporate light themes (`lamber`, `graphite`, `navy`, `forest`, `warmStone`) and system dark base configuration with light preset primary adaptation.
- [applyAppearance.ts](../src-ui/src/theme/applyAppearance.ts): Writes color tokens, font scaling weights, and interface densities to `document.documentElement` attributes and styles dynamically.
- [useAppearanceStore.ts](../src-ui/src/store/useAppearanceStore.ts): Persists user styling configuration to `localStorage`, observes brightness system media query settings, applies styles to root, and manages cross-window event synchronization.
- [SettingsView.tsx](../src-ui/src/components/settings/SettingsView.tsx): Renders custom selectors for all style choices with a real-time mini business preview card and default restoration command.

Modified:
- [index.ts](../src-ui/src/theme/index.ts): Exports all theme, presets, configuration types, and DOM applier helper utilities.
- [index.css](../src-ui/src/index.css): Dynamically scales typographic line heights by `--font-scale` to prevent overlap, and implements custom overrides for card padding, button/input heights, table cell vertical paddings, and page spacing gaps.
- [card.tsx](../src-ui/src/components/ui/card.tsx), [button.tsx](../src-ui/src/components/ui/button.tsx), [input.tsx](../src-ui/src/components/ui/input.tsx): Rewrote primitive sizes and paddings to read custom CSS density variables with standard defaults.
- [main.tsx](../src-ui/src/main.tsx): Runs synchronous appearance store hydration prior to React rendering to ensure color and layouts render without startup flashing.
- [useNavigationStore.ts](../src-ui/src/store/useNavigationStore.ts): Integrated `"settings"` view type and track previous view context to return safely on settings panel close.
- [App.tsx](../src-ui/src/App.tsx): Routes settings views and places Settings button in `HubView` header toolbar.
- [ProjectBoard.tsx](../src-ui/src/views/ProjectBoard.tsx): Places Settings button on the kanban board header toolbar next to the GlobalSaveButton.
- [WorkspaceHeader.tsx](../src-ui/src/components/WorkspaceHeader.tsx): Places Settings button on the header toolbar of all workspaces next to the GlobalSaveButton.
- [DESIGN.md](../DESIGN.md): Documented specifications for presets, mode observers, scaling bounds, density levels, and application mechanisms.
- [PROJECT_STATUS.md](../docs/PROJECT_STATUS.md), [ARCHITECTURE_MAP.md](../docs/ARCHITECTURE_MAP.md), [AI_CONTEXT.md](../docs/AI_CONTEXT.md): Updated project milestones, directory indexes, and visual token guidelines.

Decision:
- Appearance preferences are strictly application-level settings and must never be written to project databases or workspaces.
- Color presets are dynamically set via raw space-separated numbers on root styles to respect Tailwind's HSL wrapper.
- Cross-window visual synchronizations are dispatched via Tauri window event listener bridges.

### Front-End Global Visual Foundation Refactoring (Phase 1)

Created:
- [tokens.ts](../src-ui/src/theme/tokens.ts): Define design system color schemes, border radius (`lg: var(--radius)`, `md`, `sm`), and sizing.
- [typography.ts](../src-ui/src/theme/typography.ts): Declare base line-heights, weight bindings, and typographic roles mapped to `--font-scale`.
- [index.ts](../src-ui/src/theme/index.ts): Centralized theme exports.

Modified:
- [index.css](../src-ui/src/index.css): Added new HSL color variables and typography scale calculations. Integrated `.numeric-value` class.
- [tailwind.config.js](../src-ui/tailwind.config.js): Extended Tailwind configuration with semantic color names and typographic roles.
- [button.tsx](../src-ui/src/components/ui/button.tsx), [input.tsx](../src-ui/src/components/ui/input.tsx), [card.tsx](../src-ui/src/components/ui/card.tsx), [label.tsx](../src-ui/src/components/ui/label.tsx): Refactored primitive components to adhere to visual standards.
- [App.tsx](../src-ui/src/App.tsx), [ProjectBoard.tsx](../src-ui/src/views/ProjectBoard.tsx), [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx), [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Migrated views to utilize semantic styles, typography variables, and tabular numbers.
- [DataManagement.tsx](../src-ui/src/views/DataManagement.tsx): Refactored path lists, workspace cards, status labels, and relocation tables to use semantic HSL tokens.
- [AiChatPanel.tsx](../src-ui/src/components/ai/AiChatPanel.tsx): Refactored connection status badge and quick actions.
- [MessageBubble.tsx](../src-ui/src/components/MessageBubble.tsx): Refactored inline badges, user/assistant chat bubbles, blockquotes, and code block components.
- [DESIGN.md](../DESIGN.md): Updated document with the Lamber Global Visual Specification v1 details.

Tests:
- Run `npm run build` inside `src-ui`: TypeScript compilation and Vite build succeeded.
- Run `cargo test` in `src-tauri`: All unit tests passed.

Decision:
- Visual elements must follow the "No-Line" rule using background HSL shifts instead of thick solid borders.
- Sizing and typography must support dynamic resizing based on `--font-scale` variable.
- Financial numerical representation enforces tabular-nums across the UI.

### ICT Lifecycle Balance Control UI Layout Alignment

Modified:
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Changed the balance control columns container to top-align (using `md:items-start` instead of `md:items-center`), and wrapped the input container and summary card/text elements in a matched `h-[38px]` height container. This ensures Row 1 (labels), Row 2 (inputs/summaries), and Row 3 (1% own-product prompts) align row-by-row across columns.

Tests:
- Ran `npm run build` in `src-ui`: TypeScript compilation and Vite build succeeded.
- Ran `cargo test benefit::calculator::tests` in `src-tauri`: All unit tests passed.

## 2026-06-02

### ICT Lifecycle Subject Role Configuration Optimization

Created:
- [IctSubjectRoleComponents.tsx](../src-ui/src/components/IctSubjectRoleComponents.tsx): Added modular UI components `SubjectRoleActions` (row-level menu for setting/clearing balancing and reverse roles on each subject) and `SelectedSubjectRoleSummary` (card summarizing active role, supporting locate and clear actions), along with helper utilities `scrollToSubject` and `highlightSubjectElement` for smooth-scroll tab switching and visual outline feedback.

Modified:
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Integrates role assignment directly in subject table rows, removing the legacy select dropdowns from the top control areas and right panel. Simplified the right reverse panel to automatically infer target details and reverse side from the selected target, disabling the execute button if no target is set. Removed unused imports and methods.

Tests:
- Ran `npm run build` in `src-ui`: TypeScript check and Vite build passed.
- Ran `cargo test` in `src-tauri`: 8 passed, 2 expected failures due to missing template file.
- Ran `cargo test benefit::calculator::tests`: 7 passed.
- Ran `cargo test ai_context::service::tests`: 1 passed.

Decision:
- Entry point for balancing and reverse target roles is shifted to the concrete subject rows for direct visual context.
- Mutual exclusions (balancing subjects cannot be reverse targets) are enforced with warnings on hover/click.
- Switching balancing subjects prompts for confirmation and preserves the old balancing subject's current amount value.
- Right reverse panel automatically adapts its state based on the side and reverse mode of the active target.

### ICT Model E Structure Reverse Segment Sync

Modified:
- [useIctCalculations.ts](../src-ui/src/hooks/useIctCalculations.ts): Removed the `model_e` amount-mode structure reverse entrance block and added a shared candidate/final segment synchronization path for locked-total structure reverse. The sync maps stable reverse subjects to `CashflowSegment` side/scope buckets, applies target and balancing deltas before `calculate_ict_benefit`, rejects invalid candidates, and writes the accepted synced segment array back only after the final calculation succeeds.
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Extended the structure reverse hint so `model_e` amount mode tells users that segmented cashflow amount plans will be synchronized.

Tests:
- Ran `npm run build` in `src-ui`.
- Ran `cargo test` in `src-tauri`: 8 passed, 2 failed because the local test template `项目全生命周期文件模版/效益分析表 .xlsx` is missing for the two `docfill::tests` lifecycle Excel cases.
- Ran `cargo test benefit::calculator::tests` in `src-tauri`: 7 passed.
- Ran `cargo test ai_context::service::tests` in `src-tauri`: 1 passed.

Decision:
- `model_e` amount-mode structure reverse is supported only where the selected subject and balancing subject map to existing segment buckets. Revenue IT/CT/non-IT-CT subjects map to `revenueScope`; cost IT/CT/non-IT-CT/mixed subjects map to `costScope`.
- Same-bucket transfers preserve the aggregate bucket total; cross-bucket transfers move equal and opposite deltas between buckets. Existing custom annual amount plans are scaled through the amount-mode annual adjustment helper.
- Candidates are excluded from reachability and final write-back if they require a missing bucket or would create a negative segment, bucket, or annual amount.
- CT product and line revenue still mirror to their paired CT cost subjects and now synchronize the linked cost-side segment bucket. If that cross-side linkage collides with an active locked-total investment balancing rule, the solver blocks conservatively instead of attempting a four-variable cross-side reverse.
- The revenue own-product 1% prompt remains display-only.

### ICT Locked-Total Structure Reverse Calculation

Modified:
- [ictReverseCalculation.ts](../src-ui/src/lib/ictReverseCalculation.ts): Added reverse mode resolution for `normal`, `locked_total_structure`, and `blocked`, plus locked-total structure context construction, dual-subject candidate application, and bounded sample-point generation.
- [useIctCalculations.ts](../src-ui/src/hooks/useIctCalculations.ts): Added the locked-total structure solver. It samples the finite reallocatable pool, detects metric-insensitive and unreachable targets, chooses the crossing solution closest to the current target amount, validates total preservation/non-negative amounts, and writes target plus balancing subject results together.
- [useIctState.ts](../src-ui/src/hooks/useIctState.ts): Added `updateTaxItemsInclBatch` for inclusive-amount batch writes, preserving tax-exclusive recomputation and CT revenue-to-cost amount pairing without same-group stale-state overwrite risk.
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Replaced same-side balance blocking with automatic reverse mode detection, added the structure-mode hint in the smart reverse panel, kept balancing subjects disabled, and passed the resolved reverse context into the calculation hook.

Tests:
- Ran `npm run build` in `src-ui`.

Decision:
- Same-side reverse under a valid total balance rule is now handled as structure reverse: the selected target subject changes and the balancing subject moves inversely so the side's inclusive total stays unchanged.
- `model_e` amount mode is blocked for structure reverse because the current segment data model cannot reliably map two changed subjects to distinct cashflow destinations. Normal selected-subject reverse remains supported in amount mode.
- The revenue own-product 1% prompt is still display-only in the current UI; this phase did not add a blocking validation rule for it.

### ICT Dynamic Reverse Subject Calculation

Created:
- [ictReverseCalculation.ts](../src-ui/src/lib/ictReverseCalculation.ts): Added shared helpers for reverse subject stable references, eligible subject options, selected-subject display names, candidate tax-inclusive amount application, CT paired amount mirroring, and balance-allocation conflict validation.

Modified:
- [useIctCalculations.ts](../src-ui/src/hooks/useIctCalculations.ts): Replaced fixed frontend reverse writes to `rev_it_integration` / `cost_it_integration` with a dynamic selected-subject solver. Candidate payloads now support all revenue and cost subject groups, reuse `calculate_ict_benefit`, preserve the existing margin/NPV-rate target metrics, and write final results through `updateTaxItem`.
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Added the smart reverse subject selector, clears invalid subject selections when switching reverse side, disables active balancing subjects, and blocks same-side locked-total reverse conflicts with a clear message.

Tests:
- Ran `npm run build` in `src-ui`.

Decision:
- The current balance allocation implementation is present on both revenue and investment sides, so same-side reverse conflict handling is applied to both sides.
- The reverse subject is identified by stable catalog fields (`side`, `subjectCode`, `groupId`, `key`) and display names are presentation-only.
- Locked-total structure reverse is intentionally not implemented in this phase. When a same-side balance allocation rule is valid, same-side reverse is blocked; cross-side reverse remains allowed.
- The old Rust fixed reverse commands remain registered for compatibility, but the ICT frontend smart reverse panel now evaluates selected subjects through the general `calculate_ict_benefit` path.

### ICT Balance Allocation Rules

Created:
- [ictBalanceAllocation.ts](../src-ui/src/lib/ictBalanceAllocation.ts): Added shared frontend helpers for revenue/investment balance rule normalization, serialization, subject reference matching, inclusive-amount difference calculation, and validation status reporting.

Modified:
- [useIctState.ts](../src-ui/src/hooks/useIctState.ts): Added `balanceAllocation` state with independent `revenue` and `investment` rules.
- [useIctCalculations.ts](../src-ui/src/hooks/useIctCalculations.ts): Serializes `revenue_balance_rule` and `investment_balance_rule` into lifecycle input payloads and AI context payloads.
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Replaced separate quick split panels with revenue/investment total balance controls, removes one-click quick fill and integration-service preview from the control area, shows the revenue own-product minimum prompt as 1% of the revenue total, applies valid balancing amounts through `updateTaxItem`, makes active balancing subject amount fields read-only while preserving tax-rate editing, persists rules in lifecycle/cashflow saves, restores rules from current state, and blocks cashflow/document-generation navigation on negative balance validation.
- [projectService.ts](../src-ui/src/utils/projectService.ts): Added typed balance rule payload fields to `IctInput`.
- [models.rs](../src-tauri/src/benefit/models.rs): Added optional `revenue_balance_rule` and `investment_balance_rule` to `IctInput` so benefit snapshots preserve the new rule configuration.
- [calculator.rs](../src-tauri/src/benefit/calculator.rs) and [excel.rs](../src-tauri/src/benefit/excel.rs): Updated test/import input constructors with disabled balance rules.

Tests:
- Ran `npm run build` in `src-ui`.
- Ran `cargo test` in `src-tauri`.

Decision:
- Balance differences are based on inclusive amounts and are written back only through the existing tax item update path, preserving existing inclusive/tax/exclusive linkage and CT paired update behavior.
- Switching a balancing subject clears the previous balancing subject's inclusive amount to `0` before the new subject receives the balancing difference, so the same balancing amount does not remain on both subjects.
- Negative balance differences are validation errors, not formal amounts. They block cashflow and document generation before the 0-tolerance reconciliation modal can be bypassed.
- Smart reverse calculation remains fixed to the existing current targets and algorithms in this phase. Arbitrary-subject reverse calculation and total-locked reverse solving remain a follow-up phase.

## 2026-06-01

### ICT Sign-off Project Situation Itemization

Modified:
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Reordered the sign-off-only configuration section to Project Background, IT/CT Service Content, advance-payment and post-approval-selection checkboxes, then revenue collection and expenditure payment methods.
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Removed manual sign-off billing-subject override inputs and now generates `PROJECT_INVESTMENT_SITUATION` / `PROJECT_REVENUE_SITUATION` by enumerating non-zero measurement-table subjects with billing-subject-name first, standard-subject fallback, and fixed category prefixes.
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Meeting-review project overall investment wording now reuses the same generated investment sentence as the sign-off project situation investment line.
- [docfill.rs](../src-tauri/src/docfill.rs): Added generation-time normalization for older sign-off templates that still contain the hardcoded IT/CT-only project situation wording.
- [【2025版】ICT项目立项签批表（仅适用50万以下项目）模板.docx](../项目全生命周期文件模版/【2025版】ICT项目立项签批表（仅适用50万以下项目）模板.docx): Replaced the hardcoded IT/CT-only project situation wording with full-line placeholders for generated investment and revenue situation text.

Tests:
- Ran `npm run build` in `src-ui`.

Decision:
- The optional "申请立项后甄选" wording is controlled by a sign-off checkbox and is no longer hardcoded in the template.
- Sign-off project situation wording is presentation-only and continues to use existing tax-exclusive amount fields; no financial calculation paths changed.

### ICT Billing Subject Name Extension

Modified:
- [ictSubjectCatalog.ts](../src-ui/src/lib/ictSubjectCatalog.ts): Added `billingSubjectName` / `billing_subject_name` support and the shared `resolveBillingSubjectPresentation` resolver for Excel display names, document business names, and document dedup keys.
- [useIctState.ts](../src-ui/src/hooks/useIctState.ts), [useIctCalculations.ts](../src-ui/src/hooks/useIctCalculations.ts), and [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Added the "计费科目名称（文书/计费口径）" input for every existing revenue/cost subject, synchronized it across CT paired subjects, and persisted it through lifecycle/cashflow/benefit payloads without changing amount behavior.
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Continued using catalog helpers for Excel variables, sign-off wording, meeting-review wording, and business-composition deduplication so billing subject names now take priority over product/business names.
- [models.rs](../src-tauri/src/benefit/models.rs), [excel.rs](../src-tauri/src/benefit/excel.rs), [ProjectFilesTab.tsx](../src-ui/src/components/project/ProjectFilesTab.tsx), [projectFileService.ts](../src-ui/src/services/projectFileService.ts), and [projectService.ts](../src-ui/src/utils/projectService.ts): Extended the serialized item DTOs with optional `billing_subject_name` while preserving old data without the field.

Tests:
- Ran `npm run build` in `src-ui`.
- Ran `cargo test` in `src-tauri`; existing Excel subject-row tests still verify name/G/Q writes and blank amount clearing.

Decision:
- `customSubjectName` remains the product/business name field; the new billing subject name is stored independently and only affects presentation/export/document wording.
- The resolver priority is `计费科目名称 > 具体业务/产品名称 > 标准科目名称` for Excel/page display and `计费科目名称 > 具体业务/产品名称 > existing fallback` for documents.
- No standard subject rows, Excel formulas, tax calculations, cashflow math, selection fee logic, or reverse-calculation behavior were changed.

### ICT Subject Custom Business Name Extension

Created:
- [ictSubjectCatalog.ts](../src-ui/src/lib/ictSubjectCatalog.ts): Added a fixed subject catalog mapping stable subject codes, UI groups, standard subject names, Excel variable prefixes, and document business prefixes.

Modified:
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx) and [useIctState.ts](../src-ui/src/hooks/useIctState.ts): Added custom business/product name inputs for every existing revenue and cost subject, storing the optional custom name separately from amount/tax fields and restoring old data safely.
- [useIctCalculations.ts](../src-ui/src/hooks/useIctCalculations.ts), [models.rs](../src-tauri/src/benefit/models.rs), [excel.rs](../src-tauri/src/benefit/excel.rs), [ProjectFilesTab.tsx](../src-ui/src/components/project/ProjectFilesTab.tsx), and [projectFileService.ts](../src-ui/src/services/projectFileService.ts): Persist custom subject names through lifecycle/snapshot payloads and preserve them when parsing/importing exported lifecycle Excel files.
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Generates Excel subject display names, document business names, and deduplicated business composition values from the subject catalog.
- [docfill.rs](../src-tauri/src/docfill.rs): Replaced scattered Excel cell writes with a unified `3-直接经济效益评估表` subject-row mapping that writes the subject name, `G` tax-exclusive amount, and `Q` tax-inclusive amount for every standard subject row.

Tests:
- Added a Rust unit test verifying that CT product revenue and CT other product cost write `产品收入（视频监控）` / `其他产品成本（视频监控）`, `G=41.51`, and `Q=44` into the lifecycle Excel sheet.
- Added a Rust unit test verifying that blank amount variables clear `G/Q` cells instead of writing numeric `0`.

Decision:
- Custom business/product names are metadata on fixed standard billing subjects, not new subject rows. The implementation does not support adding/deleting subjects or inserting Excel rows.
- Financial calculations remain keyed by the existing standard subjects and continue to use only amount/tax fields; custom names affect UI labels, Excel output/import, and document wording only.
- Business composition wording deduplicates repeated document business names across revenue and cost, while amount detail fields keep each side separate.
- Empty/zero frontend amount inputs remain blank in generated Excel amount cells; paired CT subject custom names follow existing amount pass-through relationships.

### Meeting Investment Subject Alignment Fix

Modified:
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Meeting-review `PROJECT_TOTAL_INVESTMENT_DETAIL` now uses the same resolved IT/CT cost subjects as the sign-off form variables instead of deriving CT subject wording from the mid-platform capability name.

Decision:
- The presales meeting-review "项目整体投入金额" wording should follow the sign-off form's investment billing subjects. This changes only document subject wording and keeps the existing tax-exclusive IT, CT, mixed-cost, and total investment calculations intact.

### ICT Cashflow Price Persistence Hydration Fix

Modified:
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Added current-state hydration that overlays `project_cashflow_states.assumptions_json`, payment model, and cashflow segments onto lifecycle/snapshot input data before filling calculator fields.

Decision:
- Price fields belong to the cashflow domain during ordinary edits and can be saved without rewriting lifecycle input payloads.
- Reopening an ICT project must treat cashflow assumptions as the latest current editor state for IT/CT revenue and cost inputs, otherwise stale lifecycle payloads can restore old prices.

### Inquiry Vendor Image State Preservation Fix

Modified:
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): One-click three-vendor quote generation now merges existing vendor screenshots into regenerated quote rows instead of clearing `images`.
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Vendor screenshot upload now uses functional state updates so async file reads cannot overwrite newer vendor quote state.

Decision:
- Vendor quote screenshots are part of the vendor row state and must survive quote amount regeneration.
- The merge prefers vendor-name matches and falls back to row index to support both regenerated default vendors and manually edited rows.

### Template Image Document Embedding Fix

Modified:
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Document-generation image payloads now use `assetId` as the primary `data` value and include `assetId` explicitly, preventing frontend preview URLs from being serialized into Word variables.
- [docfill.rs](../src-tauri/src/docfill.rs): JSON image parsing now prefers `assetId`, embeds only asset IDs or legacy `data:image` values, and suppresses unresolved image JSON so raw payload text does not appear in generated Word documents.

Decision:
- Frontend `asset://localhost/...` URLs are preview-only and must not be used as document image source data.
- The backend remains the authority for resolving image binaries through Workspace SQLite asset ownership and file resolution.

### Lifecycle Document Workspace Output Path Fix

Modified:
- [docfill.rs](../src-tauri/src/docfill.rs): Added Workspace-aware lifecycle document output resolution. Relative output folders are expanded against the active Workspace root, and `projectId` can be used to derive the project directory from the Workspace SQLite `projects` record when no explicit output directory is provided.
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Passes the active `projectId` to `generate_lifecycle_docs`.

Decision:
- ICT lifecycle generated documents belong in the bound project directory under the active Workspace, not under the Tauri backend working directory.
- The previous narrow symptom was a relative path such as `都市花园254号` being interpreted as `src-tauri/都市花园254号`, which also caused Tauri dev watch rebuild/restart behavior. The fix centralizes output path resolution in the backend instead of adding frontend-only path string workarounds.

### AI Workspace Specified Project Context Routing

Created:
- [workspaceProjectRouter.ts](../src-ui/src/ai/context/workspaceProjectRouter.ts): Added deterministic Workspace project-name and template-name routing helpers for AI chat context composition.

Modified:
- [ai_context/dto.rs](../src-tauri/src/ai_context/dto.rs), [ai_context/service.rs](../src-tauri/src/ai_context/service.rs), and [ai_context/commands.rs](../src-tauri/src/ai_context/commands.rs): Added the read-only `list_ai_workspace_projects` command, returning lightweight current-Workspace project identity, saved-state existence flags, and saved template names without paths, file contents, image bytes, or full template JSON.
- [main.rs](../src-tauri/src/main.rs): Registered `list_ai_workspace_projects`.
- [aiProjectContextService.ts](../src-ui/src/services/aiProjectContextService.ts): Added typed frontend DTOs and invoke wrapper for the Workspace project index.
- [buildAiChatContext.ts](../src-ui/src/ai/context/buildAiChatContext.ts) and [types.ts](../src-ui/src/ai/context/types.ts): Extended the composer to route explicitly named projects to real `projectId` reads, load at most two specified project contexts per turn, reuse `template_detail` for one uniquely resolved specified template, inject lightweight project index context for Workspace-level list questions, and keep current-page draft overlay scoped to its bound project.
- [AiChatPanel.tsx](../src-ui/src/components/ai/AiChatPanel.tsx): Added system rules for specified project context, ambiguity handling, projectId-based official reads, multi-project separation, and draft isolation.

Tests:
- Added a Rust unit test covering the lightweight Workspace project index flags and saved template-name metadata.

Decision:
- Project names are only current-turn routing hints inside the active Workspace; official data reads continue to use real `projectId` values through `build_ai_project_context`.
- The composer does not guess when project names are duplicate, ambiguous, or absent. It emits warnings and asks the model to request clarification rather than falling back to another project.
- Broad Workspace questions use the lightweight index and do not trigger full project/template reads across all projects.
- This phase does not add AI writes, RAG, embedding search, file full-text reads, automatic image loading, financial logic changes, or cross-Workspace queries.

## 2026-05-31

### AI Template Detail Context and Controlled Vision Assets

Created:
- [templateAssetSelection.ts](../src-ui/src/ai/templateAssetSelection.ts): Added a lightweight cross-window bridge for explicit template image analysis selections, carrying only `projectId`, `templateId`, `assetId`, and field metadata.

Modified:
- [ai_context/dto.rs](../src-tauri/src/ai_context/dto.rs), [ai_context/service.rs](../src-tauri/src/ai_context/service.rs), and [ai_context/commands.rs](../src-tauri/src/ai_context/commands.rs): Added `template_detail` support to the read-only AI Project Context Service, including saved template field sanitization, asset metadata summaries, and `load_ai_template_asset` for controlled vision reads.
- [main.rs](../src-tauri/src/main.rs): Registered the `load_ai_template_asset` command.
- [aiProjectContextService.ts](../src-ui/src/services/aiProjectContextService.ts): Added frontend DTOs for template detail and controlled template asset image loading.
- [buildAiChatContext.ts](../src-ui/src/ai/context/buildAiChatContext.ts): Requests `template_detail` only when the active ICT AI context is a template editing context, and keeps template saved detail separate from draft overlay.
- [AiChatPanel.tsx](../src-ui/src/components/ai/AiChatPanel.tsx), [AiRuntime.ts](../src-ui/src/ai/AiRuntime.ts), [AiInputBox.tsx](../src-ui/src/components/ai/AiInputBox.tsx), [ImageAttachmentPreview.tsx](../src-ui/src/components/ai/ImageAttachmentPreview.tsx), and [MessageBubble.tsx](../src-ui/src/components/MessageBubble.tsx): Reused the existing `image_url` multimodal path for explicitly selected template assets, resolving asset bytes through the backend only at send time.
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Publishes current `projectId` and `selectedTemplate` into the template AI context payload and adds explicit AI analysis actions to template image thumbnails.

Decision:
- Template detail context is read on demand for the current specified template only; ordinary project questions continue to use template summaries.
- Saved template content comes from Workspace SQLite (`project_template_states`, with `project_settings` fallback). Unsaved template edits remain a separate frontend draft overlay.
- Template image contents are never auto-loaded. Only explicitly selected images become current-turn vision attachments after backend project/asset ownership validation.
- The implementation does not add AI writes, auto-fill, RAG, embeddings, document full-text reads, unselected image reads, or financial calculation changes.

### AI Project Context Chat Integration

Created:
- [src-ui/src/ai/context/buildAiChatContext.ts](../src-ui/src/ai/context/buildAiChatContext.ts), [buildDraftOverlay.ts](../src-ui/src/ai/context/buildDraftOverlay.ts), and [types.ts](../src-ui/src/ai/context/types.ts): Added a lightweight frontend context composer that loads saved official project context from the backend AI Project Context Service and conditionally builds an unsaved frontend draft overlay from dirty page state.

Modified:
- [AiChatPanel.tsx](../src-ui/src/components/ai/AiChatPanel.tsx): Calls the context composer on every message send and passes layered saved/draft/warning context nodes to the existing `PromptAST`/`AiRuntime` streaming path.
- [AiChatPanel.tsx](../src-ui/src/components/ai/AiChatPanel.tsx): Resets the streaming parser before inserting a new assistant placeholder and overwrites the active placeholder from the current parser output, preventing the previous AI reply from appearing while a new response is still thinking.
- [useWorkspaceStore.ts](../src-ui/src/store/useWorkspaceStore.ts): Clears workspace-scoped active project/scheme identity from project, navigation, and legacy ICT local storage state whenever the active workspace is cleared or changed.
- [ProjectBoard.tsx](../src-ui/src/views/ProjectBoard.tsx) and [useAiContextStore.ts](../src-ui/src/store/useAiContextStore.ts): Project Board now replaces its AI context snapshot with the active workspace ID, project count, lightweight project card list, and selected project summary so AI can answer workspace-level board questions after switching workspaces.
- [buildAiChatContext.ts](../src-ui/src/ai/context/buildAiChatContext.ts), [buildDraftOverlay.ts](../src-ui/src/ai/context/buildDraftOverlay.ts), [useNavigationStore.ts](../src-ui/src/store/useNavigationStore.ts), and [useProjectStore.ts](../src-ui/src/store/useProjectStore.ts): Refresh the latest persisted navigation/current-project/ICT active project identity at message-send time so a floating AI window opened before project selection can still call `build_ai_project_context` for the current ICT-bound project.
- [PromptRenderer.ts](../src-ui/src/ai/PromptRenderer.ts): Renamed rendered context layers to distinguish Workspace SQLite official state from current unsaved draft overlay and loading notes.
- [aiContextKeys.ts](../src-ui/src/utils/aiContextKeys.ts), [App.tsx](../src-ui/src/App.tsx), and [ProjectBoard.tsx](../src-ui/src/views/ProjectBoard.tsx): Added Project Board AI context key support so project-detail dirty edits can be treated as page-level draft state.
- [useIctCalculations.ts](../src-ui/src/hooks/useIctCalculations.ts) and [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Added active `projectId` markers to frontend AI context payloads for draft/project consistency checks.
- [calculator.rs](../src-tauri/src/benefit/calculator.rs): Updated existing Rust unit-test fixture construction with `project_background: None` so tests compile after the earlier model field addition. This does not change calculation logic.

Decision:
- AI chat now treats Workspace SQLite data as the only saved official project state. Zustand/localStorage page snapshots are used only as unsaved draft overlays when current dirty scopes match the current project/page.
- Each chat turn owns a fresh streaming parser lifecycle; prior assistant text must remain only in its completed message and must not seed the new pending assistant bubble.
- Workspace-level Project Board summaries are allowed as current page context for list/count questions, but they remain read-only prompt context and do not replace project-level official SQLite detail retrieval.
- Active project identity may be refreshed from persisted frontend navigation/current-project selection at send time, but that identity is used only to request read-only Workspace SQLite context; local draft state is not promoted to saved official data.
- Context loading failures degrade into prompt warnings and must not block chat streaming.
- Draft overlays are sanitized to omit base64/data URL previews and absolute paths, truncate large content, and avoid reading any image/document binaries.
- This phase remains read-only and does not implement AI writes, saves, patch application, RAG, embeddings, file full-text summaries, image analysis, scans, repairs, or financial recalculation.

### AI Project Context Service

Created:
- [ai_context/mod.rs](../src-tauri/src/ai_context/mod.rs), [ai_context/dto.rs](../src-tauri/src/ai_context/dto.rs), [ai_context/service.rs](../src-tauri/src/ai_context/service.rs), and [ai_context/commands.rs](../src-tauri/src/ai_context/commands.rs): Added a read-only backend service and Tauri command `build_ai_project_context` for structured project-level AI context retrieval from the active Workspace SQLite database.
- [aiProjectContextService.ts](../src-ui/src/services/aiProjectContextService.ts): Added typed frontend invoke wrapper for later AI integration.

Modified:
- [main.rs](../src-tauri/src/main.rs): Registered the new `build_ai_project_context` command.

Decision:
- AI project context is sourced only from persisted `.lamber.sqlite` state through `WorkspaceRuntime`, not from frontend draft state or user-supplied paths.
- The service is strictly read-only: no database writes, folder scans, repairs, document full-text reads, image binary/base64 reads, or prompt injection are performed in this phase.
- Template and file contexts are summaries/metadata only. Template image assets expose counts, not absolute paths or binary content.

### Project Background, Collection/Payment and IT/CT Content Sync in Template Forms

Modified:
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Added "IT服务内容" and "CT服务内容" textareas to the 《立项签批表》 (Project Approval/Sign-off Form) section. Handled variables resolution in `handleGenerate` to prioritize these overrides with fallback to original values.

Decision:
- Allow users to override IT and CT service contents in the Sign-off Form template configuration.
- The input boxes display original default values when empty/unmodified and correctly fallback to defaults if cleared.

### Project Background and Collection/Payment Methods Synchronization in Template Forms

Modified:
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Added "项目背景" (project background), "收入侧收款方式" (revenue collection method) and "支出侧付款方式" (expenditure payment method) fields to the 《立项签批表》 (Project Approval/Sign-off Form) section. They are bound directly to the shared states.

Decision:
- Enable immediate, reactive dual-direction synchronization of the template form configuration parameters (Project Background, Collection & Payment methods) between the template configuration view and the global variables.
- Changes are fully tracked and persisted in the database state upon saving.

### Project Background Synchronization in Template Forms

Modified:
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Added a "项目背景" (project background) textarea configuration field to the 《立项签批表》 (Project Approval/Sign-off Form) section. It is bound directly to the shared `projectBackground` state.

Decision:
- Enable immediate, reactive dual-direction synchronization of the Project Background content between the Project Parameters form page and the Sign-off Form template configuration. Modifying either updates the other in real time.
- Changes are fully tracked and persisted in the database lifecycle state upon saving.

### Workspace Management Card Interaction Fix

Modified:
- [DataManagement.tsx](../src-ui/src/views/DataManagement.tsx): Made associated Workspace cards locally selectable without opening the Workspace, isolated inline open/reveal/unlink button events, and routes the explicit open action into the Project Board after the Workspace is active.

Decision:
- The Workspace Management tab should not reorder cards on plain selection. Card clicks only highlight the chosen record; opening a Workspace remains an explicit button action.
- The Workspace Management open action is an entry action, not just a background switch; successful opening should land in the Workspace's project board.

### Workspace Management UI Separation

Modified:
- [DataManagement.tsx](../src-ui/src/views/DataManagement.tsx): Moved associated Workspace records into a dedicated "Workspace Management" tab and restyled them as card-based workspace selectors matching the Project Board workspace picker.

Decision:
- Global Workspace association management should be visually and logically separate from current Workspace backup/health maintenance. The management tab handles local association records; the maintenance tab handles operations on the active Workspace.

## 2026-05-30

### Global Workspace Management

Modified:
- [workspace.rs](../src-tauri/src/workspace.rs): Added commands to unlink a remembered Workspace from local config and to close the current Workspace while clearing `lastOpenedWorkspacePath`.
- [main.rs](../src-tauri/src/main.rs): Registered the new Workspace management commands.
- [workspaceService.ts](../src-ui/src/utils/workspaceService.ts) and [useWorkspaceStore.ts](../src-ui/src/store/useWorkspaceStore.ts): Added frontend wrappers/store actions for unlinking and closing Workspaces.
- [DataManagement.tsx](../src-ui/src/views/DataManagement.tsx): Added an associated Workspace list with open, reveal, and unlink actions.

Decision:
- Workspace unlinking is a local association operation only. It removes the app's recent-workspace reference but never deletes the physical Workspace folder, database, or project files. If unlinking the active Workspace, the UI must run the dirty guard first and the backend must release the active runtime/database connection.

### Workspace Import Argument Guard

Modified:
- [workspaceMaintenanceService.ts](../src-ui/src/services/workspaceMaintenanceService.ts): Changed `importWorkspace` to accept explicit arguments instead of an options object.
- [WorkspaceGate.tsx](../src-ui/src/components/workspace/WorkspaceGate.tsx) and [DataManagement.tsx](../src-ui/src/views/DataManagement.tsx): Updated import calls to pass `openAfterImport` and `conflictStrategy` directly.
- [workspace_maintenance.rs](../src-tauri/src/workspace_maintenance.rs): Parses `openAfterImport` as JSON and normalizes boolean or nested-object payloads before building import options. Missing or malformed nested booleans now default to `false`.

Decision:
- Import should not fail at the Tauri argument decoding layer when a stale or malformed frontend payload passes an object where a boolean is expected. The command now reaches backend validation and returns controlled errors.

### Workspace Export Reveal Target

Modified:
- [DataManagement.tsx](../src-ui/src/views/DataManagement.tsx): After Workspace export, opens the generated archive's containing directory rather than passing the archive file path as the folder target.
- [workspace_maintenance.rs](../src-tauri/src/workspace_maintenance.rs): Adjusted Windows Explorer file reveal arguments to use `/select,PATH` without embedded quote characters.

Decision:
- Export completion should take users directly to the `.lamber.zip` output directory. File selection remains a backend capability, but the export action itself only needs to open the containing folder.

### Windows Workspace System Hidden Attributes

Modified:
- [workspace.rs](../src-tauri/src/workspace.rs): Added Windows hidden-attribute helpers and applies them to Workspace system entries when workspaces are inspected, opened, created, initialized, or imported.
- [workspace_maintenance.rs](../src-tauri/src/workspace_maintenance.rs): Marks `.backups`, `.exports`, and `.projects` hidden when maintenance flows create or repair those directories.
- [assets.rs](../src-tauri/src/project_files/assets.rs): Marks `.projects` hidden when template asset storage creates the internal asset sandbox.

Decision:
- Dot-prefixed names remain the portable workspace format, but Windows Explorer needs the Hidden attribute to hide those files when hidden items are disabled.

### Workspace Initialization Nonblocking Scan

Modified:
- [workspace.rs](../src-tauri/src/workspace.rs): Existing-folder Workspace initialization now drops the SQLite transaction guard and spawns project folder scanning / automatic Excel import as a background task after the Workspace is opened.
- [commands.rs](../src-tauri/src/project_files/commands.rs): Added a workspace-root/database-context helper for automatic Excel import so background initialization scans do not need to re-read the current runtime workspace.

Decision:
- Initialization must not block the modal on file scanning or Excel parsing. The Workspace opens after project records are created; scan/import remains best-effort and runs separately.

### Workspace Import Flat Arguments

Modified:
- [workspace_maintenance.rs](../src-tauri/src/workspace_maintenance.rs): `import_workspace` now accepts flat Tauri IPC arguments (`openAfterImport`, `conflictStrategy`, `destinationName`) and no longer parses a nested `options` value.
- [workspaceMaintenanceService.ts](../src-ui/src/services/workspaceMaintenanceService.ts): Flattens import options before invoking the backend command instead of sending an `options` object.
- [workspace_maintenance.rs](../src-tauri/src/workspace_maintenance.rs): Import destination resolution now treats the selected folder as the parent directory and creates `{selectedFolder}/{workspaceName}` by default.

Decision:
- The target folder selected during import is not the workspace root itself unless explicitly named through `destinationName`; it is the directory into which the imported workspace folder is created.
- Because Lamber has not been publicly released, old import IPC compatibility is intentionally removed to keep the command contract simple.

### Workspace Backup List Cleanup UI

Modified:
- [DataManagement.tsx](../src-ui/src/views/DataManagement.tsx): Added single-backup deletion and full-list clear actions to the Workspace backup list, backed by the existing `delete_workspace_backup` command.

Decision:
- Backup cleanup is a file-level maintenance action for `.backups` entries only. It does not modify or compact the active Workspace database.

### Standalone Hub Module Cleanup

Removed:
- [BenefitTool.tsx](../src-ui/src/views/BenefitTool.tsx): Removed the standalone Investment Benefit Analysis view.
- [DocfillTool.tsx](../src-ui/src/views/DocfillTool.tsx): Removed the standalone Document Material Production view.

Modified:
- [App.tsx](../src-ui/src/App.tsx): Removed Hub cards and routes for the retired standalone modules.
- [useNavigationStore.ts](../src-ui/src/store/useNavigationStore.ts): Reduced `ViewType` to `hub`, `project_board`, `ict_lifecycle`, and `data_management`.
- [aiContextKeys.ts](../src-ui/src/utils/aiContextKeys.ts), [aiContextSerializer.ts](../src-ui/src/utils/aiContextSerializer.ts), and [AiChatPanel.tsx](../src-ui/src/components/ai/AiChatPanel.tsx): Removed active AI scopes and quick actions for the retired modules while filtering legacy stored scopes.
- [main.rs](../src-tauri/src/main.rs): Unregistered standalone benefit batch/template and document-generation commands.
- [calculator.rs](../src-tauri/src/benefit/calculator.rs), [excel.rs](../src-tauri/src/benefit/excel.rs), [models.rs](../src-tauri/src/benefit/models.rs), and [docfill.rs](../src-tauri/src/docfill.rs): Removed code used only by the standalone modules while keeping shared ICT lifecycle calculation, Excel import, template listing, and lifecycle document generation code.
- [workspace_maintenance.rs](../src-tauri/src/workspace_maintenance.rs): Ignores stale app-level module paths for retired `benefit_tool` and `docfill_tool` when listing external workspace paths.
- [Cargo.toml](../src-tauri/Cargo.toml): Removed the no-longer-used `rust_xlsxwriter` dependency.

Decision:
- Retire only the two Hub modules requested by the user. Project Board, ICT Lifecycle, Data Management, template forms, Excel import, benefit schemes/snapshots, and shared document generation remain in scope and active.

### Workspace Health Repair Path Fallback Fix

Modified:
- [workspace_maintenance.rs](../src-tauri/src/workspace_maintenance.rs): Regenerating missing `project.json` now resolves the project directory using `relative_path`, `linked_folder_relative_path`, `folder_path`, then `folder_name`, matching the health-check path resolution used to report the missing manifest.
- [workspace_maintenance.rs](../src-tauri/src/workspace_maintenance.rs): External `module_path:*` health items are now repairable by resetting the module base path to `.projects/modules/{moduleId}` inside the current Workspace and updating the app-level module config.

Decision:
- Older or imported flat workspace projects that only have `folder_path` should remain repairable without requiring a separate path-conversion step first.
- Repairing a module path does not move or delete files from the old external path; it only changes where the module reads/writes templates and output going forward.

## 2026-05-29

### Workspace Refactoring Phase 4: Local Portability, Backup, Restore, and Health

Created:
- [workspace_maintenance.rs](../src-tauri/src/workspace_maintenance.rs): Workspace maintenance commands for daily/manual SQLite backups, backup listing/deletion/restore, `.lamber.zip` export/import/validation, read-only workspace health checks, repairable issue execution, external path listing, dry-run internal absolute path conversion, and file-manager reveal.
- [workspaceMaintenanceService.ts](../src-ui/src/services/workspaceMaintenanceService.ts): Frontend service wrapper for workspace maintenance IPC commands.

Modified:
- [workspace.rs](../src-tauri/src/workspace.rs): Exposed workspace manifest/database path helpers for maintenance flows, added reserved workspace entry detection, added safe database connection closing for restore, changed recent workspace updates to support import-without-open, and triggers a best-effort daily SQLite backup when a workspace opens.
- [commands.rs](../src-tauri/src/benefit/commands.rs): Rejects new flat workspace project folders that collide with reserved workspace entries and skips reserved entries during root-level workspace inspection.
- [assets.rs](../src-tauri/src/project_files/assets.rs): Resolves project folders relative to the active workspace for template asset storage and retrieval, writes new internal assets inside the current workspace, and keeps AppData lookup only as a legacy read fallback.
- [DataManagement.tsx](../src-ui/src/views/DataManagement.tsx): Added a Workspace Maintenance tab for current workspace metadata, manual backup, backup restore, export/import, health check results, repair buttons, external path listing, and internal path conversion.
- [WorkspaceGate.tsx](../src-ui/src/components/workspace/WorkspaceGate.tsx): Added `.lamber.zip` import entry when no workspace is active.
- [main.rs](../src-tauri/src/main.rs): Registered workspace maintenance commands.

Decisions:
- Direct folder copy remains the primary workspace migration path. Export/import is a convenience layer and preserves `workspaceId` by default.
- `.lamber.zip` archives use the workspace root as the archive root and include `export-manifest.json`; no random top-level folder is added. Import still accepts archives with one wrapper folder for compatibility.
- `.backups` and `.exports` are excluded from export by default. The pre-export database backup protects the live workspace and is not itself included unless backups are explicitly selected.
- `run_workspace_health_check` is read-only. Repairs and internal absolute path conversion are explicit user actions and create a database backup before modification.
- Backup restore releases the active SQLite connection before replacing `.lamber.sqlite`, then reopens the workspace and attempts rollback/reopen if replacement fails.
- `project_roots` representing external roots are only checked and reported; automatic path conversion does not rewrite them.

## 2026-05-28

### Workspace Refactoring Phase 3: Domain Save Boundaries and Dirty State

Created:
- [project_state/mod.rs](../src-tauri/src/project_state/mod.rs): Workspace-scoped project state commands for project detail, lifecycle state, cashflow state, benefit analysis, template states, template assets listing, and full project state loading.
- [useSaveStore.ts](../src-ui/src/store/useSaveStore.ts): Global dirty scope store with registered save handlers, context checks, partial failure handling, and Ctrl/Command+S integration.
- [domainSaveService.ts](../src-ui/src/services/domainSaveService.ts): Frontend domain service wrapping project detail, lifecycle, cashflow, benefit analysis, template state, and full-state commands.
- [GlobalSaveButton.tsx](../src-ui/src/components/GlobalSaveButton.tsx), [useGlobalSaveShortcut.ts](../src-ui/src/hooks/useGlobalSaveShortcut.ts), and [useUnsavedChangesGuard.ts](../src-ui/src/hooks/useUnsavedChangesGuard.ts).

Modified:
- [db.rs](../src-tauri/src/db.rs): Added schema version 5 with `project_lifecycle_states`, `project_cashflow_states`, and `project_template_states`; normalized `project_template_assets` creation for fresh databases and added `template_id` compatibility.
- [main.rs](../src-tauri/src/main.rs): Registered project-state Tauri commands.
- [service.rs](../src-tauri/src/benefit/service.rs): Prevented `update_project` from inserting a missing project into the current workspace database.
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Registered lifecycle and cashflow save handlers, loaded new project full state with legacy snapshot fallback, displayed unsaved status, and guarded project/template navigation.
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Routed template form persistence through the template domain state while preserving legacy fallback and asset references.
- [ProjectBoard.tsx](../src-ui/src/views/ProjectBoard.tsx), [WorkspaceGate.tsx](../src-ui/src/components/workspace/WorkspaceGate.tsx), and [WorkspaceHeader.tsx](../src-ui/src/components/WorkspaceHeader.tsx): Integrated global save, project/workspace switching guards, and project-detail dirty handling.

Decisions:
- Current ICT editing state no longer depends on `benefit_snapshots` as its only persistence boundary.
- Benefit方案 buttons still save schemes and snapshots, but global save / Ctrl+S save current editor state through domain handlers.
- Template field values and template assets remain separate from lifecycle/cashflow state; legacy `project_settings` template payloads remain readable and are kept as compatibility mirrors during saves.

Phase 3B updates:
- `useSaveStore` save handlers now return explicit `savedScopes`; missing handlers, partial failures, and workspace/project switches no longer clear unrelated dirty state.
- `TemplateForms` propagates template save failures to the global save handler so `template-forms` cannot be marked saved unless `saveTemplateState` succeeds.
- ICT lifecycle header shows project status and save status globally, while project, workspace, and template switches are guarded before navigation.
- ICT lifecycle project loading now prefers `project_lifecycle_states` current editor state over benefit scheme snapshots even when a default scheme id is present; snapshots remain a fallback for old projects.

### Workspace Structure Flattening, Hidden System Files, and Auto Migration

Modified:
- [workspace.rs](../src-tauri/src/workspace.rs): Changed manifest and database filenames to `.lamber.workspace.json` and `.lamber.sqlite` (prefixed with dot to make them hidden files on macOS/Linux). Changed backups and exports folder names to `.backups` and `.exports`. Removed the projects subfolder layer to place all project subdirectories directly at the workspace root. Added `migrate_legacy_workspace_files` to automatically rename existing visible files/folders (`lamber.workspace.json`, `lamber.sqlite`, `backups`, `exports`) to dot-prefixed hidden ones on workspace inspection and loading.
- [commands.rs](../src-tauri/src/benefit/commands.rs): Adjusted `create_project_in_workspace` and `inspect_workspace_projects` to create and scan folders directly under the workspace root, and to ignore hidden dot-prefixed system folders.
- [service.rs](../src-tauri/src/project_files/service.rs): Updated files creation and relative path references in `add_project_file` to skip the `projects/` subfolder.
- [assets.rs](../src-tauri/src/project_files/assets.rs): Changed sandboxed fallback files folder to `.projects/` (hidden) to keep the workspace root clean of database project IDs.

Decisions:
- Flattered workspace layout so that only the actual, user-visible project directories (e.g. `项目A`, `项目B`) are visible directly under the workspace root.
- Hidden all metadata, database, and cache files/directories behind Unix dot-prefixes (`.backups`, `.exports`, `.projects`, `.lamber.sqlite`, `.lamber.workspace.json`).
- Automatically detect and rename legacy visible workspace configuration files and directories when opened to ensure backward compatibility and seamless updates.

### Automatic Excel Calculations Import on Scan and Initialization

Modified:
- [commands.rs](../src-tauri/src/project_files/commands.rs): Implemented `auto_import_excel_if_needed` to search for files starting with "效益分析表" and ending with ".xlsx" / ".xls", sorting by modification date descending to pick the newest, and parsing/importing using the `ProjectService`. Updated `scan_project_folder` and `bind_project_folder` commands to trigger this check.
- [workspace.rs](../src-tauri/src/workspace.rs): Updated `initialize_workspace_from_existing_directory` to keep track of successfully imported project IDs, run directory scans, and call the auto-import logic for all projects after committing.

Decisions:
- Enforced a safety boundary: only trigger auto-import when the target project currently has 0 schemes to protect existing user work.
- Decoupled files service scanning from the benefit service calculations: `auto_import_excel_if_needed` resolves relative workspace paths against the active workspace root and passes them to the parsing module.

### Excel Importing Workspace Relative Path Resolution Fix

Modified:
- [excel.rs](../src-tauri/src/benefit/excel.rs): Updated the `parse_benefit_excel` command to accept the active `WorkspaceRuntime` state and resolve workspace-relative paths (e.g. scanned files in workspace directories) against the active workspace root path prior to performing existence checks and opening the workbook.

Decisions:
- Passed `State<'_, Arc<WorkspaceRuntime>>` to `parse_benefit_excel`.
- Checked and resolved non-absolute paths relative to the current workspace root path inside the command handler before spawning the blocking task to ensure proper lifetime handling of borrowed state objects.

### Database Migration and Workspace Initialization Robustness Fixes

Modified:
- [db.rs](../src-tauri/src/db.rs): Fixed a database migration bug on fresh databases. Previously, fresh databases were initialized with a schema version of `'2'`, which triggered the version 3 and 4 migrations (including `ALTER TABLE projects ADD COLUMN folder_name TEXT`) and failed with `duplicate column name: folder_name` since the table was already created using the latest schema. We added a column check on `projects` (`folder_name`) to detect fresh databases and set the initial schema version to `'4'` so migrations are skipped cleanly.
- [workspace.rs](../src-tauri/src/workspace.rs): Added robust error cleanups inside `initialize_workspace_from_existing_directory`. If any step fails during initialization (such as folder scanning or database transaction commits), the command deletes any partially created workspace manifest (`lamber.workspace.json`) and database (`lamber.sqlite`) to prevent leaving the folder in a corrupted, uninitializable state. It also automatically cleans up orphan manifests from failed runs on entry.

### Workspace Initialization & Subdirectories Bulk Import

Created/Modified:
- [workspace.rs](../src-tauri/src/workspace.rs): Implemented plain directories inspection logic, returning `importablePlainDirectory` when eligible subdirectories are found. Added options deserialization and `initialize_workspace_from_existing_directory` command which sets up directories, writes workspace manifest, establishes database transaction, deduplicates entries, writes/補足 `project.json` manifests and creates standard nested directories (`assets`, `documents`, `analyses`).
- [main.rs](../src-tauri/src/main.rs): Registered the `initialize_workspace_from_existing_directory` invoke handler.
- [workspaceService.ts](../src-ui/src/utils/workspaceService.ts): Added the status and new initialize command wrapper.
- [useWorkspaceStore.ts](../src-ui/src/store/useWorkspaceStore.ts): Exposed store method `initializeWorkspaceFromExisting`.
- [WorkspaceGate.tsx](../src-ui/src/components/workspace/WorkspaceGate.tsx): Unified workspace selection pickers under local inspection, rendering a checklist selection modal showing candidate project subdirectories and optional parameter options when opening plain directories.

Decisions:
- Standardized candidate directories to exclude backups, exports, standard projects, .git, and common target build directories.
- Stored project paths relative to workspaceRoot, keeping folders at the workspace root rather than moving them under `projects/`.
- Ensured transaction rollback and duplicates check to prevent workspace half-state or duplicate project database rows.

### Navigation Defaults and Back Button Adjustments

Modified:
- [useNavigationStore.ts](../src-ui/src/store/useNavigationStore.ts): Adjusted the initial view logic to always default `currentView` to `"hub"` upon application launch, avoiding restoring last active view directly (which bypasses the Hub and could immediately trigger a workspace gate or board view).
- [WorkspaceGate.tsx](../src-ui/src/components/workspace/WorkspaceGate.tsx): Allowed the back button to show even when there is no current workspace loaded (removed `currentWorkspace` check).
- [App.tsx](../src-ui/src/App.tsx): Passed the `onBack` handler and custom label to `WorkspaceGate` in the data management view, and wrapped the view in a header when `isWorkspaceReady` is false, ensuring the '返回集市' button is always consistently visible.
- [ProjectBoard.tsx](../src-ui/src/views/ProjectBoard.tsx): Passed the `onBack` handler and custom label to `WorkspaceGate` in the project board view when `isWorkspaceReady` is false. Standardized both header back buttons to use the text-button `← 返回集市` styling.
- [DataManagement.tsx](../src-ui/src/views/DataManagement.tsx): Standardized the header back button to the text-button `← 返回集市` style.

### Workspace Refactoring Phase 2: Project Board & Workspace Binding

Created:
- [useProjectStore.ts](../src-ui/src/store/useProjectStore.ts): Zustand store managing active currentProject context with local storage recovery and workspace linkage.

Modified:
- [workspace.rs](../src-tauri/src/workspace.rs): Added workspace paths normalization, verification functions, safe folder naming, standard subdirectories creation, and workspace root auto-healing logic on workspace relocation.
- [db.rs](../src-tauri/src/db.rs): Expanded `projects` table columns with migration check to version 4 (added `folder_name`, `relative_path`, `progress`, `deadline`, `linked_folder_type`, `linked_folder_relative_path`, `linked_folder_external_path`).
- [models.rs (benefit)](../src-tauri/src/benefit/models.rs): Extended Rust `Project` model with Phase 2 workspace mapping attributes.
- [repository.rs (benefit)](../src-tauri/src/benefit/repository.rs): Added DB mapping of extended project attributes for SQLite database repository.
- [commands.rs (benefit)](../src-tauri/src/benefit/commands.rs): Implemented workspace-scoped project operations: `create_project_in_workspace` (with directory structure creation and `project.json` manifest writing), `list_workspace_projects`, and `inspect_workspace_projects`.
- [service.rs (project_files)](../src-tauri/src/project_files/service.rs): Updated files scanning and sandboxing paths mapping to store documents relative to the workspace projects folder and to prevent delete-on-unbind behavior.
- [commands.rs (project_files)](../src-tauri/src/project_files/commands.rs) & [health.rs](../src-tauri/src/project_files/health.rs): Passed the active workspace root to file service calls.
- [ProjectBoard.tsx](../src-ui/src/views/ProjectBoard.tsx): Rebuilt creation flow using standard workspace paths, bound cards to missing directory warnings, removed old manual folder selection modals, and integrated project context selection.
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Synced global project state store context and added support for saving free calculations into any existing workspace project.

Decisions:
- Standardized project folders creation inside `workspaceRoot/{safeProjectName}/` with `assets/`, `documents/`, `analyses/` subfolders and a backup `project.json` manifest.
- SQLite remains the primary source of truth, with `project.json` serving as redundant metadata.
- Internal folders bound to projects are stored as relative workspace paths to guarantee portable workspaces, while external folders trigger visual alerts warning users of non-portable linkages.
- Unbinding local folders clears metadata and scanner references in the DB, but keeps actual disk files untouched.

### Workspace Runtime Foundation

Created:
- [workspace.rs](../src-tauri/src/workspace.rs): Workspace manifest model, UUID v4 workspace creation, current workspace runtime, database connection binding, recent workspace updates, last-opened restore, and directory-state inspection.
- [useWorkspaceStore.ts](../src-ui/src/store/useWorkspaceStore.ts): Frontend global workspace state.
- [WorkspaceGate.tsx](../src-ui/src/components/workspace/WorkspaceGate.tsx): Gate shown when database-backed modules are accessed without an open workspace.
- [workspaceService.ts](../src-ui/src/utils/workspaceService.ts): Frontend Workspace IPC wrapper and error parsing.

Modified:
- [main.rs](../src-tauri/src/main.rs): Removed startup initialization of AppData `projects_store.db`; registered Workspace commands; attempts to restore `lastOpenedWorkspacePath`.
- [config_manager.rs](../src-tauri/src/config_manager.rs): Added `recentWorkspaces` and `lastOpenedWorkspacePath` to local AppConfig.
- Project, file, root, health, relocation, import, template asset, and docfill commands now obtain the active SQLite connection from `WorkspaceRuntime`.
- [App.tsx](../src-ui/src/App.tsx): Keeps Hub focused on module selection and blocks database-backed non-board modules behind WorkspaceGate.
- [ProjectBoard.tsx](../src-ui/src/views/ProjectBoard.tsx): Moved workspace selection into the Project Board flow, making "Project Workspace" the first layer before project cards are loaded. Switching workspaces now opens a workspace overview instead of immediately opening a folder picker.
- [WorkspaceGate.tsx](../src-ui/src/components/workspace/WorkspaceGate.tsx): Refactored into a card-based workspace overview using locally recorded recent workspaces, with current-workspace marking and explicit open/create actions.
- [WorkspaceGate.tsx](../src-ui/src/components/workspace/WorkspaceGate.tsx) and [ProjectBoard.tsx](../src-ui/src/views/ProjectBoard.tsx): Clicking the current workspace card now directly closes the workspace overview and returns to the existing project board. Only selecting a different workspace performs a workspace switch.

Decisions:
- `lamber.workspace.json` stores only workspace identity metadata. Recent workspace history remains local to the machine.
- A directory with `lamber.sqlite` but no manifest is treated as `legacySuspected`; the app does not overwrite or migrate it in this phase.
- Opening or creating a workspace automatically registers the workspace root as a default project root when missing, so project folders under the workspace do not trigger the legacy "register as new root" prompt.
- Workspace selection belongs to the Project Board workflow. Hub should open the Project Board module, while ProjectBoard presents the Project Workspace layer before project list operations.
- Project workspace switching should show the recorded workspace overview first. Opening an arbitrary directory remains an explicit secondary action.
- Legacy JSON/AppData migration is deferred and no longer runs automatically on startup.

## 2026-05-27

### Project Background Persistence in Calculator Snapshots

Modified:
- [models.rs (benefit)](../src-tauri/src/benefit/models.rs): Added `project_background` field to `IctInput` struct.
- [useIctCalculations.ts](../src-ui/src/hooks/useIctCalculations.ts): Serialized `project_background` inside `buildInputDataPayload` to send it with the snapshot.
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Restored the `projectBackground` state inside `fillCalculatorState` when loading snapshot properties.

### Project Nested Assets Folder Renaming to Project Name Suffix

Modified:
- [assets.rs (project_files)](../src-tauri/src/project_files/assets.rs): Renamed the bound project folder's template asset folder from `.lamber` to `{project_name}-图片`. Added `sanitize_folder_name` helper to clean folder names, updated `get_project_folder_info_from_db` to retrieve project names, and implemented self-adaptive fallback path checks to support legacy `.lamber/assets/` images and project renames.

### Project Template Data & Asset Separation Persistence

Created:
- [assets.rs (project_files)](../src-tauri/src/project_files/assets.rs): Core sandboxed image uploads, size (<= 20MB) and MIME type checks, soft-delete, and orphan asset garbage collector.

Modified:
- [db.rs](../src-tauri/src/db.rs): Added `project_template_assets` table structure and transaction-wrapped Version 3 schema upgrade. Fix connection mutable borrow in init.
- [repository.rs (benefit)](../src-tauri/src/benefit/repository.rs): Added `get_project_setting` and `save_project_setting` to `ProjectRepository` and its Sqlite, Json, and Dual implementations.
- [commands.rs (project_files)](../src-tauri/src/project_files/commands.rs): Registered six new Tauri commands for managing project settings and assets.
- [main.rs](../src-tauri/src/main.rs): Registered the six new Tauri command handlers in the application runtime.
- [docfill.rs](../src-tauri/src/docfill.rs): Refactored `internal_generate_docx` to load sandboxed images by resolving `assetId` values using database connection validation, bypassing frontend absolute path leakage.
- [projectService.ts](../src-ui/src/utils/projectService.ts): Mapped backend template setting and asset commands to frontend APIs.
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Implemented loading, 1s auto-saving of forms, instant paste/drop upload, and legacy base64 image migrations. Fully refactored uncontrolled form fields (including demand checklist, customer confirm, env requirements, public URL, and security detail) using controlled state binding `getBind` to eliminate component-reload reset issues.
- [default.json (capabilities)](../src-tauri/capabilities/default.json): Removed the invalid `"core:protocol:allow-asset"` capability.
- [tauri.conf.json](../src-tauri/tauri.conf.json): Configured `assetProtocol`'s scope correctly for security sandbox file loading.

Decisions:
- Stripped base64 strings from JSON objects prior to writing to the SQLite database, storing only `assetId` references.
- Forced backend-only resolution of physical image paths using project database ownership validation.
- Allowed in-memory local base64 preview rendering in React frontend to ensure instant feedback while upload and saving happen in background.

## 2026-05-26


### Path Resilience & Database Cascade Deletion Fixes

Modified:
- [service.rs (project_files)](../src-tauri/src/project_files/service.rs):
  - Updated `bind_project_folder` to extract the parent directory of `folder_path` when creating a new project root, registering the parent directory as root, and mapping the subfolder name as the project's relative subpath.
  - Updated `scan_project_folder` and `add_project_file` to dynamically save the matched `project_directories` entry prior to saving files, and only set `directory_id` if a root is matched. If rootless (absolute-only mode), set `directory_id` to `None` to prevent SQLite `FOREIGN KEY constraint failed` errors.
- [import_scanner.rs](../src-tauri/src/project_files/import_scanner.rs):
  - Corrected `directory_id` assignment during project files insertion: if the imported project doesn't match any global root, set `directory_id` to `None` instead of `Some(dir_id)`. Removed unused `PathBuf` import.
- [benefit/repository.rs](../src-tauri/src/benefit/repository.rs):
  - Refactored `save_project` in `SqliteProjectRepository` to check row existence via `SELECT EXISTS` and run `UPDATE` or `INSERT` instead of `INSERT OR REPLACE`. This prevents SQLite's delete conflict resolution from triggering `ON DELETE CASCADE` deletions on child tables (`project_directories`, `project_files`, `benefit_schemes`, `benefit_snapshots`), which was previously wiping folder bindings and scheme calculations during Excel imports.

Decisions:
- Enforced strict nullable constraint safety for `directory_id` in `project_files` when directories are rootless, avoiding invalid foreign key references in SQLite.
- Adjusted root registration to logically target parent folders, enabling automatic grouping of adjacent project folders under the same root drive or parent directory.
- Replaced `INSERT OR REPLACE` with transactional `UPDATE`/`INSERT` checks on the main `projects` table to prevent unexpected cascade deletion cascades on child tables.

## 2026-05-25

### Project Files System and Path Resilience Upgrade (Phase 2)

Created:
- [roots.rs](../src-tauri/src/project_files/roots.rs) - Global Project Roots Configuration CRUD, defaults manager, and Tauri commands.
- [health.rs](../src-tauri/src/project_files/health.rs) - Health Check Service analyzing path linkages, exists states, mismatch counts, and auto-healing directories.
- [relocation.rs](../src-tauri/src/project_files/relocation.rs) - Transactional bulk relocation service to preview and swap directories paths across drives.
- [import_scanner.rs](../src-tauri/src/project_files/import_scanner.rs) - Recursive candidate folder scanner and importer utilizing SQL transactions to bulk import subfolders as projects (auto-determining file roles).
- [DataManagement.tsx](../src-ui/src/views/DataManagement.tsx) - Unified Data Management Dashboard supporting Roots, Health Checker, and Relocator.

Modified:
- [db.rs](../src-tauri/src/db.rs) - Updated schema versions and added new tables/migrations for project_roots and directories. Escaped `"exists"` column name to resolve SQLite syntax error.
- [repository.rs (project_files)](../src-tauri/src/project_files/repository.rs) - Extended repositories (Json, SQLite, Dual) with roots lookup, directories association, and files sync. Escaped `"exists"` column name in queries.
- [service.rs (project_files)](../src-tauri/src/project_files/service.rs) - Updated `bind_project_folder` to support `force_mode` parameter, added 5-level path resolving and metadata extraction, and integrated self-healing.
- [ProjectFilesTab.tsx](../src-ui/src/components/project/ProjectFilesTab.tsx) - Handled `NOT_IN_ROOT` error during folder binding by triggering a custom modal offering to create a root or bind as absolute-only.
- [ProjectBoard.tsx](../src-ui/src/views/ProjectBoard.tsx) - Added a "批量扫描导入" button in the board header, and implemented a custom modal showing candidate folders, file roles, and conflict resolution selector (`merge`/`new`/`skip`).
- [main.rs](../src-tauri/src/main.rs) - Registered new commands and services for roots, relocation, health, and scanner.
- [relocation.rs](../src-tauri/src/project_files/relocation.rs) & [import_scanner.rs](../src-tauri/src/project_files/import_scanner.rs) - Escaped `"exists"` column name in UPDATE and INSERT queries.

Decisions:
- Standardized relative path mapping using a combination of `project_roots` + `relative_path` + `file_fingerprint` to ensure path durability across different environments and drives.
- Enforced transactional integrity during batch operations (relocation, candidate imports) using SQLite database transactions to prevent corruption or locking.
- Escaped standard SQL reserved keyword `"exists"` within SQLite table structure and query strings to prevent runtime panics during boot.
- Retained "No-Line Rule" surface styling for all new components.

## 2026-05-23

### SQLite Database Integration and Dynamic Hot-Swapping (Phase 1)

Created:
- [db.rs](../src-tauri/src/db.rs) - SQLite connection management, foreign key configurations, and creation of 12 core tables.
- [migration.rs](../src-tauri/src/migration.rs) - JSON-to-SQLite transactional database migration service, auto-backup, report generator, and Tauri commands.

Modified:
- [Cargo.toml](../src-tauri/Cargo.toml) - Added `rusqlite` dependency with `bundled` feature.
- [repository.rs (benefit)](../src-tauri/src/benefit/repository.rs) - Implemented `SqliteProjectRepository` and Arc-wrapped `DualProjectRepository` wrapper for hot-swapping.
- [repository.rs (project_files)](../src-tauri/src/project_files/repository.rs) - Implemented `SqliteProjectFileRepository` and `DualProjectFileRepository` wrapper.
- [main.rs](../src-tauri/src/main.rs) - Registered new modules, initialized database connection, set up initial repository backends, and registered Tauri migration commands.
- [App.tsx](../src-ui/src/App.tsx) - Embedded startup check for SQLite database migration and created a premium ledger-style overlay popup modal to prompt user for migration and display statistics.

Decisions:
- Standardized SQLite database as primary storage, enabling transaction management instead of concurrency-sensitive JSON file writing.
- Kept `projects_store.json` backward compatibility and implemented automatic timestamped JSON backup before database inserts.
- Utilized dynamic dual repository swapping to support hot swapping repositories without app restarts after migration or skip actions.

### Initial project state mapping

Created:
- [AGENTS.md](../AGENTS.md)
- [AI_CONTEXT.md](./AI_CONTEXT.md)
- [PROJECT_STATUS.md](./PROJECT_STATUS.md)
- [ARCHITECTURE_MAP.md](./ARCHITECTURE_MAP.md)
- [CHANGELOG_AI.md](./CHANGELOG_AI.md)

Observed:
- The project is a desktop application combining a Rust/Tauri v2 backend and React 18 / Zustand frontend.
- Calculations (payback, NPV, selection fee, reverse-calculations) use Rust's `rust_decimal` to prevent float rounding errors.
- Real-time serialization of form states to HSL-formatted AI prompts allows context-aware chat capabilities.
- Local folder scanning maps folder structures into project file entities, with safety toggles between sandbox copies and raw links.
- 0-tolerance financial reconciliation blocks invalid user states during workflow transitions.

Decisions:
- Standardized UI borders are replaced by color surface adjustments to fulfill the design guidelines of "The Architectural Ledger".
- Navigational routes track entry sources (`entrySource`) to support bidirectional back-navigation without context degradation.
- AI direct write capabilities are locked behind user interaction boundaries.

Risks:
- Coordinates for Excel coordinate back-filling (`excel.rs` and `parse_benefit_excel`) are hardcoded. Edits to templates will break structural mappings.
- The 0-tolerance reconciliation check may experience rounding edge cases when converting user inputs.
- Concurrency during atomic JSON writes (`projects_store.json`) might result in race conditions.

Open questions:
- Should the local JSON-based storage layer be migrated to SQLite in the next phase to improve transactional consistency?
- Do we need to support dynamic annual discount rates rather than flat project-wide rates?
