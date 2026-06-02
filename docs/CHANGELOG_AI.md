# CHANGELOG_AI.md

This changelog records structural modifications, business rules, and context changes made by AI agents to maintain a reliable project state mapping.

## 2026-06-01

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
