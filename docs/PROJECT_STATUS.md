# PROJECT_STATUS.md

Last updated: 2026-06-01 (ICT Billing Subject Name Extension)

## 0. ICT Sign-off Project Situation Itemization

The ICT project sign-off form's project situation wording now expands investment and revenue details from the measurement subject catalog instead of relying on four manually edited IT/CT subject fields. Investment and revenue lines enumerate every non-zero standard subject by category using the measurement-table billing subject name when present, otherwise the standard subject name, prepend the fixed category prefix such as `IT-` or `CT-`, and keep tax-exclusive amounts unchanged.

The sign-off form no longer exposes manual "IT投入 / CT投入 / IT收入 / CT收入" billing-subject inputs. Its dedicated configuration order is Project Background, IT/CT Service Content, advance-payment and post-approval-selection checkboxes, then revenue collection and expenditure payment methods. The "立项后甄选" phrase is included in the investment wording only when the new checkbox is selected.

Meeting-review "项目整体投入金额" now reuses the same generated investment wording as the sign-off project situation investment line, keeping subject prefixes and billing-subject fallback consistent. This change affects document wording and template placeholders only. It does not change calculations, tax rates, cashflow, Excel row mapping, or investment amount calculations.

## 0. ICT Billing Subject Name Extension

ICT lifecycle revenue and cost measurement pages now support a second optional name field, "计费科目名称", for every fixed standard billing subject. The previous "具体业务/产品名称" field remains unchanged and is still saved as the product/business name; the new billing subject name is saved separately and is optional for old project data.

A shared subject presentation resolver now centralizes subject naming for UI display hints, Excel output variables, project sign-off wording, meeting-review wording, and document aggregation deduplication. The display priority is `计费科目名称 > 具体业务/产品名称 > 标准科目名称` for Excel and page subject labels. Document business names use the same priority before adding the existing category prefix (`IT-`, `CT-`, `非IT/CT-`, `综合类-`), while old fallback behavior remains in place when neither custom field is present.

The new field is persisted through existing lifecycle input payloads, benefit scheme snapshots, and cashflow assumptions JSON as `billing_subject_name` / `billingSubjectName`; no new SQLite table or schema migration is required. CT paired subjects (`产品收入` ↔ `其他产品成本`, `专线收入` ↔ `专线带宽成本`) keep billing subject names synchronized like their amount and product-name pass-through path.

This change does not add/delete standard subjects, does not add Excel rows, and does not alter G/Q amount writing, tax rates, cashflow, NPV, selection fee, reconciliation, or reverse-calculation formulas.

## 0. ICT Subject Custom Business Name Extension

ICT lifecycle revenue and cost measurement pages now support a "具体业务/产品名称" field for every built-in standard billing subject. The standard subject remains the fixed internal identity for calculation, mapping, Excel row targeting, and backward compatibility; the custom name is stored separately on the tax item and is optional for old project data.

Excel generation for `3-直接经济效益评估表` now writes each subject's display name as `标准科目（具体名称）` when a custom name exists, while keeping the standard name when it does not. The same row's `G` column is written with tax-exclusive amount and `Q` column with tax-inclusive amount for all mapped revenue/cost subjects, including `Q10` and `Q25`, through a single subject-row mapping.

Project sign-off and meeting-review document variables now derive custom document business names such as `IT-高楼消防监测平台集成服务` and `CT-视频监控` from the subject catalog. Business composition-style values are deduplicated across revenue and cost when the same custom business appears on both sides; amount detail variables still keep income and cost records separate.

Follow-up refinement: empty/zero frontend amounts are serialized as blank Excel amount inputs rather than `0`, so unused subject rows remain unfilled in generated benefit analysis sheets. CT product revenue and CT other-product cost custom names now stay synchronized in both directions, matching the existing amount pass-through behavior; CT line revenue and CT bandwidth cost names follow the same paired-subject synchronization.

This change does not add or delete billing subject rows, does not insert Excel rows, and does not alter tax, cashflow, NPV, selection fee, or reverse-calculation formulas.

## 0. Meeting Investment Subject Alignment Fix

ICT presales meeting-review document generation now builds the "项目整体投入金额" detail from the same IT/CT cost subject fields used by the project sign-off form. The meeting-review wording no longer derives CT investment subject names from the mid-platform capability name, while the existing tax-exclusive IT, CT, mixed-cost, and total investment calculations remain unchanged.

## 0. ICT Cashflow Price Persistence Hydration Fix

ICT lifecycle reload now hydrates calculator inputs from the latest cashflow domain state when available. Price edits in IT/CT revenue and cost fields are saved under `project_cashflow_states.assumptions_json`; reopening the project overlays those assumptions onto lifecycle input payloads or legacy scheme snapshots before filling the calculator. This prevents stale `project_lifecycle_states.input_payload_json` values from restoring old prices after a cashflow-only save.

## 0. Inquiry Vendor Image State Preservation Fix

The one-click three-vendor quote generator now preserves vendor quote screenshots instead of rebuilding every vendor row with an empty `images` array. Generated quote rows merge existing images by vendor name first and by row position as a fallback, so recalculating quote amounts does not discard uploaded screenshots.

Vendor image upload updates now use functional `setInqVendors` updates, preventing asynchronous file reads from writing image results into stale vendor arrays after quote rows are regenerated or edited.

## 0. Template Image Document Embedding Fix

Template image fields used during Word generation now serialize project template assets by `assetId`, not by frontend preview URLs such as `asset://localhost/...`. The backend image replacement path prefers `assetId` inside JSON image payloads, resolves the asset through Workspace SQLite ownership checks, embeds the binary image into the generated docx, and suppresses unresolved image JSON so raw payload text cannot leak into Word output.

## 0. Lifecycle Document Workspace Output Path Fix

ICT lifecycle document generation now resolves output folders through the active Workspace context before writing files. Workspace-relative project paths such as `项目A` are expanded against the current Workspace root instead of being treated as process-relative paths under `src-tauri`, preventing generated documents from landing inside the Tauri source tree and triggering dev-mode rebuild/restart loops.

The frontend passes the active `projectId` along with document-generation requests. The backend keeps explicit `outputDir` compatibility, but always resolves relative output paths with `workspace::resolve_workspace_path` and can derive the target project folder from the Workspace SQLite project record when no output directory is provided.

## 0. AI Workspace Specified Project Context Routing

The AI chat context composer can now resolve explicitly named projects from the currently opened Workspace before a message is sent. A new read-only Workspace project index command returns lightweight project identity and saved-state existence metadata from `.lamber.sqlite`, including project names, real `projectId` values, lifecycle/cashflow/template/benefit presence flags, and saved template names. It does not return project paths, file contents, image bytes, or full project/template JSON.

Project names are used only as routing hints within the current Workspace. After a unique deterministic match, all official context reads reuse `build_ai_project_context` with the matched real `projectId`. If the user names a different project while another project is currently open, the explicitly named project wins. If the user compares two explicitly named projects, the composer loads at most two separated official project contexts. Ambiguous duplicate names, unresolved named-project references, and overly broad multi-project requests degrade into warnings instead of guessing.

Specified project template reads reuse the existing `template_detail` path. The composer first loads a target project's saved summary, resolves an explicitly mentioned template name or known template alias such as "立项签批表", and only then requests that one saved template detail. Template image content remains protected: images are metadata-only unless the existing explicit template asset selection flow attaches a validated asset for the current turn.

Workspace-level questions such as "哪些项目已经填写了立项签批表" can use the lightweight Workspace project index without deep-loading every project. Current unsaved frontend draft overlay is still injected only for its bound active project and is never attached to another specified project.

## 0. AI Template Detail Context and Controlled Vision Assets

The AI chat context composer now loads a specified template's saved detail only when the active ICT page is a template editing context. The detail source is the current Workspace SQLite database through `build_ai_project_context` with `requestedSources: ["template_detail"]` and `activeTemplateId`; ordinary project questions continue to use the summary bundle.

Template detail context separates saved official content from unsaved draft overlay. Saved fields come from `project_template_states` with `project_settings` (`template_form_data::<templateId>`) as compatibility fallback. Current template-page edits remain an unsaved frontend draft overlay and must not be described as saved or written.

Template detail sanitization strips data URLs, base64/preview fields, local absolute paths, and oversized nested payloads before prompt injection. Image fields expose only lightweight asset metadata by default.

Template images can be sent to AI only through explicit user selection from a template image field. The frontend passes only `projectId`, `templateId`, and `assetId`; the backend `load_ai_template_asset` command validates that the asset belongs to the current project in the active Workspace, checks supported image MIME/size constraints, resolves only workspace-contained asset files, and returns a temporary data URL for the current vision request. The AI still cannot write template fields, auto-fill forms, run RAG/embedding, summarize project documents, or read unselected images.

## 0. AI Project Context Chat Integration

The AI chat send path now composes project context at message-send time. For active project-aware views, it reads the saved official project context through the backend `build_ai_project_context` command and injects it into the prompt as Workspace SQLite persisted state. The existing frontend AI context store remains in use only as a current-page unsaved draft overlay, and the overlay is injected only when matching dirty scopes exist for the current project/page.

The message-send composer refreshes active project identity from the existing local navigation/current-project storage, including the ICT lifecycle active project key, before deciding whether to call `build_ai_project_context`. This keeps the separate floating AI window aligned when it was opened before an ICT project was selected or when the user switches projects while the AI window remains open. Only project identity is refreshed this way; saved official project state still comes from Workspace SQLite.

Workspace switching now clears the workspace-scoped active project and scheme identity across the project store, navigation store, and legacy ICT local storage keys. This prevents a floating AI window from sending a stale project ID from the previous workspace to the current workspace's SQLite context command.

Project Board also publishes a lightweight current-workspace board summary to the AI context store, including workspace identity, project count, and compact project cards. The chat composer injects this as current page context for Project Board questions such as workspace project counts, while project-level official detail still comes from `build_ai_project_context` only when a valid active project is selected.

The prompt now separates saved official state from current unsaved changes. If backend project context loading fails, chat streaming still proceeds with a warning context node instead of pretending that the project has empty data. The integration remains read-only: it does not trigger saves, scans, repairs, document generation, financial recalculation, AI writes, patch application, RAG, embeddings, file full-text summaries, or image binary analysis.

The chat streaming UI resets the parser state before each new assistant placeholder is inserted, and stream synchronization overwrites the active placeholder with the current parser output instead of falling back to previous content. This prevents the previous assistant response from flashing inside the new response while the model is thinking.

Current context volume control uses the backend summary mode by default. Draft overlay data is scoped to the active project, sanitized to remove base64/data URL previews and absolute paths, and truncated for large arrays/objects/strings.

## 0. AI Project Context Service

The backend now exposes a read-only `build_ai_project_context` Tauri command for project-level AI context retrieval. It obtains the active SQLite connection only through `WorkspaceRuntime`, validates that the requested project exists in the current Workspace database, and returns a structured summary bundle for `overview`, `lifecycle`, `cashflow`, `benefit`, `templates`, and `files`.

The service reads only persisted official state from `.lamber.sqlite`; it does not read frontend localStorage, Zustand draft state, or page-local unsaved data. It does not write database rows, scan folders, repair paths, read document full text, read image binaries, return absolute asset paths, perform RAG/embedding work, or inject the result into the chat prompt. Frontend support is currently limited to the typed `aiProjectContextService.ts` invoke wrapper for later integration.

## 0. Project Background, Collection/Payment and IT/CT Content Sync in Template Forms

The Project Background, Collection and Payment Methods, and IT/CT Service Content inputs have been added and synchronized inside the Project Sign-off Form template configuration (`TemplateForms.tsx`). The fields are reactively shared, showing editable default value placeholders, and are fully saved/persisted via template state handlers in the workspace database.

## 0.00 Latest global workspace management

The Data Management Center now includes a global "associated workspaces" list backed by local `recentWorkspaces`. Users can open, reveal, or unlink any remembered Workspace from this list. Unlinking only removes the local association from `config.json`; it never deletes the Workspace folder, `.lamber.sqlite`, or project files on disk.

If the unlinked Workspace is currently open, the frontend runs the unsaved-change guard first, then the backend clears the active `WorkspaceRuntime` and database connection and removes `lastOpenedWorkspacePath`. Non-current Workspace unlinking only updates the recent list.

The associated Workspace list is now separated into its own "Workspace Management" tab in Data Management and uses the same card-based visual pattern as the Project Board workspace picker. Current-workspace maintenance actions remain in the separate "Workspace Maintenance" tab.

Workspace Management cards behave as local selection targets: clicking a card only highlights it and must not open the Workspace or reorder the recent Workspace list. Opening, revealing, and unlinking remain explicit inline button actions. The inline "open" action enters the selected Workspace's Project Board after the Workspace is active.

## 0.1 Latest workspace import argument guard

Workspace archive import now calls the frontend maintenance service with explicit arguments (`zipPath`, `targetDir`, `openAfterImport`, `conflictStrategy`, `destinationName`) instead of passing an options object through call sites. The Rust command also receives `openAfterImport` as JSON and normalizes either a boolean or an accidentally nested map before constructing `ImportWorkspaceOptions`; missing or malformed nested booleans default to `false` so import is not blocked before archive validation runs.

## 0.2 Latest workspace export reveal target

After a Workspace export completes, the Data Management Center opens the exported archive's containing directory instead of sending the archive path directly as the folder target. The backend file-manager reveal command also uses a simpler Windows Explorer `/select,PATH` argument format for file selection, avoiding failures that can fall back to the desktop or computer root view.

## 0.3 Latest Windows workspace system hidden attributes

Workspace system files and directories now use Windows Hidden file attributes in addition to dot-prefixed names. Opening, creating, initializing, importing, backing up, exporting, repairing, or creating `.projects` asset storage marks `.lamber.workspace.json`, `.lamber.sqlite`, `.backups`, `.exports`, and `.projects` as hidden on Windows, while keeping macOS/Linux behavior based on dot prefixes.

## 0.4 Latest workspace initialization nonblocking scan

Initializing an existing plain directory as a Workspace now commits project rows, releases the SQLite transaction lock, and returns the opened Workspace before project folder scanning and automatic Excel calculation import run in a background task. This prevents the initialization dialog and recent workspace cards from staying disabled when folder scanning or Excel parsing touches large user files.

## 0.5 Latest workspace import flat arguments

Workspace archive import now uses flat Tauri IPC arguments: `openAfterImport`, `conflictStrategy`, and `destinationName`. The frontend no longer sends a nested `options` object, and the backend no longer keeps the legacy boolean `options` compatibility path because the app has not been publicly released.

The selected import folder is treated as the parent destination. By default, importing creates `{selectedFolder}/{workspaceName}` and then handles conflicts using the configured strategy, rather than extracting `.lamber.workspace.json` and `.lamber.sqlite` directly into the selected folder.

## 0.6 Latest workspace backup cleanup UI

The Data Management Center backup list now supports deleting a single SQLite backup and clearing all backups currently shown in the list. These actions call the existing workspace backup deletion command and only remove backup files under the workspace backup area; they do not modify the active `.lamber.sqlite` database.

## 0.7 Latest standalone module cleanup

The Hub no longer exposes the standalone "Investment Benefit Analysis" and "Document Material Production" modules. Their dedicated frontend views, navigation routes, AI active scopes, quick actions, and standalone Tauri commands were removed.

The shared ICT lifecycle calculation engine, project-board Excel import path, template-form document generation path, `parse_benefit_excel`, `generate_lifecycle_docs`, and `get_available_templates` remain active because they are used by Project Board and ICT Lifecycle workflows.

Legacy app-level module path records for retired `benefit_tool` and `docfill_tool` modules are ignored by workspace external-path reporting so stale paths do not appear as repairable workspace health issues.

## 0.8 Latest workspace portability update

Workspace refactoring phase 4 adds local portability maintenance for flat Lamber Workspace roots. A workspace remains copyable as a folder: project directories live directly under `workspaceRoot/{safeProjectName}/`, while system entries use reserved names such as `.lamber.workspace.json`, `.lamber.sqlite`, `.backups`, `.exports`, `.projects`, `backups`, `exports`, and `projects`. New project creation and workspace health checks must reject or flag project folders that collide with those reserved entries.

The backend now includes workspace maintenance commands for daily/manual SQLite backups, safe backup restore, read-only health checks, repair actions, external path inspection, dry-run path conversion, `.lamber.zip` export, `.lamber.zip` validation/import, and file-manager reveal. Export archives use the workspace root as the zip root, preserve `workspaceId`, include `.lamber.workspace.json`, a consistent `.lamber.sqlite` copy, `.projects/`, project folders, and `export-manifest.json`, and exclude `.backups` / `.exports` by default unless explicitly requested. Import preserves `workspaceId` as a migration/restore of the same workspace and supports both direct-root archives and archives with one top-level wrapper directory.

Health checking is read-only. It reports missing system/project directories, reserved-name conflicts, absolute internal paths, external paths, missing template assets, unregistered project manifests, database integrity/schema issues, and orphan or malformed records across `project_lifecycle_states`, `project_cashflow_states`, `project_template_states`, `project_template_assets`, `benefit_schemes`, and `benefit_snapshots`. Repair commands are separate, require explicit user action, and create a database backup before modifying files or data.

Project manifest repair uses the same directory resolution fallback as health checking: `relative_path`, `linked_folder_relative_path`, `folder_path`, then `folder_name`. This keeps older or imported flat workspace projects repairable even when `relative_path` has not been backfilled yet.

Module workspace paths (`module_path:*`) that point outside the active Workspace are repairable by resetting the module base directory to `.projects/modules/{moduleId}` inside the current Workspace and creating `templates/` and `output/`. This repair changes the app-level module configuration but does not move or delete files from the old external directory.

Template asset storage now resolves current workspace-relative paths first. New internal project assets are written inside the active workspace, with `.projects/{projectId}/assets` used as the fallback for external project folders. Legacy AppData asset lookup is kept only as a read fallback, not as the preferred write location.

## 0.9 Latest persistence update

Workspace refactoring phase 3 separates current project editing state from scheme snapshots. Project detail updates remain in `projects`; ICT lifecycle editor state is saved in `project_lifecycle_states`; funding models, assumptions, cashflow tables, sector cashflows, and metrics are saved in `project_cashflow_states`; benefit方案 history remains in `benefit_schemes` / `benefit_snapshots`; template field values, mappings, and output configuration are saved in `project_template_states` with compatibility fallback to legacy `project_settings` keys. Template binary/image assets continue to use `project_template_assets` and file-backed storage.

The frontend now has `useSaveStore`, a domain save service, a global save button, Ctrl/Command+S handling, and unsaved-change guards. Dirty scopes are cleared only after their registered domain save handler returns the scopes it actually persisted for the same workspace and project context. Template form saving must propagate failures to `useSaveStore`; autosave may be silent, but global save must keep `template-forms` dirty if `saveTemplateState` fails.

Opening an ICT project now restores current lifecycle editor state from `project_lifecycle_states` before reading benefit scheme snapshots, even when navigation includes a default scheme id. Scheme snapshots remain historical/fallback data, not the primary source for current project background or lifecycle edits.

## 1. Project summary

Lamber is a lightweight sales support and project management desktop tool designed for client managers and solution experts in the 5G/ICT domains. It addresses the inefficiencies of manual financial calculations, document generation, and project folder tracking by consolidating project board management, ICT lifecycle and cashflow assessment, workspace-backed template filling, Excel parsing/import, and local file scanning into a single Tauri-based application.

## 2. Tech stack

- **Frontend**: Vite + React 18 + TypeScript + Zustand
- **Desktop runtime**: Tauri v2 + Rust
- **State management**:
  - Global navigation: `useNavigationStore` (Zustand + local storage, always defaulting to "hub" view on application startup)
  - AI Context: `useAiContextStore` (Zustand + local storage + Tauri events)
  - Layout & view modes: Local storage (`lamber_project_board_view_mode`, etc.)
- **Database/Persistence**: SQLite (via `rusqlite` with `bundled` feature) for structured relational tables, alongside dynamic `projects_store.json` backward compatibility
- **Workspace Runtime**: Business data is now scoped to an explicit Lamber Workspace root containing hidden system files `.lamber.workspace.json`, `.lamber.sqlite`, `.backups/`, `.exports/`, and `.projects/`. Project folders (e.g. `项目A`, `项目B`) are placed directly inside the workspace root without an intermediate `projects/` layer. The app remembers `recentWorkspaces` and `lastOpenedWorkspacePath` in local AppConfig, with recent workspace records deduplicated primarily by path. Data Management exposes this local association list so users can open, reveal, or unlink remembered workspaces without deleting the physical workspace folders. It supports initializing an existing general project root directory as a Lamber Workspace and bulk importing eligible first-level subdirectories as workspace internal projects (with automatic creation of `project.json` and assets/documents/analyses folders). Accessing workspace-backed features without a workspace redirects to `WorkspaceGate` with matching headers, which now provides a standardized '← 返回集市' back button to return to the Hub (across all entry views including Project Board and Data Management, both when workspace is active or inactive).
- **Styling**: Tailwind CSS + Shadcn/UI (Radix UI) + HSL-based design system
- **AI integration**: Local SSE streaming client (Ollama / OpenAI standard endpoint) with semantic Markdown context serialization
- **File handling**:
  - Word variable replacement: Backend Rust `docx-template`
  - Excel template filling/parsing: Backend Rust `calamine` + `umya-spreadsheet`
  - Local directory opening: Native OS process spawning (`explorer` on Windows, `open` on macOS)
- **Build tools**: npm + cargo (Tauri CLI)

## 3. Core modules

### 3.1 Project Board (项目看板)

**Status**: Active (Fully Implemented)

- **Current behavior**: Displays projects scoped to the active Lamber Workspace root. It fetches project details from the current workspace SQLite database (`.lamber.sqlite`). New projects are automatically created under `workspaceRoot/{safeProjectName}/` with standard assets, documents, and analyses subdirectories and a redundant `project.json` manifest. Cards indicate directory health, flagging missing directories in the UI. Binding works with relative workspace paths for internal folders, and warns users when folders are external. Disconnecting folder links clears database records but keeps real disk files intact.
- **Known requirements**: Maintain independent rendering from ICT Lifecycle view and persist UI selections in local storage.
- **Known issues**: Large note sizes might slightly lag during immediate auto-save.
- **Related files**:
  - [ProjectBoard.tsx](../src-ui/src/views/ProjectBoard.tsx)
  - [projectService.ts](../src-ui/src/utils/projectService.ts)

### 3.2 ICT Lifecycle Calculator (ICT生命周期测算)

**Status**: Active (Fully Implemented)

- **Current behavior**: Computes 10-year cashflows, NPV, NPV Rate, Margin Rate, IRR, and payback period. Supports bound-project mode and standalone "free" calculator mode. Includes quick split calculators (requiring 1% own product revenue allocation), selection fee calculators, and smart back-calculations. The project background, discount rates, property rights, and cashflow details are fully persisted in the scheme snapshots. It enforces a 0-tolerance tax validation check before allowing users to see cashflows or generate documents.
- **Known requirements**: Maintain risk analysis criteria defined in backend Rust code.
- **Known issues**: Binary search limit for back-calculation is capped at 10 billion CNY.
- **Related files**:
  - [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx)
  - [calculator.rs](../src-tauri/src/benefit/calculator.rs)

### 3.3 AI Assistant (Lamber 智能顾问)

**Status**: Active (Fully Implemented)

- **Current behavior**: SSE stream chatbot utilizing a local Ollama server (or OpenAI compatible API). Automatically gathers frontend serialized business states and injects them as system prompts, together with a product capabilities catalog (`midThreeConstants.ts`) to make suggestions.
- **Known requirements**: Multi-round history optimization and PDF/Markdown local RAG.
- **Known issues**: High token consumption on large templates since the entire form structure is serialized.
- **Related files**:
  - [AiChatPanel.tsx](../src-ui/src/components/ai/AiChatPanel.tsx)
  - [useAiContextStore.ts](../src-ui/src/store/useAiContextStore.ts)
  - [AiRuntime.ts](../src-ui/src/ai/AiRuntime.ts)

### 3.4 File / Excel / Word Integration (文件及模板集成)

**Status**: Active (Fully Implemented)

- **Current behavior**: Scans directories for docx/xlsx files. Back-fills form fields (from `TemplateForms.tsx`) into files via Rust backend. Word templates are generated via `docx-template`. Excels can be filled or parsed back (via `parse_benefit_excel`, which resolves workspace-relative paths against the active workspace root) to overwrite project states. Supports sandbox "copied" files or raw disk "linked" files. Template forms (`TemplateForms.tsx`) and embedded image resources are persisted separately: lightweight forms/table settings are saved in `project_settings` (under key `template_form_data::<template_name>`), while pasted/dropped images are uploaded instantly to project assets sandboxes (bound project folder under `{project_name}-图片/assets/` if linked, otherwise falling back to `.projects/{project_id}/assets/` inside the active workspace) and tracked in the `project_template_assets` metadata table. Document generation reads binary contents directly from sandbox files via backend validation. Automatically imports Excel calculations: when a folder scan is triggered (manually, via folder binding, or during workspace project bulk initialization), if the project has 0 schemes, it filters for Excel files whose names start with "效益分析表" and end with ".xlsx"/".xls", chooses the newest one by modification date, parses its economic parameters, and saves it as the default scheme "Excel导入测算方案".
- **Known requirements**: Scan timestamps are updated without modifying physical files.
- **Known issues**: Cell coordinates mapping in Excel template is fragile if the spreadsheet structure changes.
- **Related files**:
  - [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx)
  - [docfill.rs](../src-tauri/src/docfill.rs)
  - [project_files/service.rs](../src-tauri/src/project_files/service.rs)

### 3.5 Data Management Center & Path Resilience (数据管理中心与路径韧性)

**Status**: Active (Workspace Maintenance extended in Phase 4)

- **Current behavior**: Avoids hardcoded absolute paths by employing global project roots (`project_roots`), relative paths (`project_directories`), and file fingerprints (`size:modified_at:first_8kb_hash`). Supports global project roots CRUD, folder binding warning triggers (offering auto-parent folder registration as root with relative project subpaths, or absolute-only paths), a file-link Health Checks Dashboard, bulk path relocation, and candidate scanner to import folders as new projects. Phase 4 adds a card-based Workspace Management tab for associated Workspace records, plus a separate Workspace Maintenance tab for current workspace information, manual SQLite backup, backup listing/restore, `.lamber.zip` export/import, read-only workspace health checks, repairable issue actions, external path listing, and dry-run conversion of internal absolute paths to workspace-relative paths. Allows scanning and syncing of rootless project directories safely without triggering foreign key constraint failures.
- **Related files**:
  - [DataManagement.tsx](../src-ui/src/views/DataManagement.tsx)
  - [ProjectFilesTab.tsx](../src-ui/src/components/project/ProjectFilesTab.tsx)
  - [ProjectBoard.tsx](../src-ui/src/views/ProjectBoard.tsx)
  - [project_files/roots.rs](../src-tauri/src/project_files/roots.rs)
  - [project_files/health.rs](../src-tauri/src/project_files/health.rs)
  - [project_files/relocation.rs](../src-tauri/src/project_files/relocation.rs)
  - [project_files/import_scanner.rs](../src-tauri/src/project_files/import_scanner.rs)
  - [workspace_maintenance.rs](../src-tauri/src/workspace_maintenance.rs)

## 4. Important business rules

- **Project Board Independence**: The Project Board is independent of the ICT Lifecycle. Navigating to the calculator tracks the origin via `entrySource` so the back button knows where to return.
- **1% Own-Product Rule**: Projects must integrate at least 1% of total revenue as own-product CT revenue. The quick split calculator helps compute and fill this automatically.
- **0-Tolerance Tax Reconciliation**: In the calculator, tax-inclusive amount must match tax-exclusive + tax amount (`incl = excl * (1 + rate)`). If there is a rounding mismatch, navigation is blocked. If the error is <= 0.10, the user can override it.
- **Backend Risk Assessment**: Risk levels (高风险, 中风险, 低风险) are strictly evaluated by the Rust backend service `calculate_risk_level` in `service.rs` based on thresholds (margin: 8%; NPV rate: 4%).
- **Selection Fee Brackets**: Spanning quote ranges determine selection fees (e.g. quote <= 12,100 -> fee is 100; quote <= 48,500 -> fee is quote * 0.00825). Highest limit can also be reverse-calculated.
- **AI Direct Write Lock**: AI is allowed to read and analyze any active workspace variables, but is prohibited from updating the database files directly. Binders or "apply" buttons are the only allowed write mechanisms.

## 5. Important architecture decisions

### ADR-006: Workspace-Scoped Primary Database
- **Decision**: Introduce a Lamber Workspace root as the owner of the primary SQLite database. Project, file, roots, relocation, import, template asset, and document image resolution commands must require an opened workspace and obtain the active `lamber.sqlite` through `WorkspaceRuntime`.
- **Reason**: Avoid scattering the primary database in the OS AppData directory and prepare the app for portable project workspaces.
- **Impact**: Startup no longer initializes `AppData/projects_store.db`. If last workspace auto-restore fails, the app shows a WorkspaceGate and does not create an empty database or fall back to AppData. Legacy AppData/JSON migration is deferred to a compatibility phase.

### ADR-007: Workspace Portability Maintenance
- **Decision**: Treat direct folder copy as the primary migration path and `.lamber.zip` export/import as a convenience format. Workspace archives preserve the same `workspaceId` by default, use the workspace root as the archive root, and keep internal paths relative to the active workspace wherever possible.
- **Reason**: Users should be able to copy a workspace folder to another machine and open it directly without AppData dependencies or path rebinding.
- **Impact**: Health checks must surface absolute internal paths, external references, reserved-name conflicts, missing files, and orphan state rows without writing to the database. Repair actions are explicit and backup-protected. Backup restore must release the active SQLite connection before replacing `.lamber.sqlite`, especially on Windows.

### ADR-001: File-based JSON Database Repository
- **Decision**: Store projects and files in `projects_store.json` using atomic loading, modifying, and saving operations. (Superceded by SQLite in Phase 1 upgrade).
- **Reason**: Simplifies desktop deployments without DB servers, preparing interfaces for an easy future SQLite migration.
- **Impact**: Concurrency is managed via Rust load-modify-save cycles. (Migrated to SQLite transaction management).

### ADR-005: SQLite Core Database with Dual-Repository Swapping
- **Decision**: Introduce SQLite as the primary data store and implement a dynamic `DualProjectRepository` wrapper to allow hot-swapping between JSON and SQLite.
- **Reason**: To support relational features (ICT lifecycle fields, AI knowledge, scanned page indexes) and handle transactional integrity safely.
- **Impact**: JSON files are backed up automatically on startup, migrated using transactional inserts, and dynamically swapped to SQLite without requiring an application restart.

### ADR-002: Tonal Shifting (No solid borders) UI Theme
- **Decision**: Avoid standard `border-t`, `border-l` 1px lines in page segments, replacing them with background colors (`bg-muted` vs `bg-card`) and surface nesting.
- **Reason**: Implements the professional ledger feel defined in `DESIGN.md`.
- **Impact**: Keep border classes minimal, reserved only for inputs or critical errors.

### ADR-003: Decoupled AI Context Serialization
- **Decision**: Serialize active views into Markdown inside the frontend React application and send them as LLM system context.
- **Reason**: Keeps frontend state modifications synced to the AI assistant in real-time, even before changes are written to the database.
- **Impact**: Requires views to continually update `useAiContextStore` on form change events.

### ADR-004: Navigational Origin-Aware Routing
- **Decision**: Record `entrySource` in `useNavigationStore` when launching the calculator.
- **Reason**: Ensures clicking the back button in the calculator correctly returns the user to the Project Board or the main Hub, irrespective of scheme switches during the session.
- **Impact**: Prevents navigation loops or broken paths when changing options midway.

## 6. Current priorities

1. Multi-round conversation history optimization to prevent Ollama context window overflow.
2. Private file upload RAG implementation.
3. Install package size clean up.
4. i18n English localization.

## 7. Known fragile areas

- **Excel Cell Coordinates mapping**: Cell coordinates in `excel.rs` and `parse_benefit_excel` are hardcoded. Changing Excel template sheets will break parsing.
- **Zip / Word XML extraction**: docx variable extraction depends on direct XML parsing. Malformed documents or tables inside docx templates can cause extraction failures.
- **0-tolerance reconciliation check**: Floating-point rounding in JavaScript vs Rust Decimal might trigger false positives in validation. All UI inputs are handled via strings or decimal crates.
- **SQLite INSERT OR REPLACE Cascade Deletion Risk**: Avoid using `INSERT OR REPLACE` on parent tables (e.g. `projects`) that have child tables referencing them with `ON DELETE CASCADE` (like `project_directories`, `project_files`, `benefit_schemes`, `benefit_snapshots`). SQLite's `REPLACE` executes as a `DELETE` followed by an `INSERT`, which fires foreign key cascade deletions, wiping out all related child rows. Use manual existence checks and separate `UPDATE`/`INSERT` commands instead.

## 8. Do not break

- **No solid borders theme**: Do not add solid grey or black borders separating card contents. Use shifts in shades of blue/grey.
- **No direct AI database updates**: Never write backend file-writing commands triggered directly by AI agents without intermediate form confirmation.
- **No-physical file deletion in scans**: Directory scans must only toggle the `exists` flag on linked files, never delete files physically.
- **Health check read-only boundary**: `run_workspace_health_check` must never mutate database rows or filesystem structure; only explicit repair commands may write, and they must create a database backup first.

## 9. Open questions

1. *Should we migrate the JSON storage to SQLite in the next phase?*
2. *Do we need to support custom discount rates per year in the 10-year cashflow simulation instead of a single flat rate?*
