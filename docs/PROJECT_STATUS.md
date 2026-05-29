# PROJECT_STATUS.md

Last updated: 2026-05-28 (Workspace Save Boundaries)

## 0. Latest persistence update

Workspace refactoring phase 3 separates current project editing state from scheme snapshots. Project detail updates remain in `projects`; ICT lifecycle editor state is saved in `project_lifecycle_states`; funding models, assumptions, cashflow tables, sector cashflows, and metrics are saved in `project_cashflow_states`; benefit方案 history remains in `benefit_schemes` / `benefit_snapshots`; template field values, mappings, and output configuration are saved in `project_template_states` with compatibility fallback to legacy `project_settings` keys. Template binary/image assets continue to use `project_template_assets` and file-backed storage.

The frontend now has `useSaveStore`, a domain save service, a global save button, Ctrl/Command+S handling, and unsaved-change guards. Dirty scopes are cleared only after their registered domain save handler returns the scopes it actually persisted for the same workspace and project context. Template form saving must propagate failures to `useSaveStore`; autosave may be silent, but global save must keep `template-forms` dirty if `saveTemplateState` fails.

Opening an ICT project now restores current lifecycle editor state from `project_lifecycle_states` before reading benefit scheme snapshots, even when navigation includes a default scheme id. Scheme snapshots remain historical/fallback data, not the primary source for current project background or lifecycle edits.

## 1. Project summary

Lamber is a lightweight sales support and project management desktop tool designed for client managers and solution experts in the 5G/ICT domains. It addresses the inefficiencies of manual financial calculations, document generation, and project folder tracking by consolidating ICT project lifecycle assessment, benefit calculator, contract template filling (Word), economic evaluation sheets (Excel), and local files scanning into a single Tauri-based application.

## 2. Tech stack

- **Frontend**: Vite + React 18 + TypeScript + Zustand
- **Desktop runtime**: Tauri v2 + Rust
- **State management**:
  - Global navigation: `useNavigationStore` (Zustand + local storage, always defaulting to "hub" view on application startup)
  - AI Context: `useAiContextStore` (Zustand + local storage + Tauri events)
  - Layout & view modes: Local storage (`lamber_project_board_view_mode`, etc.)
- **Database/Persistence**: SQLite (via `rusqlite` with `bundled` feature) for structured relational tables, alongside dynamic `projects_store.json` backward compatibility
- **Workspace Runtime**: Business data is now scoped to an explicit Lamber Workspace root containing hidden system files `.lamber.workspace.json`, `.lamber.sqlite`, `.backups/`, and `.exports/`. Project folders (e.g. `项目A`, `项目B`) are placed directly inside the workspace root without an intermediate `projects/` layer. The app remembers `recentWorkspaces` and `lastOpenedWorkspacePath` in local AppConfig. It supports initializing an existing general project root directory as a Lamber Workspace and bulk importing eligible first-level subdirectories as workspace internal projects (with automatic creation of `project.json` and assets/documents/analyses folders). Accessing workspace-backed features without a workspace redirects to `WorkspaceGate` with matching headers, which now provides a standardized '← 返回集市' back button to return to the Hub (across all entry views including Project Board and Data Management, both when workspace is active or inactive).
- **Styling**: Tailwind CSS + Shadcn/UI (Radix UI) + HSL-based design system
- **AI integration**: Local SSE streaming client (Ollama / OpenAI standard endpoint) with semantic Markdown context serialization
- **File handling**:
  - Word variable replacement: Backend Rust `docx-template`
  - Excel template filling/parsing: Backend Rust `calamine` + `umya-spreadsheet` + `rust_xlsxwriter`
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

### 3.3 Investment Benefit Analysis (项目效益分析)

**Status**: Active (Fully Implemented)

- **Current behavior**: Standalone benefit tool `BenefitTool.tsx` allowing single-item custom parameter calculations and batch Excel files processing using `process_excel_batch`.
- **Known requirements**: Excel batch processing outputs files into the workspace `output/` directory.
- **Known issues**: Batch processor requires Excel headers to match exact templates.
- **Related files**:
  - [BenefitTool.tsx](../src-ui/src/views/BenefitTool.tsx)
  - [excel.rs](../src-tauri/src/benefit/excel.rs)

### 3.4 AI Assistant (Lamber 智能顾问)

**Status**: Active (Fully Implemented)

- **Current behavior**: SSE stream chatbot utilizing a local Ollama server (or OpenAI compatible API). Automatically gathers frontend serialized business states and injects them as system prompts, together with a product capabilities catalog (`midThreeConstants.ts`) to make suggestions.
- **Known requirements**: Multi-round history optimization and PDF/Markdown local RAG.
- **Known issues**: High token consumption on large templates since the entire form structure is serialized.
- **Related files**:
  - [AiChatPanel.tsx](../src-ui/src/components/ai/AiChatPanel.tsx)
  - [useAiContextStore.ts](../src-ui/src/store/useAiContextStore.ts)
  - [AiRuntime.ts](../src-ui/src/ai/AiRuntime.ts)

### 3.5 File / Excel / Word Integration (文件及模板集成)

**Status**: Active (Fully Implemented)

- **Current behavior**: Scans directories for docx/xlsx files. Back-fills form fields (from `TemplateForms.tsx`) into files via Rust backend. Word templates are generated via `docx-template`. Excels can be filled or parsed back (via `parse_benefit_excel`, which resolves workspace-relative paths against the active workspace root) to overwrite project states. Supports sandbox "copied" files or raw disk "linked" files. Template forms (`TemplateForms.tsx`) and embedded image resources are persisted separately: lightweight forms/table settings are saved in `project_settings` (under key `template_form_data::<template_name>`), while pasted/dropped images are uploaded instantly to project assets sandboxes (bound project folder under `{project_name}-图片/assets/` if linked, otherwise falling back to `{app_data_dir}/projects/{project_id}/assets/`) and tracked in the `project_template_assets` metadata table. Document generation reads binary contents directly from sandbox files via backend validation. Automatically imports Excel calculations: when a folder scan is triggered (manually, via folder binding, or during workspace project bulk initialization), if the project has 0 schemes, it filters for Excel files whose names start with "效益分析表" and end with ".xlsx"/".xls", chooses the newest one by modification date, parses its economic parameters, and saves it as the default scheme "Excel导入测算方案".
- **Known requirements**: Scan timestamps are updated without modifying physical files.
- **Known issues**: Cell coordinates mapping in Excel template is fragile if the spreadsheet structure changes.
- **Related files**:
  - [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx)
  - [docfill.rs](../src-tauri/src/docfill.rs)
  - [project_files/service.rs](../src-tauri/src/project_files/service.rs)

### 3.6 Data Management Center & Path Resilience (数据管理中心与路径韧性)

**Status**: Active (Fully Implemented in Phase 2)

- **Current behavior**: Avoids hardcoded absolute paths by employing global project roots (`project_roots`), relative paths (`project_directories`), and file fingerprints (`size:modified_at:first_8kb_hash`). Supports global project roots CRUD, folder binding warning triggers (offering auto-parent folder registration as root with relative project subpaths, or absolute-only paths), a Health Checks Dashboard showing file links health metrics and self-healing reports, bulk path relocation (relocating directories across volumes with previews), and candidate scanner to import folders as new projects (auto-identifying file roles and resolving naming conflicts via merge/new/skip options). Allows scanning and syncing of rootless project directories safely without triggering foreign key constraint failures.
- **Related files**:
  - [DataManagement.tsx](../src-ui/src/views/DataManagement.tsx)
  - [ProjectFilesTab.tsx](../src-ui/src/components/project/ProjectFilesTab.tsx)
  - [ProjectBoard.tsx](../src-ui/src/views/ProjectBoard.tsx)
  - [project_files/roots.rs](../src-tauri/src/project_files/roots.rs)
  - [project_files/health.rs](../src-tauri/src/project_files/health.rs)
  - [project_files/relocation.rs](../src-tauri/src/project_files/relocation.rs)
  - [project_files/import_scanner.rs](../src-tauri/src/project_files/import_scanner.rs)

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

## 9. Open questions

1. *Should we migrate the JSON storage to SQLite in the next phase?*
2. *Do we need to support custom discount rates per year in the 10-year cashflow simulation instead of a single flat rate?*
