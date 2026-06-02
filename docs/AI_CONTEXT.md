# AI_CONTEXT.md

Last updated: 2026-06-01 (ICT Billing Subject Name Extension)

## What this project is

Lamber is a lightweight sales support desktop tool built with **Tauri, React, and Rust**. It helps client managers and solutions experts manage ICT project lifecycles, perform economic benefit calculations (NPV, margin, cashflow), and fill out standardized bidding and project review documents (Word/Excel) using structured templates.

## Workspace foundation

Business persistence is scoped to an explicit Lamber Workspace root:

```text
LamberWorkspace/
├─ .lamber.workspace.json
├─ .lamber.sqlite
├─ .backups/
├─ .exports/
├─ .projects/ (用于模板沙箱图片资源等)
├─ 项目A/
├─ 项目B/
```

`WorkspaceRuntime` owns the current workspace state and active SQLite connection. Project, file, template asset, root, health, relocation, import, and document generation commands must require an open workspace and use `{workspaceRoot}/.lamber.sqlite`. Startup may restore `lastOpenedWorkspacePath`; if it fails, the app must show WorkspaceGate and must not create an empty DB or fall back to AppData.

When a workspace is opened or created, the workspace root is automatically registered in `project_roots` if it is missing. All new projects are created directly inside the workspace at `{safeProjectName}/` with standard `assets/`, `documents/`, and `analyses/` subfolders. Workspaces can also be initialized from existing directories, bulk-importing their first-level subdirectories as projects at the root of the workspace without moving them, writing or complementing `project.json` manifests, and using workspace-relative paths (which are dynamically resolved against the active workspace root by backend commands like `parse_benefit_excel`) for portable durability. Project folders created, imported, or bound under the workspace should therefore not ask the user to register an additional root directory.

When initializing a Workspace from an existing folder, commit project inserts, release the SQLite mutex, and return the opened Workspace before scanning imported project folders or attempting automatic Excel import. Scan/import runs as a background best-effort task and uses the initialized workspace root plus database handle captured at initialization time.

Flat workspace roots reserve `.lamber.workspace.json`, `.lamber.sqlite`, `.backups`, `.exports`, `.projects`, `backups`, `exports`, and `projects`; project creation and health checks must prevent project directories from using those names.

Dot-prefixed system entries are not enough on Windows. The backend must also mark `.lamber.workspace.json`, `.lamber.sqlite`, `.backups`, `.exports`, and `.projects` with the Windows Hidden attribute whenever those entries are created, repaired, imported, or opened.

Workspace portability maintenance is implemented in `workspace_maintenance.rs`. It provides daily/manual SQLite backups, safe restore from backup, backup file deletion, read-only workspace health checks, explicit repair commands, external path listing, dry-run internal absolute path conversion, fixed-root `.lamber.zip` export, `.lamber.zip` validation/import, and file-manager reveal. Export archives preserve `workspaceId`, use the workspace root as the zip root, include `.lamber.workspace.json`, a consistent `.lamber.sqlite`, `.projects/`, project folders, and `export-manifest.json`, and exclude `.backups` / `.exports` by default. Import preserves `workspaceId`, can read both direct-root archives and archives with a single top-level wrapper folder, treats the selected folder as the parent destination, and creates `{selectedFolder}/{workspaceName}` by default. The import command uses flat IPC arguments (`openAfterImport`, `conflictStrategy`, `destinationName`) rather than a nested `options` object. Frontend callers should call the maintenance service with explicit arguments, and the backend normalizes `openAfterImport` from JSON to avoid Tauri rejecting stale object-shaped payloads before import validation runs; absent or malformed nested booleans default to `false`.

After Workspace export, the UI should open the generated `.lamber.zip` file's containing directory, not a generic system location. The backend reveal command still supports file selection, using a Windows Explorer `/select,PATH` argument without embedded quotes for better path parsing.

`run_workspace_health_check` must stay read-only. Any repair, including creating missing folders, regenerating `project.json`, or converting workspace-internal absolute paths to relative paths, must be invoked separately and must create a database backup first. External roots and external template assets are reported but not moved or rewritten automatically.

When regenerating a missing project `project.json`, repair code should resolve the project directory with the same fallback as health checking: `relative_path`, `linked_folder_relative_path`, `folder_path`, then `folder_name`. This preserves repairability for older or imported flat workspace records before full path normalization.

External `module_path:*` references can be repaired by resetting the module base directory to `.projects/modules/{moduleId}` inside the active Workspace. The repair updates app-level module config and creates `templates/` / `output/`, but it must not move or delete files from the old external directory. Retired standalone module paths for `benefit_tool` and `docfill_tool` are ignored by workspace external-path reporting.

Workspace selection is part of the Project Board workflow: entering Project Board first shows the Project Workspace layer when no workspace is open, and only loads project cards after a workspace is selected.

Workspace switching should present a workspace overview similar to the project list, backed by locally recorded `recentWorkspaces`. Users choose a recorded workspace card first; opening another folder is an explicit secondary action.

Data Management also exposes the local `recentWorkspaces` association list in a dedicated Workspace Management tab, using a card-based Workspace registry. Card clicks are local selection only: they highlight the record without opening the Workspace or updating `lastOpenedAt`, so the list order must not change. Inline open/reveal/unlink buttons must stop propagation so each action stays independent. The open button should switch to the target Workspace and navigate directly to the Project Board after the Workspace is active; opening the already active Workspace should also enter the Project Board. Users can open, reveal, or unlink remembered Workspaces there. Unlinking means removing the local association from `config.json`; it must never delete the physical Workspace folder, `.lamber.sqlite`, or project files. If the unlinked Workspace is currently open, the UI must run the dirty guard first, then the backend clears the active runtime/database connection and `lastOpenedWorkspacePath`. Active Workspace backup, export, restore, and health actions belong in the separate Workspace Maintenance tab.

Current project editing state is split by domain. Project detail fields are stored in `projects`; ICT lifecycle state is stored in `project_lifecycle_states`; funding model and cashflow state are stored in `project_cashflow_states`; benefit方案 history is stored in `benefit_schemes` and `benefit_snapshots`; template form state is stored in `project_template_states` with legacy `project_settings` fallback; template assets are stored as files plus `project_template_assets` metadata. The frontend `useSaveStore` tracks dirty scopes and only clears a scope after its registered domain handler returns the scopes it actually saved for the same workspace/project context. Template form failures must propagate to the global save handler; otherwise `template-forms` remains dirty.

When opening an ICT project, the frontend must restore `project_lifecycle_states` current editor state before falling back to benefit scheme snapshots, then overlay the latest `project_cashflow_states.assumptions_json` for cashflow-domain price inputs. This keeps global save / Ctrl+S edits, such as project background and IT/CT revenue/cost prices, independent from older scheme snapshot history.

On application launch, the view always defaults to the HubView (`"hub"`) to guarantee the user starts at the central tool selection panel, rather than automatically restoring the last active view. The Hub currently exposes Project Board, ICT Lifecycle Calculator, and Data Management. The old standalone Investment Benefit Analysis and Document Material Production modules have been retired; their shared calculation, Excel import, and template generation engines remain available through ICT/project workflows. The `WorkspaceGate` component and view headers support standardized back-navigation to the Hub across Project Board and Data Management states.

## Read order for new AI sessions

Before starting any task, read the status and architecture documentation in this order:

1. [PROJECT_STATUS.md](./PROJECT_STATUS.md) - Current development status, milestones, business rules, and constraints.
2. [ARCHITECTURE_MAP.md](./ARCHITECTURE_MAP.md) - Directory layout, Boot processes, data flow, and state serialization.
3. [CHANGELOG_AI.md](./CHANGELOG_AI.md) - History of AI contributions, decisions, and ongoing questions.
4. Relevant source files only (do not scan the entire codebase).

## Core modules

- **Project Board (项目看板)**: Visual hub displaying project indicators, project phases, and local folders linkage.
- **ICT Lifecycle Calculator (ICT生命周期测算)**: Main calculation engine handling 10-year cashflows, NPV, margins, and selection fee mappings. Standardizes persistence of project backgrounds and economic assumptions directly in the scheme snapshots database storage.
- **ICT Standard Subject Presentation Names**: Every fixed revenue/cost billing subject can carry an optional product/business name and an optional billing subject name. The fixed subject catalog remains the source of truth for `subjectCode`, standard subject label, Excel row mapping, and document prefix. Product/business names (`customSubjectName` / `custom_subject_name`) and billing subject names (`billingSubjectName` / `billing_subject_name`) are stored separately and must not replace the standard subject identity used for calculations.
- **AI Assistant / Copilot (智能顾问)**: Streamed (SSE) local LLM chatbot that pulls the active view's serialized state to give contextual recommendations based on a built-in capabilities database.
- **Template Word/Excel Engine (文档填报与填充)**: Backend Rust logic mapping form values and calculation results into docx variables and excel cells. Now utilizes separate template form data (saved in `project_settings` under key `template_form_data::<template_name>`) and sandboxed image assets (saved to the bound project folder under `{project_name}-图片/assets/` if linked, or falling back to `.projects/{project_id}/assets/` inside the active workspace and tracked in `project_template_assets`), ensuring form configurations are kept lightweight and database volume is small.
- **Lifecycle Excel Subject Name and Q Column Fill**: `3-直接经济效益评估表` output writes all mapped standard subject rows through a unified mapping. Name cells receive the shared resolver display name using `计费科目名称 > 具体业务/产品名称 > 标准科目名称`, `G` receives tax-exclusive amount, and `Q` receives tax-inclusive amount for every mapped subject. Empty/zero frontend amounts must remain blank in generated Excel amount cells rather than being written as `0`. Do not insert/delete Excel rows or alter formula regions outside this controlled input mapping.
- **Shared Document Subject Presentation**: Project sign-off and meeting-review wording must use the shared subject resolver from `ictSubjectCatalog.ts`. Document business names use `计费科目名称 > 具体业务/产品名称 > existing fallback` before adding the fixed category prefix, and aggregation/deduplication uses the final resolved document name. Do not derive CT investment subject wording from mid-platform capability names; keep the underlying tax-exclusive investment calculations unchanged.
- **Sign-off and Meeting Investment Wording**: The ICT project sign-off template uses full-line variables `PROJECT_INVESTMENT_SITUATION` and `PROJECT_REVENUE_SITUATION`. `TemplateForms.tsx` generates these by enumerating non-zero measurement-table subjects with `计费科目名称` first, standard subject fallback, and fixed category prefixes across IT, CT, non-IT/CT, and comprehensive categories. Meeting-review `PROJECT_TOTAL_INVESTMENT_DETAIL` / meeting `PROJECT_TOTAL_INVESTMENT` reuse the same generated investment wording. The optional "申请立项后甄选" phrase is controlled by the sign-off checkbox and must not be hardcoded in the template.
- **Template Image Embedding Boundary**: Frontend preview URLs (`asset://localhost/...`) are UI-only. Word generation image payloads should carry `assetId` values, and the backend must resolve and embed the image bytes through Workspace SQLite asset validation rather than writing paths or JSON text into the document.
- **Inquiry Vendor Screenshot State**: Inquiry vendor quote screenshots live on each vendor row's `images` array. Regenerating three-vendor quote amounts must preserve those images by vendor name or row position, and async upload callbacks must update the latest vendor state.
- **Lifecycle Document Output Ownership**: ICT lifecycle document generation must resolve output directories through the active Workspace. Relative project folder paths are Workspace-relative, not process-relative, so generated files should land in the bound project directory under the Workspace and never under `src-tauri`.
- **Local File & Scan Engine (本地文件扫描管理)**: Service linking projects to local directory folders, scanning files (Word/Excel/PDF), managing linked vs. sandbox-copied storage modes, handling elastic path resilience via project roots and relative subpaths, and automatically importing calculations from Excel files starting with "效益分析表" and ending with ".xlsx"/".xls" (choosing the newest by modification date) when folder scans are triggered for projects with 0 schemes.

Additional current AI context capability:

- **AI Project Context Service**: Read-only Rust backend module `src-tauri/src/ai_context` exposing `build_ai_project_context`. It reads only persisted official project state from the active Workspace `.lamber.sqlite` through `WorkspaceRuntime`, validates `projectId`, and returns structured summaries for overview, lifecycle, cashflow, benefit, templates, and files. It does not read full documents, image binaries/base64, unsaved frontend drafts, or perform RAG/embedding work.
- **AI Chat Context Composer**: Frontend module `src-ui/src/ai/context` used by `AiChatPanel` on each message send. It loads saved official project context from Workspace SQLite when an active project is available, and separately injects the current page's unsaved draft overlay only when `useSaveStore` reports matching dirty scopes for the same project/page. The overlay is sanitized to remove base64/data URL previews and absolute paths and to truncate large content.
- The chat composer refreshes the latest active project identity from existing navigation/current-project storage, including the ICT lifecycle active project key, every time a message is sent. This supports a floating AI window that was opened before project selection or kept open while switching projects. The refreshed identity only determines which Workspace SQLite project context to request; saved official state still comes from `build_ai_project_context`, and frontend dirty data remains an unsaved draft overlay.
- Workspace changes must clear stale active project/scheme IDs before AI context lookup. Project Board publishes a compact current-workspace board summary for workspace-level AI questions such as project counts; selected project detail remains validated through Workspace SQLite when available.
- `AiChatPanel` resets the streaming parser before adding a new assistant placeholder. The active streaming bubble is overwritten from the current parser output, keeping previous completed responses isolated from the new thinking/streaming response.
- **AI Template Detail Context**: The AI chat composer can request one active template's saved detail through `build_ai_project_context` using `requestedSources: ["template_detail"]` and `activeTemplateId`. The backend reads only the specified template from Workspace SQLite (`project_template_states`, with legacy `project_settings` fallback), sanitizes base64/data URL previews, local absolute paths, and oversized nested payloads, and returns image asset metadata only.
- **Controlled Template Vision Assets**: Template image thumbnails expose an explicit AI analysis action. The frontend sends only `projectId + templateId + assetId` metadata to the AI window. On message send, `load_ai_template_asset` validates the asset belongs to the current project in the active Workspace, checks PNG/JPEG/WEBP and size constraints, resolves a workspace-contained file, and returns a temporary data URL for the existing multimodal `image_url` request. Unselected images are not read.
- **AI Workspace Specified Project Routing**: The AI chat composer can request `list_ai_workspace_projects` to build a lightweight current-Workspace project index before each non-Hub message. Explicitly named projects are matched deterministically against that index and then loaded through `build_ai_project_context` using real `projectId` values. Explicit project mentions override the currently open project, and at most two specified projects are deep-loaded per turn for comparison. Workspace-level list questions can use the lightweight index instead of full project/template reads.
- **Specified Project Template Detail Routing**: When a user names a project and a uniquely identifiable template, the composer first loads that project's template summary, resolves the template name or known alias, and then reuses the existing `template_detail` source for that one template only. Ambiguous template names produce warnings and do not trigger full template scans.

## Current priorities

1. **AI Conversation History Window Context Optimization**: Optimize messaging queues and history truncation to prevent context window overflow.
2. **Local RAG Integration**: Allow users to upload local PDF/Markdown files as part of the AI Copilot's private knowledge base.
3. **App Build Size Optimization**: Clean up redundant npm and cargo dependencies to reduce the installer footprint.
4. **Internationalization (i18n)**: Implement `i18next` for Chinese and English multi-language toggling.

## High-risk areas

- **0-tolerance reconciliation check**: Any changes to input fields or calculations must satisfy `excl_tax = incl_tax / (1 + rate)` within a zero-tolerance margin. Mismatches block navigation and document generation.
- **Subject presentation names are non-financial metadata**: Editing a product/business name or billing subject name must not alter tax rates, inclusive/exclusive amounts, cashflow, NPV, selection fee, or reverse-calculation behavior. Old project data without custom names must display/export the standard subject names. CT paired subjects (`产品收入` ↔ `其他产品成本`, `专线收入` ↔ `专线带宽成本`) keep product/business names and billing subject names synchronized like their amount pass-through path.
- **SQLite Connection & Transaction management**: The storage layer has migrated to SQLite. When writing multi-row operations or migrating data, always wrap operations in a transaction (`tx`) to prevent database locks and ensure structural integrity. Use `Arc<Mutex<rusqlite::Connection>>` to share connection locks. Avoid `INSERT OR REPLACE` when saving parent records (like `projects`) that have foreign key cascade-delete relations, as `REPLACE` acts as a `DELETE` followed by an `INSERT` in SQLite, wiping out related child rows in `project_directories`, `project_files`, `benefit_schemes`, and `benefit_snapshots`.
- **Workspace readiness**: Do not run database-backed project operations unless `WorkspaceRuntime::require_workspace()` and `require_db()` succeed. Do not reintroduce startup fallback to AppData `projects_store.db`.
- **Workspace backup and restore safety**: Restoring `.lamber.sqlite` from backup must release or close the active SQLite connection before replacing the file, then reopen the workspace. On Windows this is required to avoid file-lock replacement failures.
- **Workspace archive safety**: Importing `.lamber.zip` must validate structure and reject zip entries with absolute paths, `..` components, or extraction targets outside the chosen destination directory.
- **Local Folder binding & scan synchronization**: Scanning folders updates physical file existence, but must never delete files physically on linked mode. Sandboxed (`copied`) files should be physically deleted only after user confirmation. Folder binding warning options allow auto-parent folder registration as project roots. Scanning and adding files inside absolute-only (rootless) folders must keep `directory_id` set to `None` to prevent SQLite `FOREIGN KEY constraint failed` errors on the `project_files` table.
- **Template Form & Image Assets Separation**: Form configurations are saved in `project_settings` under key `template_form_data::<template_name>`, and any large base64 image data is stripped beforehand to prevent database bloat. Pasted/dropped images are uploaded directly to the backend project asset sandbox, tracked in `project_template_assets`, and represented using `assetId` references. When generating documents, the frontend must NOT pass absolute file paths; the backend validates project ownership of the `assetId`, loads the physical file from the sandbox, and embeds it directly. Legacy base64 images must be migrated automatically during the next form save. Image uploads are constrained to PNG, JPEG, and WEBP formats and must not exceed 20MB.
- **Built-in Product Recommendations**: Recommended products must be cross-checked with codes in `midThreeConstants.ts`. If matched, append `[系统内置]` label; otherwise, append `【系统外扩展】`.

Additional read-only AI boundary:

- **AI Project Context Read Boundary**: `build_ai_project_context` only performs SQLite `SELECT` reads through the active `WorkspaceRuntime` database. It does not accept workspace/database paths from the frontend, does not mutate business data, does not trigger scans/repairs/saves, and does not expose template asset absolute paths or image/document binary content.
- **AI Chat Context Boundary**: AI chat prompt construction separates saved official state from current unsaved draft overlay. Saved state comes from Workspace SQLite; draft state comes from frontend dirty page state and must be described as unsaved. Chat context loading failures are represented as warnings and must not be interpreted as empty project data. The AI still cannot write project data, auto-save, apply patches, run RAG/embedding work, generate file summaries, read image binaries, or change financial formulas.
- **AI Template Detail and Image Boundary**: Saved template detail is official only when loaded from Workspace SQLite. Template dirty state remains an unsaved draft overlay. Template image binaries are provided to the model only after explicit user selection and backend `projectId + assetId` ownership validation; the frontend does not pass or receive physical paths for AI analysis, and image base64 is not stored back into the database.
- **AI Workspace Routing Boundary**: Project names are only used to locate a target project inside the current Workspace for the current chat turn. Names are never used as business keys after matching; official data reads use the resolved `projectId`. Duplicate, ambiguous, missing, or overly broad project requests must not fall back to another project. Current frontend dirty state is injected only for the project it belongs to and must not contaminate another specified project.

## Where to start for common tasks

- **Project Board changes**:
  - Frontend: [ProjectBoard.tsx](../src-ui/src/views/ProjectBoard.tsx)
  - Backend: [project_files/service.rs](../src-tauri/src/project_files/service.rs)
- **Benefit Calculation math / NPV changes**:
  - Math Engine: [calculator.rs](../src-tauri/src/benefit/calculator.rs)
  - Lifecycle View: [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx)
- **ICT subject naming / Excel row mapping changes**:
  - Subject Catalog: [ictSubjectCatalog.ts](../src-ui/src/lib/ictSubjectCatalog.ts)
  - Lifecycle View: [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx)
  - Excel Fill: [docfill.rs](../src-tauri/src/docfill.rs)
- **AI Chat Copilot changes**:
  - Chat Box Component: [AiChatPanel.tsx](../src-ui/src/components/ai/AiChatPanel.tsx)
  - AI State Store: [useAiContextStore.ts](../src-ui/src/store/useAiContextStore.ts)
  - LLM Runtime: [AiRuntime.ts](../src-ui/src/ai/AiRuntime.ts)
- **Word/Excel Template Filling changes**:
  - Template forms: [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx)
  - Docfill service: [docfill.rs](../src-tauri/src/docfill.rs)
  - Excel parser: [excel.rs](../src-tauri/src/benefit/excel.rs)
- **UI Design modifications**:
  - Main styling: [index.css](../src-ui/src/index.css)
  - Design Tokens: [DESIGN.md](../DESIGN.md)
- **Data Management Center changes**:
  - Frontend: [DataManagement.tsx](../src-ui/src/views/DataManagement.tsx)
  - Backend: [roots.rs](../src-tauri/src/project_files/roots.rs), [health.rs](../src-tauri/src/project_files/health.rs), [relocation.rs](../src-tauri/src/project_files/relocation.rs), [import_scanner.rs](../src-tauri/src/project_files/import_scanner.rs)

## Do not break

- Do not move the Project Board back under the ICT Lifecycle view. They must remain independent.
- Do not let the AI rewrite `projects_store.json` directly without an explicit user transaction.
- Do not replace the tonal-shift design system with high-saturation solid borders or solid purple/blue panels.
