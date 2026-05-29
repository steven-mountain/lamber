# AI_CONTEXT.md

Last updated: 2026-05-28

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

Workspace selection is part of the Project Board workflow: entering Project Board first shows the Project Workspace layer when no workspace is open, and only loads project cards after a workspace is selected.

Workspace switching should present a workspace overview similar to the project list, backed by locally recorded `recentWorkspaces`. Users choose a recorded workspace card first; opening another folder is an explicit secondary action.

Current project editing state is split by domain. Project detail fields are stored in `projects`; ICT lifecycle state is stored in `project_lifecycle_states`; funding model and cashflow state are stored in `project_cashflow_states`; benefit方案 history is stored in `benefit_schemes` and `benefit_snapshots`; template form state is stored in `project_template_states` with legacy `project_settings` fallback; template assets are stored as files plus `project_template_assets` metadata. The frontend `useSaveStore` tracks dirty scopes and only clears a scope after its registered domain handler returns the scopes it actually saved for the same workspace/project context. Template form failures must propagate to the global save handler; otherwise `template-forms` remains dirty.

When opening an ICT project, the frontend must restore `project_lifecycle_states` current editor state before falling back to benefit scheme snapshots. This keeps global save / Ctrl+S edits, such as project background, independent from the older scheme snapshot history.

On application launch, the view always defaults to the HubView (`"hub"`) to guarantee the user starts at the central tool selection panel, rather than automatically restoring the last active view. The `WorkspaceGate` component and view headers now support a standardized `← 返回集市` back-navigation to the Hub (across all entry views including Project Board and Data Management, in both active and inactive workspace states).

## Read order for new AI sessions

Before starting any task, read the status and architecture documentation in this order:

1. [PROJECT_STATUS.md](./PROJECT_STATUS.md) - Current development status, milestones, business rules, and constraints.
2. [ARCHITECTURE_MAP.md](./ARCHITECTURE_MAP.md) - Directory layout, Boot processes, data flow, and state serialization.
3. [CHANGELOG_AI.md](./CHANGELOG_AI.md) - History of AI contributions, decisions, and ongoing questions.
4. Relevant source files only (do not scan the entire codebase).

## Core modules

- **Project Board (项目看板)**: Visual hub displaying project indicators, project phases, and local folders linkage.
- **ICT Lifecycle Calculator (ICT生命周期测算)**: Main calculation engine handling 10-year cashflows, NPV, margins, and selection fee mappings. Standardizes persistence of project backgrounds and economic assumptions directly in the scheme snapshots database storage.
- **AI Assistant / Copilot (智能顾问)**: Streamed (SSE) local LLM chatbot that pulls the active view's serialized state to give contextual recommendations based on a built-in capabilities database.
- **Template Word/Excel Engine (文档填报与填充)**: Backend Rust logic mapping form values and calculation results into docx variables and excel cells. Now utilizes separate template form data (saved in `project_settings` under key `template_form_data::<template_name>`) and sandboxed image assets (saved to the bound project folder under `{project_name}-图片/assets/` if linked, or falling back to `{app_data_dir}/projects/{project_id}/assets/` and tracked in `project_template_assets`), ensuring form configurations are kept lightweight and database volume is small.
- **Local File & Scan Engine (本地文件扫描管理)**: Service linking projects to local directory folders, scanning files (Word/Excel/PDF), managing linked vs. sandbox-copied storage modes, handling elastic path resilience via project roots and relative subpaths, and automatically importing calculations from Excel files starting with "效益分析表" and ending with ".xlsx"/".xls" (choosing the newest by modification date) when folder scans are triggered for projects with 0 schemes.

## Current priorities

1. **AI Conversation History Window Context Optimization**: Optimize messaging queues and history truncation to prevent context window overflow.
2. **Local RAG Integration**: Allow users to upload local PDF/Markdown files as part of the AI Copilot's private knowledge base.
3. **App Build Size Optimization**: Clean up redundant npm and cargo dependencies to reduce the installer footprint.
4. **Internationalization (i18n)**: Implement `i18next` for Chinese and English multi-language toggling.

## High-risk areas

- **0-tolerance reconciliation check**: Any changes to input fields or calculations must satisfy `excl_tax = incl_tax / (1 + rate)` within a zero-tolerance margin. Mismatches block navigation and document generation.
- **SQLite Connection & Transaction management**: The storage layer has migrated to SQLite. When writing multi-row operations or migrating data, always wrap operations in a transaction (`tx`) to prevent database locks and ensure structural integrity. Use `Arc<Mutex<rusqlite::Connection>>` to share connection locks. Avoid `INSERT OR REPLACE` when saving parent records (like `projects`) that have foreign key cascade-delete relations, as `REPLACE` acts as a `DELETE` followed by an `INSERT` in SQLite, wiping out related child rows in `project_directories`, `project_files`, `benefit_schemes`, and `benefit_snapshots`.
- **Workspace readiness**: Do not run database-backed project operations unless `WorkspaceRuntime::require_workspace()` and `require_db()` succeed. Do not reintroduce startup fallback to AppData `projects_store.db`.
- **Local Folder binding & scan synchronization**: Scanning folders updates physical file existence, but must never delete files physically on linked mode. Sandboxed (`copied`) files should be physically deleted only after user confirmation. Folder binding warning options allow auto-parent folder registration as project roots. Scanning and adding files inside absolute-only (rootless) folders must keep `directory_id` set to `None` to prevent SQLite `FOREIGN KEY constraint failed` errors on the `project_files` table.
- **Template Form & Image Assets Separation**: Form configurations are saved in `project_settings` under key `template_form_data::<template_name>`, and any large base64 image data is stripped beforehand to prevent database bloat. Pasted/dropped images are uploaded directly to the backend project asset sandbox, tracked in `project_template_assets`, and represented using `assetId` references. When generating documents, the frontend must NOT pass absolute file paths; the backend validates project ownership of the `assetId`, loads the physical file from the sandbox, and embeds it directly. Legacy base64 images must be migrated automatically during the next form save. Image uploads are constrained to PNG, JPEG, and WEBP formats and must not exceed 20MB.
- **Built-in Product Recommendations**: Recommended products must be cross-checked with codes in `midThreeConstants.ts`. If matched, append `[系统内置]` label; otherwise, append `【系统外扩展】`.

## Where to start for common tasks

- **Project Board changes**:
  - Frontend: [ProjectBoard.tsx](../src-ui/src/views/ProjectBoard.tsx)
  - Backend: [project_files/service.rs](../src-tauri/src/project_files/service.rs)
- **Benefit Calculation math / NPV changes**:
  - Math Engine: [calculator.rs](../src-tauri/src/benefit/calculator.rs)
  - Lifecycle View: [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx)
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
