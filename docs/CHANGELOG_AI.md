# CHANGELOG_AI.md

This changelog records structural modifications, business rules, and context changes made by AI agents to maintain a reliable project state mapping.

## 2026-05-28

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
- Standardized project folders creation inside `workspaceRoot/projects/{safeProjectName}/` with `assets/`, `documents/`, `analyses/` subfolders and a backup `project.json` manifest.
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
