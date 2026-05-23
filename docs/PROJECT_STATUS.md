# PROJECT_STATUS.md

Last updated: 2026-05-23

## 1. Project summary

Lamber is a lightweight sales support and project management desktop tool designed for client managers and solution experts in the 5G/ICT domains. It addresses the inefficiencies of manual financial calculations, document generation, and project folder tracking by consolidating ICT project lifecycle assessment, benefit calculator, contract template filling (Word), economic evaluation sheets (Excel), and local files scanning into a single Tauri-based application.

## 2. Tech stack

- **Frontend**: Vite + React 18 + TypeScript + Zustand
- **Desktop runtime**: Tauri v2 + Rust
- **State management**:
  - Global navigation: `useNavigationStore` (Zustand + local storage)
  - AI Context: `useAiContextStore` (Zustand + local storage + Tauri events)
  - Layout & view modes: Local storage (`lamber_project_board_view_mode`, etc.)
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

- **Current behavior**: Displays projects in a vertical waterfall row layout (list view) or an adaptive CSS Grid responsive layout (card/grid view). Allows filtering by phase pills (全部, 立项中, 审批中, 实施中, 已完成), searching by keyword, adding notes via auto-saving textareas, binding local folder paths, and viewing key financial metrics (margin, NPV, NPVR, IRR, payback period, risk level).
- **Known requirements**: Maintain independent rendering from ICT Lifecycle view and persist UI selections in local storage.
- **Known issues**: Large note sizes might slightly lag during immediate auto-save.
- **Related files**:
  - [ProjectBoard.tsx](file:///d:/HermesJang/CMCC/tools/lambert/lamber/src-ui/src/views/ProjectBoard.tsx)
  - [projectService.ts](file:///d:/HermesJang/CMCC/tools/lambert/lamber/src-ui/src/utils/projectService.ts)

### 3.2 ICT Lifecycle Calculator (ICT生命周期测算)

**Status**: Active (Fully Implemented)

- **Current behavior**: Computes 10-year cashflows, NPV, NPV Rate, Margin Rate, IRR, and payback period. Supports bound-project mode and standalone "free" calculator mode. Includes quick split calculators (requiring 1% own product revenue allocation), selection fee calculators, and smart back-calculations. It enforces a 0-tolerance tax validation check before allowing users to see cashflows or generate documents.
- **Known requirements**: Maintain risk analysis criteria defined in backend Rust code.
- **Known issues**: Binary search limit for back-calculation is capped at 10 billion CNY.
- **Related files**:
  - [IctLifecycle.tsx](file:///d:/HermesJang/CMCC/tools/lambert/lamber/src-ui/src/views/IctLifecycle.tsx)
  - [calculator.rs](file:///d:/HermesJang/CMCC/tools/lambert/lamber/src-tauri/src/benefit/calculator.rs)

### 3.3 Investment Benefit Analysis (项目效益分析)

**Status**: Active (Fully Implemented)

- **Current behavior**: Standalone benefit tool `BenefitTool.tsx` allowing single-item custom parameter calculations and batch Excel files processing using `process_excel_batch`.
- **Known requirements**: Excel batch processing outputs files into the workspace `output/` directory.
- **Known issues**: Batch processor requires Excel headers to match exact templates.
- **Related files**:
  - [BenefitTool.tsx](file:///d:/HermesJang/CMCC/tools/lambert/lamber/src-ui/src/views/BenefitTool.tsx)
  - [excel.rs](file:///d:/HermesJang/CMCC/tools/lambert/lamber/src-tauri/src/benefit/excel.rs)

### 3.4 AI Assistant (Lamber 智能顾问)

**Status**: Active (Fully Implemented)

- **Current behavior**: SSE stream chatbot utilizing a local Ollama server (or OpenAI compatible API). Automatically gathers frontend serialized business states and injects them as system prompts, together with a product capabilities catalog (`midThreeConstants.ts`) to make suggestions.
- **Known requirements**: Multi-round history optimization and PDF/Markdown local RAG.
- **Known issues**: High token consumption on large templates since the entire form structure is serialized.
- **Related files**:
  - [AiChatPanel.tsx](file:///d:/HermesJang/CMCC/tools/lambert/lamber/src-ui/src/components/ai/AiChatPanel.tsx)
  - [useAiContextStore.ts](file:///d:/HermesJang/CMCC/tools/lambert/lamber/src-ui/src/store/useAiContextStore.ts)
  - [AiRuntime.ts](file:///d:/HermesJang/CMCC/tools/lambert/lamber/src-ui/src/ai/AiRuntime.ts)

### 3.5 File / Excel / Word Integration (文件及模板集成)

**Status**: Active (Fully Implemented)

- **Current behavior**: Scans directories for docx/xlsx files. Back-fills form fields (from `TemplateForms.tsx`) into files via Rust backend. Word templates are generated via `docx-template`. Excels can be filled or parsed back (via `parse_benefit_excel`) to overwrite project states. Supports sandbox "copied" files or raw disk "linked" files.
- **Known requirements**: Scan timestamps are updated without modifying physical files.
- **Known issues**: Cell coordinates mapping in Excel template is fragile if the spreadsheet structure changes.
- **Related files**:
  - [TemplateForms.tsx](file:///d:/HermesJang/CMCC/tools/lambert/lamber/src-ui/src/views/TemplateForms.tsx)
  - [docfill.rs](file:///d:/HermesJang/CMCC/tools/lambert/lamber/src-tauri/src/docfill.rs)
  - [project_files/service.rs](file:///d:/HermesJang/CMCC/tools/lambert/lamber/src-tauri/src/project_files/service.rs)

## 4. Important business rules

- **Project Board Independence**: The Project Board is independent of the ICT Lifecycle. Navigating to the calculator tracks the origin via `entrySource` so the back button knows where to return.
- **1% Own-Product Rule**: Projects must integrate at least 1% of total revenue as own-product CT revenue. The quick split calculator helps compute and fill this automatically.
- **0-Tolerance Tax Reconciliation**: In the calculator, tax-inclusive amount must match tax-exclusive + tax amount (`incl = excl * (1 + rate)`). If there is a rounding mismatch, navigation is blocked. If the error is <= 0.10, the user can override it.
- **Backend Risk Assessment**: Risk levels (高风险, 中风险, 低风险) are strictly evaluated by the Rust backend service `calculate_risk_level` in `service.rs` based on thresholds (margin: 8%; NPV rate: 4%).
- **Selection Fee Brackets**: Spanning quote ranges determine selection fees (e.g. quote <= 12,100 -> fee is 100; quote <= 48,500 -> fee is quote * 0.00825). Highest limit can also be reverse-calculated.
- **AI Direct Write Lock**: AI is allowed to read and analyze any active workspace variables, but is prohibited from updating the database files directly. Binders or "apply" buttons are the only allowed write mechanisms.

## 5. Important architecture decisions

### ADR-001: File-based JSON Database Repository
- **Decision**: Store projects and files in `projects_store.json` using atomic loading, modifying, and saving operations.
- **Reason**: Simplifies desktop deployments without DB servers, preparing interfaces for an easy future SQLite migration.
- **Impact**: Concurrency is managed via Rust load-modify-save cycles. Simultaneous database writes must be controlled.

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

## 8. Do not break

- **No solid borders theme**: Do not add solid grey or black borders separating card contents. Use shifts in shades of blue/grey.
- **No direct AI database updates**: Never write backend file-writing commands triggered directly by AI agents without intermediate form confirmation.
- **No-physical file deletion in scans**: Directory scans must only toggle the `exists` flag on linked files, never delete files physically.

## 9. Open questions

1. *Should we migrate the JSON storage to SQLite in the next phase?*
2. *Do we need to support custom discount rates per year in the 10-year cashflow simulation instead of a single flat rate?*
