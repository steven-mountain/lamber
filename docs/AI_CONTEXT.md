# AI_CONTEXT.md

Last updated: 2026-05-23

## What this project is

Lamber is a lightweight sales support desktop tool built with **Tauri, React, and Rust**. It helps client managers and solutions experts manage ICT project lifecycles, perform economic benefit calculations (NPV, margin, cashflow), and fill out standardized bidding and project review documents (Word/Excel) using structured templates.

## Read order for new AI sessions

Before starting any task, read the status and architecture documentation in this order:

1. [PROJECT_STATUS.md](file:///d:/HermesJang/CMCC/tools/lambert/lamber/docs/PROJECT_STATUS.md) - Current development status, milestones, business rules, and constraints.
2. [ARCHITECTURE_MAP.md](file:///d:/HermesJang/CMCC/tools/lambert/lamber/docs/ARCHITECTURE_MAP.md) - Directory layout, Boot processes, data flow, and state serialization.
3. [CHANGELOG_AI.md](file:///d:/HermesJang/CMCC/tools/lambert/lamber/docs/CHANGELOG_AI.md) - History of AI contributions, decisions, and ongoing questions.
4. Relevant source files only (do not scan the entire codebase).

## Core modules

- **Project Board (项目看板)**: Visual hub displaying project indicators, project phases, and local folders linkage.
- **ICT Lifecycle Calculator (ICT生命周期测算)**: Main calculation engine handling 10-year cashflows, NPV, margins, and selection fee mappings.
- **AI Assistant / Copilot (智能顾问)**: Streamed (SSE) local LLM chatbot that pulls the active view's serialized state to give contextual recommendations based on a built-in capabilities database.
- **Template Word/Excel Engine (文档填报与填充)**: Backend Rust logic mapping form values and calculation results into docx variables and excel cells.
- **Local File & Scan Engine (本地文件扫描管理)**: Service linking projects to local directory folders, scanning files (Word/Excel/PDF), and managing linked vs. sandbox-copied storage modes.

## Current priorities

1. **AI Conversation History Window Context Optimization**: Optimize messaging queues and history truncation to prevent context window overflow.
2. **Local RAG Integration**: Allow users to upload local PDF/Markdown files as part of the AI Copilot's private knowledge base.
3. **App Build Size Optimization**: Clean up redundant npm and cargo dependencies to reduce the installer footprint.
4. **Internationalization (i18n)**: Implement `i18next` for Chinese and English multi-language toggling.

## High-risk areas

- **0-tolerance reconciliation check**: Any changes to input fields or calculations must satisfy `excl_tax = incl_tax / (1 + rate)` within a zero-tolerance margin. Mismatches block navigation and document generation.
- **Json Store load-modify-save cycle**: All repo files perform atomic load-modify-save cycles. Be very careful with concurrency when writing to `projects_store.json`.
- **Local Folder binding & scan synchronization**: Scanning folders updates physical file existence, but must never delete files physically on linked mode. Sandboxed (`copied`) files should be physically deleted only after user confirmation.
- **Built-in Product Recommendations**: Recommended products must be cross-checked with codes in `midThreeConstants.ts`. If matched, append `[系统内置]` label; otherwise, append `【系统外扩展】`.

## Where to start for common tasks

- **Project Board changes**:
  - Frontend: [ProjectBoard.tsx](file:///d:/HermesJang/CMCC/tools/lambert/lamber/src-ui/src/views/ProjectBoard.tsx)
  - Backend: [project_files/service.rs](file:///d:/HermesJang/CMCC/tools/lambert/lamber/src-tauri/src/project_files/service.rs)
- **Benefit Calculation math / NPV changes**:
  - Math Engine: [calculator.rs](file:///d:/HermesJang/CMCC/tools/lambert/lamber/src-tauri/src/benefit/calculator.rs)
  - Lifecycle View: [IctLifecycle.tsx](file:///d:/HermesJang/CMCC/tools/lambert/lamber/src-ui/src/views/IctLifecycle.tsx)
- **AI Chat Copilot changes**:
  - Chat Box Component: [AiChatPanel.tsx](file:///d:/HermesJang/CMCC/tools/lambert/lamber/src-ui/src/components/ai/AiChatPanel.tsx)
  - AI State Store: [useAiContextStore.ts](file:///d:/HermesJang/CMCC/tools/lambert/lamber/src-ui/src/store/useAiContextStore.ts)
  - LLM Runtime: [AiRuntime.ts](file:///d:/HermesJang/CMCC/tools/lambert/lamber/src-ui/src/ai/AiRuntime.ts)
- **Word/Excel Template Filling changes**:
  - Template forms: [TemplateForms.tsx](file:///d:/HermesJang/CMCC/tools/lambert/lamber/src-ui/src/views/TemplateForms.tsx)
  - Docfill service: [docfill.rs](file:///d:/HermesJang/CMCC/tools/lambert/lamber/src-tauri/src/docfill.rs)
  - Excel parser: [excel.rs](file:///d:/HermesJang/CMCC/tools/lambert/lamber/src-tauri/src/benefit/excel.rs)
- **UI Design modifications**:
  - Main styling: [index.css](file:///d:/HermesJang/CMCC/tools/lambert/lamber/src-ui/src/index.css)
  - Design Tokens: [DESIGN.md](file:///d:/HermesJang/CMCC/tools/lambert/lamber/DESIGN.md)

## Do not break

- Do not move the Project Board back under the ICT Lifecycle view. They must remain independent.
- Do not let the AI rewrite `projects_store.json` directly without an explicit user transaction.
- Do not replace the tonal-shift design system with high-saturation solid borders or solid purple/blue panels.
