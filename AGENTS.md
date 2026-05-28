# AGENTS.md

## Scope

This file applies to the entire repository.

## Context initialization

Before starting any task, read these files in order:

1. [AI_CONTEXT.md](./docs/AI_CONTEXT.md)
2. [PROJECT_STATUS.md](./docs/PROJECT_STATUS.md)
3. [ARCHITECTURE_MAP.md](./docs/ARCHITECTURE_MAP.md)
4. [CHANGELOG_AI.md](./docs/CHANGELOG_AI.md)

Do not perform a full repository scan at the beginning of a task.  
Use the context files first, then read only the source files relevant to the current task.

## Project principles

Lamber is a sales support toolkit built on **Tauri + React + Rust**, designed to automate project benefit calculations, generate bidding and project review documents (Word/Excel) using structured templates, and manage 5G/ICT project lifecycles.

## Working rules

- **Incremental changes**: Prefer incremental modifications over broad rewrites.
- **Preserve business logic**: Keep existing formulas, rules, and logic unless explicitly requested.
- **Financial logic alignment**: When editing calculation math, make sure not to break cashflow, NPV, tax separation, selection fee, or reverse-calculation logic.
- **0-tolerance reconciliation check**: Respect the 0-tolerance financial check before transitions to cashflow tables or document generation.
- **AI features control**: Ensure the AI copilot only reads serialized business states and never writes core project data directly without user action (such as manual clicking or confirmations).
- **UI Design System**: Follow "The Architectural Ledger" guidelines in [DESIGN.md](./DESIGN.md). Adhere to the "No-Line Rule" (tonal surface changes instead of borders), using rounded corners (`ROUND_FOUR`), the `Inter` font family, and tabular numbers for numerical displays.

## Required documentation updates

After meaningful code changes, update these files in order:

- [PROJECT_STATUS.md](./docs/PROJECT_STATUS.md)
- [ARCHITECTURE_MAP.md](./docs/ARCHITECTURE_MAP.md)
- [CHANGELOG_AI.md](./docs/CHANGELOG_AI.md)
- [AI_CONTEXT.md](./docs/AI_CONTEXT.md)

Record only long-term valuable project knowledge. Do not log transient debugging details.

## Forbidden behavior

- **No direct AI database writes**: Do not bypass user confirmation to let the AI directly modify `projects_store.json`.
- **No high-saturation UI backgrounds**: Do not restore large-area high-saturation blue/purple background panels. Respect the grey/pale-blue design system.
- **No traditional 1px borders**: Avoid using arbitrary `border` classes; use the layout surface shift tokens (`bg-muted`, nesting, or HSL container backgrounds) to demarcate sections.
- **No dependency bloating**: Do not introduce large third-party crates or npm packages without explicit approval.
