# AGENTS.md

## Scope

This file applies to the entire repository.

## Context initialization

Before starting any task, follow these on-demand reading rules to load the context:

1. **Every task must read (lightweight entrance)**:
   * [PROJECT_INDEX.md](./docs/PROJECT_INDEX.md)
   * [CURRENT_TASK.md](./docs/CURRENT_TASK.md)
2. **On-demand reads depending on task scope**:
   * If modifying appearance, style, or themes: Read [docs/modules/appearance.md](./docs/modules/appearance.md) and [DESIGN.md](./DESIGN.md).
   * If modifying global architecture or core flow: Read [docs/ARCHITECTURE_MAP.md](./docs/ARCHITECTURE_MAP.md).
   * Do **not** read [docs/CHANGELOG_AI.md](./docs/CHANGELOG_AI.md), [docs/PROJECT_STATUS.md](./docs/PROJECT_STATUS.md), or [docs/AI_CONTEXT.md](./docs/AI_CONTEXT.md) unless tracing history/regressions or specifically instructed.

Do not perform a full repository scan at the beginning of a task. Use the entry files first, then read only the source files relevant to the current task.

## Project principles

Lamber is a sales support toolkit built on **Tauri + React + Rust**, designed to automate project benefit calculations, generate bidding and project review documents (Word/Excel) using structured templates, and manage 5G/ICT project lifecycles.

The project should favor stable, maintainable, business-safe implementation over quick local patches.
When a bug exposes fragile architecture, duplicated state, inconsistent data flow, or hidden coupling, prefer a controlled structural fix instead of adding another narrow workaround.

## Working rules

* **Incremental but robust changes**: Prefer incremental modifications over broad rewrites, but do not optimize only for the smallest diff. The goal is the most stable and maintainable fix within a controlled scope.
* **Root-cause first**: Before fixing a bug, identify the underlying cause by tracing the relevant data flow, state flow, lifecycle, persistence path, and UI interaction path.
* **Reduce related technical debt**: If the task reveals directly related duplicated logic, fragile assumptions, hidden coupling, inconsistent state handling, or scattered fallback behavior, clean it up as part of the same fix.
* **Avoid superficial patches**: Do not rely on narrow `if` guards, timeout hacks, one-off fallbacks, duplicated condition branches, or local-only workarounds unless they are clearly the safest long-term solution.
* **Keep scope controlled**: Robust repair does not mean unlimited refactoring. Refactor only the components, modules, or data paths directly involved in the problem. Preserve public behavior unless the requested change explicitly requires otherwise.
* **Explain repair tradeoffs**: For meaningful fixes, briefly explain the root cause, why a minimal patch would be fragile, what technical debt is being reduced, and why the chosen scope is safe.
* **Preserve business logic**: Keep existing formulas, rules, and logic unless explicitly requested.
* **Financial logic alignment**: When editing calculation math, make sure not to break cashflow, NPV, tax separation, selection fee, or reverse-calculation logic.
* **Validate the effective write**: Candidate evaluation and validation must use the exact state that the commit will write, including normalization, linked subjects, split invalidation, and funding plans. Share the transformation with the writer; a post-write warning cannot replace a pre-write rejection.
* **Discrete target reachability**: Being within a sampled minimum/maximum does not prove that a target is reachable. Validate the final metric against the unchanged tolerance before writing; never rewrite the user's target to make a result pass.
* **0-tolerance reconciliation check**: Respect the 0-tolerance financial check before transitions to cashflow tables or document generation.
* **AI features control**: Ensure the AI copilot only reads serialized business states and never writes core project data directly without user action, such as manual clicking or confirmations.
* **UI Design System**: Use the pinned official dsh WebUI as the design baseline for all of Lamber, as specified in [DESIGN.md](./DESIGN.md). Align the main application and AI extensions with its component roles, semantic colors, corner hierarchy, typography, spacing, and interaction states through shared tokens/components. The former Architectural Ledger, ROUND_FOUR, and separate blue-grey palette are no longer target standards. Preserve Lamber branding, business semantics, accessibility, and tabular numbers for financial displays; do not restyle upstream WebUI to match the old Lamber design.

## Fix strategy

When fixing bugs or implementing changes, do not default to the smallest possible patch if the issue reveals deeper architectural, state-management, data-flow, lifecycle, persistence, or maintainability problems.

The preferred sequence is:

1. **Understand the failure path**

   * Locate where the bug is introduced.
   * Trace how data, state, props, events, persistence, and side effects move through the relevant modules.
   * Check whether the issue is caused by duplicated sources of truth, stale derived state, inconsistent serialization, race conditions, fragile UI assumptions, or backend/frontend mismatch.

2. **Choose the most stable bounded fix**

   * Prefer a fix that removes the failure mode instead of masking it.
   * Consolidate duplicated logic when duplication is part of the bug.
   * Normalize state ownership when multiple modules are competing as sources of truth.
   * Strengthen interfaces or data contracts when the bug comes from loose assumptions.
   * Replace brittle timing/order dependencies with explicit lifecycle or state synchronization.

3. **Avoid debt-increasing shortcuts**

   * Do not add a new workaround on top of an already fragile path.
   * Do not copy-paste existing flawed logic into another module.
   * Do not patch only the visible symptom while leaving the same bug trigger active elsewhere.
   * Do not preserve a broken abstraction only because it produces a smaller diff.

4. **Validate the repair**

   * Add or update tests when practical.
   * If automated tests are not available, provide clear manual validation steps.
   * Cover the original bug, related regression points, and at least one realistic business workflow.

In short: **do not optimize for the smallest diff; optimize for the most stable, maintainable fix within a controlled scope.**

## Required documentation updates

After meaningful code or visual changes, update documentation as follows:

* Update [CURRENT_TASK.md](./docs/CURRENT_TASK.md) to record progress.
* Update corresponding [docs/modules/*.md](./docs/modules/) design documentation (e.g., [appearance.md](./docs/modules/appearance.md)) for key modular designs/rules.
* Add a summary record in [CHANGELOG_AI.md](./docs/CHANGELOG_AI.md).
* Do **not** update `PROJECT_STATUS.md` or `AI_CONTEXT.md` unless explicitly required; these files are now kept for historical reference only.

Record only long-term valuable project knowledge. Do not log transient debugging details.

## Forbidden behavior

* **No direct AI database writes**: Do not bypass user confirmation to let the AI directly modify `projects_store.json`.
* **No high-saturation UI backgrounds**: Do not restore large-area high-saturation blue/purple background panels. Follow the pinned dsh WebUI semantic surfaces across the product.
* **No arbitrary visual separators**: Use the corresponding dsh component's semantic surfaces, separators, and shadows through the shared design mapping; do not add arbitrary dark borders or impose the old Lamber surface palette on official components.
* **No dependency bloating**: Do not introduce large third-party crates or npm packages without explicit approval.
* **No debt-increasing quick patches**: Do not choose a minimal local patch when the same effort can safely remove the related failure path or reduce directly connected technical debt.
