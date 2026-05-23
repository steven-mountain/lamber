# CHANGELOG_AI.md

This changelog records structural modifications, business rules, and context changes made by AI agents to maintain a reliable project state mapping.

## 2026-05-23

### Initial project state mapping

Created:
- [AGENTS.md](file:///d:/HermesJang/CMCC/tools/lambert/lamber/AGENTS.md)
- [AI_CONTEXT.md](file:///d:/HermesJang/CMCC/tools/lambert/lamber/docs/AI_CONTEXT.md)
- [PROJECT_STATUS.md](file:///d:/HermesJang/CMCC/tools/lambert/lamber/docs/PROJECT_STATUS.md)
- [ARCHITECTURE_MAP.md](file:///d:/HermesJang/CMCC/tools/lambert/lamber/docs/ARCHITECTURE_MAP.md)
- [CHANGELOG_AI.md](file:///d:/HermesJang/CMCC/tools/lambert/lamber/docs/CHANGELOG_AI.md)

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
