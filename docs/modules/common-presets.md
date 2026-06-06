# 常用资料与项目预设模块

## Phase 1.5 Field-Level Presets

Phase 1.5 adds an opt-in field capability without changing the ownership of formal project data.

- `src-ui/src/lib/presetFieldKeys.ts` is the central field metadata registry. Each entry includes `fieldKey`, user-facing label/description, templates, groups, `fieldType`, `presetEligible`, recommended categories, optional aliases, kind, and default enabled state.
- Unregistered fields are ineligible by default. User-visible fallback text is “未命名字段 / 暂未配置”; raw field keys are reserved for internal matching and diagnostics.
- Existing Phase 1 fields remain enabled by default for compatibility. Newly connected ordinary fields may set `defaultEnabled: false` and show “+ 预设” until the user enables them.
- Workspace-specific overrides are stored in SQLite `preset_field_settings(field_key, enabled, updated_at)`, introduced in schema version 7.
- `CommonPresetQuickFill` owns the shared enable/select/save/replace/append/disable interaction. Multiple rendered controls for the same fieldKey synchronize through one workspace-scoped setting.
- `common_presets.rs` enforces an eligible field allowlist for both field enablement and `applicableFieldKeys` writes. Amount, percent, and computed fields cannot be bound by bypassing the UI.

The preset center and field-side panels display the field label, applicable templates, and business groups. For example, the shared payment keys are presented as “收入条款” and “支出条款”, applicable to “立项签批表、会审纪要” under “商务条款”.

Explicitly excluded field types include `amount`, `percent`, and `computed`, covering revenue/cost amounts, tax rates, ratios, annual cashflow, NPV, margin, reverse-calculation targets, and balancing amounts.

The first opt-in representative field is ICT “产权归属”. Phase 1 connected fields remain available, including project background/solution, reviewers, department, project owner, customer unit, income/expenditure terms, service content, and related meeting/sign-off text.

### Field Capability Close Behavior

Enabled fields expose a compact more-actions menu beside “选择常用 / 保存当前”. “关闭预设” writes only `preset_field_settings.enabled = false` for the current stable fieldKey.

Closing a field capability must:
- Keep the current form value unchanged.
- Keep every `common_presets` record unchanged.
- Keep existing `applicableFieldKeys` bindings unchanged, including bindings shared with other fields.
- Synchronize other mounted controls using the same fieldKey.
- Restore the lightweight “+ 预设” state after closing.

The confirmation text must state that form content and reusable materials are retained. “选择常用” uses the `presetLibrary` bookmark icon; lightning, magic-wand, and AI-generation icons are not valid for reusable-material selection.

## Phase 1 Scope

Phase 1 implements reusable material management and field-level quick fill only.

Implemented:
- Workspace-scoped common short fields.
- Workspace-scoped common long text snippets.
- Independent Hub entry and `PresetCenterView` management page.
- SQLite persistence through `common_presets`.
- Stable fieldKey binding through `presetFieldKeys.ts`.
- Reusable `CommonPresetQuickFill` form-side picker.
- Active user-triggered save-current-field and fill actions.

Not implemented in Phase 1:
- Full project preset templates.
- Selecting a preset while creating a new project.
- One-click applying multiple fields.
- Automatic extraction from old projects.
- AI recommendation, AI auto-fill, or AI direct writing.

## SQLite Model

Table: `common_presets`

Important columns:
- `id`
- `scope`: currently `workspace`; `user` is reserved.
- `kind`: `short_value` or `text_snippet`.
- `category`
- `name`
- `content`
- `tags_json`
- `applicable_field_keys_json`
- `usage_count`
- `last_used_at`
- `enabled`
- `created_at`
- `updated_at`
- `deleted_at`

The table is initialized by `db.rs` and added in schema version 6. Because it lives inside the active Workspace `.lamber.sqlite`, normal workspace backup/export/import flows include it.

## FieldKey Rules

Field binding must use stable keys from `src-ui/src/lib/presetFieldKeys.ts`. Do not match by visible Chinese labels.

Initial keys:
- `project_basic.customer_name`
- `project_basic.background`
- `project_basic.solution`
- `approval.reviewers`
- `approval.department`
- `approval.project_manager`
- `approval.it_service_content`
- `approval.ct_service_content`
- `demand.unit`
- `demand.service_content`
- `demand.customer_confirmation`
- `demand.deployment_environment`
- `meeting.onsite_support`
- `meeting.it_construction_content`
- `meeting.ct_construction_content`
- `meeting.time_requirement`
- `payment.revenue_collection_method`
- `payment.expenditure_payment_method`
- `service.description`
- `risk.description`

An empty `applicableFieldKeys` list means the preset remains category/kind managed. A non-empty list limits quick-fill matching to those exact fields.

## Fill Flow

1. A field renders `CommonPresetQuickFill` with `fieldKey`, `kind`, current value, and an owning setter.
2. The component lists enabled matching presets from the current Workspace.
3. User selects replace or append. Replacement of non-empty long text requires explicit confirmation.
4. The component calls the owning setter only.
5. Existing project/template dirty state and save handlers persist the resulting formal field data.
6. The preset record's usage count and last-used timestamp are updated.

Presets are not read during Word/Excel generation, benefit calculation, or AI project context construction.

## Management Page Layout

`PresetCenterView` should prioritize the reusable-material list over the create/edit form. The default state is a list-first workspace with kind tabs, category filter, sort selector, search box, and a prominent create action. The full form is opened only when the user creates a new material or edits an existing one.

The right-side create/edit surface is an auxiliary panel, not a permanent second source of visual weight. On desktop it should stay around 360-420px wide while the list keeps the majority of the page. On narrow screens it should behave as an overlay drawer so the list is not squeezed into an unusable column. The panel owns its own scrolling and keeps save/cancel actions fixed at the bottom of the panel.

The panel must be constrained to the available viewport height; form sections scroll inside the panel and must not push the fixed action area below the visible page. Clicking outside the panel should request close on both desktop and narrow screens, while preserving the unsaved-change confirmation guard.

Preset cards should show enough metadata for scanning without requiring edit mode: name, category, enabled state, short content摘要, tags, applicable-field summary, usage count, and last-used time. Search is a frontend list filter over name, category, content, and tags; it must not change the persisted `common_presets` shape.

## Form-Side Layout Rule

Preset actions in forms should be attached through `CommonPresetFieldHeader`, which renders the visible field label and the `CommonPresetQuickFill` actions in one responsive header row. Plain neighboring fields should use `CommonPresetLabelHeader` where row alignment matters, so fields with and without presets share the same label-header height. Do not place preset buttons in a separate `justify-end` row below the label, and do not align them with hardcoded offsets or fixed pixel positioning. Narrow containers may wrap naturally, but normal desktop form fields should keep the label and preset actions on the same line.

## First Integrated Fields

- ICT basic customer name.
- ICT project background.
- Template project background.
- Template technical solution.
- Template meeting reviewers.
- Template branch/department name.
- Template risk/project owner.
- Demand import form: project demand unit, service content, customer confirmation, and deployment environment requirement.
- Meeting review form: onsite support staff, IT construction content, CT construction content, revenue collection method, expenditure payment method, and time requirement.
- Sign-off form: IT service content, CT service content, revenue collection method, and expenditure payment method.
