# 常用资料与项目预设模块

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
- Selection result sign-off form: project background, selected partner, selection content, scope, industry/scenario, method, rule, standard-plan description, revenue collection method, and expenditure payment method.
