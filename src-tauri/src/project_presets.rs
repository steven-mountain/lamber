use rusqlite::{params, Connection, OptionalExtension};
use serde::{Deserialize, Serialize};
use std::collections::BTreeSet;
use std::sync::Arc;
use tauri::State;

fn now_iso() -> String {
    chrono::Utc::now().to_rfc3339()
}

fn generate_id(prefix: &str) -> String {
    format!("{}_{}", prefix, uuid::Uuid::new_v4().simple())
}

fn normalize_list(values: Vec<String>) -> Vec<String> {
    let values = values
        .into_iter()
        .map(|value| value.trim().to_string())
        .filter(|value| !value.is_empty())
        .collect::<BTreeSet<_>>();
    values.iter().cloned().collect()
}

const PROJECT_PRESET_FIELD_KEYS: &[&str] = &[
    "project_basic.customer_name",
    "project_basic.background",
    "project_basic.solution",
    "project_basic.property_rights",
    "approval.reviewers",
    "approval.department",
    "approval.branch_attendees",
    "approval.project_manager",
    "approval.it_service_content",
    "approval.ct_service_content",
    "demand.unit",
    "demand.service_content",
    "demand.customer_confirmation",
    "demand.deployment_environment",
    "meeting.onsite_support",
    "meeting.it_construction_content",
    "meeting.ct_construction_content",
    "meeting.time_requirement",
    "meeting.threeization",
    "meeting.strategic_value",
    "meeting.technical_conclusion",
    "meeting.review_accuracy",
    "payment.revenue_collection_method",
    "payment.expenditure_payment_method",
    "service.description",
    "risk.description",
    "procurement.single_source_basis",
    "procurement.other_method",
    "implementation.construction_interface",
    "demand.device_list",
    "demand.security_detail",
    "template.it_business_mode",
    "template.it_funding_source",
    "demand.it_business_mode",
    "procurement.method",
    "tender.is_joint",
    "procurement.single_source",
];

fn validate_field_key(field_key: &str) -> Result<(), String> {
    if PROJECT_PRESET_FIELD_KEYS.contains(&field_key) {
        Ok(())
    } else {
        Err(format!("ProjectPresetFieldNotEligible::{}", field_key))
    }
}

fn validate_scope(scope: &str) -> Result<(), String> {
    match scope {
        "workspace" => Ok(()),
        "user" => Err("UserScopedProjectPresetsNotImplemented".to_string()),
        _ => Err(format!("UnsupportedProjectPresetScope::{}", scope)),
    }
}

fn validate_value_type(value_type: &str) -> Result<(), String> {
    match value_type {
        "text" | "long_text" | "dictionary_value" | "boolean" => Ok(()),
        _ => Err(format!("UnsupportedProjectPresetValueType::{}", value_type)),
    }
}

fn validate_source_type(source_type: &str) -> Result<(), String> {
    match source_type {
        "manual" | "from_project" | "preset_item" | "dictionary" => Ok(()),
        _ => Err(format!(
            "UnsupportedProjectPresetSourceType::{}",
            source_type
        )),
    }
}

pub(crate) fn ensure_schema(conn: &Connection) -> rusqlite::Result<()> {
    conn.execute_batch(
        "CREATE TABLE IF NOT EXISTS project_preset_templates (
            id TEXT PRIMARY KEY,
            scope TEXT NOT NULL DEFAULT 'workspace',
            name TEXT NOT NULL,
            description TEXT,
            category TEXT NOT NULL DEFAULT '',
            tags_json TEXT NOT NULL DEFAULT '[]',
            enabled INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            deleted_at TEXT
        );
        CREATE TABLE IF NOT EXISTS project_preset_template_entries (
            id TEXT PRIMARY KEY,
            template_id TEXT NOT NULL,
            field_key TEXT NOT NULL,
            value_json TEXT NOT NULL,
            value_type TEXT NOT NULL,
            source_type TEXT NOT NULL DEFAULT 'manual',
            sort_order INTEGER NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            UNIQUE(template_id, field_key),
            FOREIGN KEY(template_id) REFERENCES project_preset_templates(id) ON DELETE CASCADE
        );
        CREATE INDEX IF NOT EXISTS idx_project_preset_templates_scope
            ON project_preset_templates(scope, enabled, updated_at);
        CREATE INDEX IF NOT EXISTS idx_project_preset_entries_template
            ON project_preset_template_entries(template_id, sort_order);",
    )
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ProjectPresetTemplateEntry {
    pub id: String,
    pub template_id: String,
    pub field_key: String,
    pub value: serde_json::Value,
    pub value_type: String,
    pub source_type: String,
    pub sort_order: i64,
    pub created_at: String,
    pub updated_at: String,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ProjectPresetTemplate {
    pub id: String,
    pub scope: String,
    pub name: String,
    pub description: Option<String>,
    pub category: String,
    pub tags: Vec<String>,
    pub enabled: bool,
    pub created_at: String,
    pub updated_at: String,
    pub entries: Vec<ProjectPresetTemplateEntry>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProjectPresetTemplateEntryInput {
    pub id: Option<String>,
    pub field_key: String,
    pub value: serde_json::Value,
    pub value_type: String,
    pub source_type: Option<String>,
    pub sort_order: Option<i64>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProjectPresetTemplateInput {
    pub id: Option<String>,
    pub scope: Option<String>,
    pub name: String,
    pub description: Option<String>,
    pub category: Option<String>,
    pub tags: Option<Vec<String>>,
    pub enabled: Option<bool>,
    pub entries: Vec<ProjectPresetTemplateEntryInput>,
}

fn row_to_entry(row: &rusqlite::Row<'_>) -> rusqlite::Result<ProjectPresetTemplateEntry> {
    let value_json: String = row.get(3)?;
    Ok(ProjectPresetTemplateEntry {
        id: row.get(0)?,
        template_id: row.get(1)?,
        field_key: row.get(2)?,
        value: serde_json::from_str(&value_json).unwrap_or(serde_json::Value::Null),
        value_type: row.get(4)?,
        source_type: row.get(5)?,
        sort_order: row.get(6)?,
        created_at: row.get(7)?,
        updated_at: row.get(8)?,
    })
}

fn list_entries_locked(
    conn: &Connection,
    template_id: &str,
) -> Result<Vec<ProjectPresetTemplateEntry>, String> {
    let mut stmt = conn
        .prepare(
            "SELECT id, template_id, field_key, value_json, value_type, source_type,
                sort_order, created_at, updated_at
             FROM project_preset_template_entries
             WHERE template_id = ?1
             ORDER BY sort_order ASC, created_at ASC",
        )
        .map_err(|error| error.to_string())?;
    let entries = stmt
        .query_map([template_id], row_to_entry)
        .map_err(|error| error.to_string())?
        .collect::<rusqlite::Result<Vec<_>>>()
        .map_err(|error| error.to_string())?;
    Ok(entries)
}

fn get_template_locked(
    conn: &Connection,
    template_id: &str,
) -> Result<ProjectPresetTemplate, String> {
    let template = conn
        .query_row(
            "SELECT id, scope, name, description, category, tags_json, enabled,
                created_at, updated_at
             FROM project_preset_templates
             WHERE id = ?1 AND deleted_at IS NULL",
            [template_id],
            |row| {
                let tags_json: String = row.get(5)?;
                Ok(ProjectPresetTemplate {
                    id: row.get(0)?,
                    scope: row.get(1)?,
                    name: row.get(2)?,
                    description: row.get(3)?,
                    category: row.get(4)?,
                    tags: serde_json::from_str(&tags_json).unwrap_or_default(),
                    enabled: row.get::<_, i64>(6)? != 0,
                    created_at: row.get(7)?,
                    updated_at: row.get(8)?,
                    entries: Vec::new(),
                })
            },
        )
        .optional()
        .map_err(|error| error.to_string())?
        .ok_or_else(|| format!("ProjectPresetTemplateNotFound::{}", template_id))?;
    Ok(ProjectPresetTemplate {
        entries: list_entries_locked(conn, template_id)?,
        ..template
    })
}

fn list_templates_locked(
    conn: &Connection,
    include_disabled: bool,
) -> Result<Vec<ProjectPresetTemplate>, String> {
    ensure_schema(conn).map_err(|error| error.to_string())?;
    let mut stmt = conn
        .prepare(
            "SELECT id FROM project_preset_templates
             WHERE scope = 'workspace' AND deleted_at IS NULL
               AND (?1 = 1 OR enabled = 1)
             ORDER BY updated_at DESC, name ASC",
        )
        .map_err(|error| error.to_string())?;
    let ids = stmt
        .query_map([if include_disabled { 1 } else { 0 }], |row| {
            row.get::<_, String>(0)
        })
        .map_err(|error| error.to_string())?
        .collect::<rusqlite::Result<Vec<_>>>()
        .map_err(|error| error.to_string())?;
    ids.iter().map(|id| get_template_locked(conn, id)).collect()
}

fn save_template_locked(
    conn: &mut Connection,
    input: ProjectPresetTemplateInput,
) -> Result<ProjectPresetTemplate, String> {
    ensure_schema(conn).map_err(|error| error.to_string())?;
    let scope = input.scope.unwrap_or_else(|| "workspace".to_string());
    validate_scope(&scope)?;
    let name = input.name.trim().to_string();
    if name.is_empty() {
        return Err("ProjectPresetTemplateNameRequired".to_string());
    }
    if input.entries.is_empty() {
        return Err("ProjectPresetTemplateEntriesRequired".to_string());
    }

    let mut seen = BTreeSet::new();
    for entry in &input.entries {
        let field_key = entry.field_key.trim();
        validate_field_key(field_key)?;
        validate_value_type(entry.value_type.trim())?;
        validate_source_type(entry.source_type.as_deref().unwrap_or("manual"))?;
        if entry.value.is_null()
            || entry
                .value
                .as_str()
                .map(|value| value.trim().is_empty())
                .unwrap_or(false)
        {
            return Err(format!("ProjectPresetEntryValueRequired::{}", field_key));
        }
        if !seen.insert(field_key.to_string()) {
            return Err(format!("DuplicateProjectPresetField::{}", field_key));
        }
    }

    let now = now_iso();
    let template_id = input.id.unwrap_or_else(|| generate_id("project_preset"));
    let exists: bool = conn
        .query_row(
            "SELECT EXISTS(
                SELECT 1 FROM project_preset_templates
                WHERE id = ?1 AND deleted_at IS NULL
             )",
            [&template_id],
            |row| row.get(0),
        )
        .map_err(|error| error.to_string())?;
    let description = input
        .description
        .map(|value| value.trim().to_string())
        .filter(|value| !value.is_empty());
    let category = input.category.unwrap_or_default().trim().to_string();
    let tags = normalize_list(input.tags.unwrap_or_default());
    let tags_json = serde_json::to_string(&tags).map_err(|error| error.to_string())?;
    let enabled = input.enabled.unwrap_or(true);

    let tx = conn.transaction().map_err(|error| error.to_string())?;
    if exists {
        tx.execute(
            "UPDATE project_preset_templates
             SET scope = ?1, name = ?2, description = ?3, category = ?4,
                tags_json = ?5, enabled = ?6, updated_at = ?7
             WHERE id = ?8 AND deleted_at IS NULL",
            params![
                scope,
                name,
                description,
                category,
                tags_json,
                if enabled { 1 } else { 0 },
                now,
                template_id,
            ],
        )
        .map_err(|error| error.to_string())?;
        tx.execute(
            "DELETE FROM project_preset_template_entries WHERE template_id = ?1",
            [&template_id],
        )
        .map_err(|error| error.to_string())?;
    } else {
        tx.execute(
            "INSERT INTO project_preset_templates (
                id, scope, name, description, category, tags_json, enabled,
                created_at, updated_at
             ) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?8)",
            params![
                template_id,
                scope,
                name,
                description,
                category,
                tags_json,
                if enabled { 1 } else { 0 },
                now,
            ],
        )
        .map_err(|error| error.to_string())?;
    }

    for (index, entry) in input.entries.into_iter().enumerate() {
        let entry_id = entry
            .id
            .unwrap_or_else(|| generate_id("project_preset_entry"));
        tx.execute(
            "INSERT INTO project_preset_template_entries (
                id, template_id, field_key, value_json, value_type, source_type,
                sort_order, created_at, updated_at
             ) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?8)",
            params![
                entry_id,
                template_id,
                entry.field_key.trim(),
                serde_json::to_string(&entry.value).map_err(|error| error.to_string())?,
                entry.value_type.trim(),
                entry.source_type.unwrap_or_else(|| "manual".to_string()),
                entry.sort_order.unwrap_or((index as i64 + 1) * 10),
                now,
            ],
        )
        .map_err(|error| error.to_string())?;
    }
    tx.commit().map_err(|error| error.to_string())?;
    get_template_locked(conn, &template_id)
}

pub(crate) fn initialize_new_project_locked(
    conn: &Connection,
    project_id: &str,
    project_name: &str,
    template_id: &str,
) -> Result<(), String> {
    ensure_schema(conn).map_err(|error| error.to_string())?;
    let template = get_template_locked(conn, template_id)?;
    if !template.enabled {
        return Err("ProjectPresetTemplateDisabled".to_string());
    }

    let mut customer_name = None;
    let mut project_background = None;
    let mut property_rights = None;
    for entry in &template.entries {
        let value = entry
            .value
            .as_str()
            .map(str::to_string)
            .unwrap_or_else(|| entry.value.to_string());
        match entry.field_key.as_str() {
            "project_basic.customer_name" => customer_name = Some(value),
            "project_basic.background" => project_background = Some(value),
            "project_basic.property_rights" => property_rights = Some(value),
            _ => {}
        }
    }

    let existing_customer_name: String = conn
        .query_row(
            "SELECT customer_name FROM projects WHERE id = ?1",
            [project_id],
            |row| row.get(0),
        )
        .map_err(|error| error.to_string())?;
    let should_seed_customer =
        existing_customer_name.trim().is_empty() || existing_customer_name == "未知客户";
    if should_seed_customer {
        if let Some(customer_name) = customer_name.as_deref() {
            conn.execute(
                "UPDATE projects SET customer_name = ?1, updated_at = ?2 WHERE id = ?3",
                params![customer_name, now_iso(), project_id],
            )
            .map_err(|error| error.to_string())?;
        }
    }

    let now = now_iso();
    let profile = serde_json::json!({
        "projectName": project_name,
        "customerName": if should_seed_customer {
            customer_name.unwrap_or(existing_customer_name)
        } else {
            existing_customer_name
        },
        "propertyRights": property_rights.unwrap_or_default(),
    });
    let background = serde_json::json!({
        "projectBackground": project_background.unwrap_or_default(),
    });
    let input_payload = serde_json::json!({
        "project_name": project_name,
        "customer_name": profile["customerName"],
        "property_rights": profile["propertyRights"],
        "project_background": background["projectBackground"],
    });
    conn.execute(
        "INSERT INTO project_lifecycle_states (
            id, project_id, lifecycle_version, profile_json, parameters_json,
            background_json, input_payload_json, created_at, updated_at
         ) VALUES (?1, ?2, 1, ?3, '{}', ?4, ?5, ?6, ?6)",
        params![
            generate_id("lifecycle"),
            project_id,
            profile.to_string(),
            background.to_string(),
            input_payload.to_string(),
            now,
        ],
    )
    .map_err(|error| error.to_string())?;

    let seed = serde_json::json!({
        "templateId": template.id,
        "templateName": template.name,
        "entries": template.entries,
    });
    conn.execute(
        "INSERT INTO project_settings (project_id, key, value, updated_at)
         VALUES (?1, 'project_preset_seed', ?2, ?3)
         ON CONFLICT(project_id, key) DO UPDATE SET
            value = excluded.value,
            updated_at = excluded.updated_at",
        params![project_id, seed.to_string(), now],
    )
    .map_err(|error| error.to_string())?;
    Ok(())
}

#[tauri::command]
pub async fn list_project_preset_templates(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    include_disabled: Option<bool>,
) -> Result<Vec<ProjectPresetTemplate>, String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|error| error.to_string())?;
    list_templates_locked(&conn, include_disabled.unwrap_or(false))
}

#[tauri::command]
pub async fn save_project_preset_template(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    template: ProjectPresetTemplateInput,
) -> Result<ProjectPresetTemplate, String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let mut conn = db.lock().map_err(|error| error.to_string())?;
    save_template_locked(&mut conn, template)
}

#[tauri::command]
pub async fn set_project_preset_template_enabled(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    id: String,
    enabled: bool,
) -> Result<ProjectPresetTemplate, String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|error| error.to_string())?;
    let changed = conn
        .execute(
            "UPDATE project_preset_templates
             SET enabled = ?1, updated_at = ?2
             WHERE id = ?3 AND deleted_at IS NULL",
            params![if enabled { 1 } else { 0 }, now_iso(), id],
        )
        .map_err(|error| error.to_string())?;
    if changed == 0 {
        return Err(format!("ProjectPresetTemplateNotFound::{}", id));
    }
    get_template_locked(&conn, &id)
}

#[tauri::command]
pub async fn delete_project_preset_template(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    id: String,
) -> Result<(), String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|error| error.to_string())?;
    let now = now_iso();
    let changed = conn
        .execute(
            "UPDATE project_preset_templates
             SET enabled = 0, deleted_at = ?1, updated_at = ?1
             WHERE id = ?2 AND deleted_at IS NULL",
            params![now, id],
        )
        .map_err(|error| error.to_string())?;
    if changed == 0 {
        return Err(format!("ProjectPresetTemplateNotFound::{}", id));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sample_input() -> ProjectPresetTemplateInput {
        ProjectPresetTemplateInput {
            id: None,
            scope: None,
            name: "测试项目预设".to_string(),
            description: Some("测试".to_string()),
            category: Some("ICT".to_string()),
            tags: Some(vec!["常用".to_string()]),
            enabled: Some(true),
            entries: vec![
                ProjectPresetTemplateEntryInput {
                    id: None,
                    field_key: "project_basic.customer_name".to_string(),
                    value: serde_json::json!("某客户"),
                    value_type: "text".to_string(),
                    source_type: Some("from_project".to_string()),
                    sort_order: None,
                },
                ProjectPresetTemplateEntryInput {
                    id: None,
                    field_key: "procurement.method".to_string(),
                    value: serde_json::json!("采购"),
                    value_type: "dictionary_value".to_string(),
                    source_type: Some("dictionary".to_string()),
                    sort_order: None,
                },
            ],
        }
    }

    #[test]
    fn project_preset_crud_and_safety_work() {
        let mut conn = Connection::open_in_memory().unwrap();
        conn.execute("PRAGMA foreign_keys = ON", []).unwrap();
        ensure_schema(&conn).unwrap();
        let created = save_template_locked(&mut conn, sample_input()).unwrap();
        assert_eq!(created.entries.len(), 2);
        assert_eq!(list_templates_locked(&conn, false).unwrap().len(), 1);

        let mut invalid = sample_input();
        invalid.entries[0].field_key = "finance.npv".to_string();
        assert!(save_template_locked(&mut conn, invalid)
            .unwrap_err()
            .contains("ProjectPresetFieldNotEligible"));
    }

    #[test]
    fn new_project_initialization_writes_safe_form_state() {
        let mut conn = Connection::open_in_memory().unwrap();
        conn.execute("PRAGMA foreign_keys = ON", []).unwrap();
        conn.execute_batch(
            "CREATE TABLE projects (
                id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                customer_name TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );
            CREATE TABLE project_lifecycle_states (
                id TEXT PRIMARY KEY,
                project_id TEXT NOT NULL UNIQUE,
                lifecycle_version INTEGER NOT NULL,
                profile_json TEXT NOT NULL,
                parameters_json TEXT NOT NULL,
                background_json TEXT NOT NULL,
                input_payload_json TEXT NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
            );
            CREATE TABLE project_settings (
                project_id TEXT NOT NULL,
                key TEXT NOT NULL,
                value TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                PRIMARY KEY(project_id, key),
                FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
            );",
        )
        .unwrap();
        ensure_schema(&conn).unwrap();
        conn.execute(
            "INSERT INTO projects (id, name, customer_name, updated_at)
             VALUES ('p1', '项目一', '未知客户', 'now')",
            [],
        )
        .unwrap();
        let created = save_template_locked(&mut conn, sample_input()).unwrap();

        initialize_new_project_locked(&conn, "p1", "项目一", &created.id).unwrap();

        let customer: String = conn
            .query_row(
                "SELECT customer_name FROM projects WHERE id = 'p1'",
                [],
                |row| row.get(0),
            )
            .unwrap();
        assert_eq!(customer, "某客户");
        let profile: String = conn
            .query_row(
                "SELECT profile_json FROM project_lifecycle_states WHERE project_id = 'p1'",
                [],
                |row| row.get(0),
            )
            .unwrap();
        assert!(profile.contains("某客户"));
        let seed: String = conn
            .query_row(
                "SELECT value FROM project_settings
                 WHERE project_id = 'p1' AND key = 'project_preset_seed'",
                [],
                |row| row.get(0),
            )
            .unwrap();
        assert!(seed.contains("procurement.method"));
    }
}
