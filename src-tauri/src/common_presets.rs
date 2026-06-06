use rusqlite::{params, Connection, OptionalExtension};
use serde::{Deserialize, Serialize};
use std::collections::BTreeSet;
use std::sync::Arc;
use tauri::State;

fn now_iso() -> String {
    chrono::Utc::now().to_rfc3339()
}

fn generate_id() -> String {
    format!("preset_{}", uuid::Uuid::new_v4().simple())
}

fn normalize_list(values: Vec<String>) -> Vec<String> {
    let mut set = BTreeSet::new();
    for value in values {
        let trimmed = value.trim();
        if !trimmed.is_empty() {
            set.insert(trimmed.to_string());
        }
    }
    set.into_iter().collect()
}

fn validate_kind(kind: &str) -> Result<(), String> {
    match kind {
        "short_value" | "text_snippet" => Ok(()),
        _ => Err(format!("UnsupportedPresetKind::{}", kind)),
    }
}

fn validate_scope(scope: &str) -> Result<(), String> {
    match scope {
        "workspace" => Ok(()),
        "user" => Err("UserScopedPresetsNotImplemented".to_string()),
        _ => Err(format!("UnsupportedPresetScope::{}", scope)),
    }
}

const PRESET_ELIGIBLE_FIELD_KEYS: &[&str] = &[
    "project_basic.customer_name",
    "project_basic.background",
    "project_basic.solution",
    "project_basic.property_rights",
    "approval.reviewers",
    "approval.department",
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
    "payment.revenue_collection_method",
    "payment.expenditure_payment_method",
    "contract.income_terms",
    "contract.expenditure_terms",
    "service.description",
    "risk.description",
];

fn validate_preset_field_key(field_key: &str) -> Result<(), String> {
    if PRESET_ELIGIBLE_FIELD_KEYS.contains(&field_key) {
        Ok(())
    } else {
        Err(format!("PresetFieldNotEligible::{}", field_key))
    }
}

fn validate_preset_field_keys(field_keys: &[String]) -> Result<(), String> {
    for field_key in field_keys {
        validate_preset_field_key(field_key)?;
    }
    Ok(())
}

pub(crate) fn ensure_schema(conn: &Connection) -> rusqlite::Result<()> {
    conn.execute(
        "CREATE TABLE IF NOT EXISTS common_presets (
            id TEXT PRIMARY KEY,
            scope TEXT NOT NULL DEFAULT 'workspace',
            kind TEXT NOT NULL CHECK(kind IN ('short_value', 'text_snippet')),
            category TEXT NOT NULL,
            name TEXT NOT NULL,
            content TEXT NOT NULL,
            tags_json TEXT NOT NULL DEFAULT '[]',
            applicable_field_keys_json TEXT NOT NULL DEFAULT '[]',
            usage_count INTEGER NOT NULL DEFAULT 0,
            last_used_at TEXT,
            enabled INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            deleted_at TEXT
        );",
        [],
    )?;
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_common_presets_scope_kind_category ON common_presets(scope, kind, category);",
        [],
    )?;
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_common_presets_usage ON common_presets(scope, usage_count, last_used_at);",
        [],
    )?;
    conn.execute(
        "CREATE TABLE IF NOT EXISTS preset_field_settings (
            field_key TEXT PRIMARY KEY,
            enabled INTEGER NOT NULL DEFAULT 0,
            updated_at TEXT NOT NULL
        );",
        [],
    )?;
    Ok(())
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CommonPreset {
    pub id: String,
    pub scope: String,
    pub kind: String,
    pub category: String,
    pub name: String,
    pub content: String,
    pub tags: Vec<String>,
    pub applicable_field_keys: Vec<String>,
    pub usage_count: i64,
    pub last_used_at: Option<String>,
    pub enabled: bool,
    pub created_at: String,
    pub updated_at: String,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CommonPresetInput {
    pub id: Option<String>,
    pub scope: Option<String>,
    pub kind: String,
    pub category: String,
    pub name: String,
    pub content: String,
    pub tags: Option<Vec<String>>,
    pub applicable_field_keys: Option<Vec<String>>,
    pub enabled: Option<bool>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CommonPresetFilter {
    pub kind: Option<String>,
    pub category: Option<String>,
    pub field_key: Option<String>,
    pub include_disabled: Option<bool>,
    pub sort_by: Option<String>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PresetFieldSetting {
    pub field_key: String,
    pub enabled: bool,
    pub updated_at: String,
}

fn row_to_preset(row: &rusqlite::Row<'_>) -> rusqlite::Result<CommonPreset> {
    let tags_raw: String = row.get(6)?;
    let field_keys_raw: String = row.get(7)?;
    Ok(CommonPreset {
        id: row.get(0)?,
        scope: row.get(1)?,
        kind: row.get(2)?,
        category: row.get(3)?,
        name: row.get(4)?,
        content: row.get(5)?,
        tags: serde_json::from_str(&tags_raw).unwrap_or_default(),
        applicable_field_keys: serde_json::from_str(&field_keys_raw).unwrap_or_default(),
        usage_count: row.get(8)?,
        last_used_at: row.get(9)?,
        enabled: row.get::<_, i64>(10)? != 0,
        created_at: row.get(11)?,
        updated_at: row.get(12)?,
    })
}

fn get_preset_locked(conn: &Connection, id: &str) -> Result<CommonPreset, String> {
    conn.query_row(
        "SELECT id, scope, kind, category, name, content, tags_json, applicable_field_keys_json,
            usage_count, last_used_at, enabled, created_at, updated_at
         FROM common_presets
         WHERE id = ?1 AND deleted_at IS NULL",
        [id],
        row_to_preset,
    )
    .optional()
    .map_err(|e| e.to_string())?
    .ok_or_else(|| format!("CommonPresetNotFound::{}", id))
}

fn list_presets_locked(
    conn: &Connection,
    filter: CommonPresetFilter,
) -> Result<Vec<CommonPreset>, String> {
    ensure_schema(conn).map_err(|e| e.to_string())?;
    if let Some(kind) = filter.kind.as_deref() {
        validate_kind(kind)?;
    }

    let order_by = match filter.sort_by.as_deref() {
        Some("usage") => {
            "usage_count DESC, last_used_at IS NULL ASC, last_used_at DESC, updated_at DESC"
        }
        _ => "last_used_at IS NULL ASC, last_used_at DESC, usage_count DESC, updated_at DESC",
    };
    let sql = format!(
        "SELECT id, scope, kind, category, name, content, tags_json, applicable_field_keys_json,
            usage_count, last_used_at, enabled, created_at, updated_at
         FROM common_presets
         WHERE deleted_at IS NULL
            AND scope = 'workspace'
            AND (?1 IS NULL OR kind = ?1)
            AND (?2 IS NULL OR category = ?2)
            AND (?3 = 1 OR enabled = 1)
         ORDER BY {}",
        order_by
    );
    let mut stmt = conn.prepare(&sql).map_err(|e| e.to_string())?;
    let rows = stmt
        .query_map(
            params![
                filter.kind,
                filter.category.filter(|value| !value.trim().is_empty()),
                if filter.include_disabled.unwrap_or(false) {
                    1
                } else {
                    0
                },
            ],
            row_to_preset,
        )
        .map_err(|e| e.to_string())?;

    let field_key = filter.field_key.map(|value| value.trim().to_string());
    let mut presets = Vec::new();
    for row in rows {
        let preset = row.map_err(|e| e.to_string())?;
        if let Some(field_key) = field_key.as_deref() {
            if field_key.is_empty() {
                presets.push(preset);
            } else if preset.applicable_field_keys.is_empty()
                || preset
                    .applicable_field_keys
                    .iter()
                    .any(|key| key == field_key)
            {
                presets.push(preset);
            }
        } else {
            presets.push(preset);
        }
    }
    Ok(presets)
}

fn save_preset_locked(conn: &Connection, input: CommonPresetInput) -> Result<CommonPreset, String> {
    ensure_schema(conn).map_err(|e| e.to_string())?;
    validate_kind(&input.kind)?;
    let scope = input.scope.unwrap_or_else(|| "workspace".to_string());
    validate_scope(&scope)?;

    let category = input.category.trim();
    let name = input.name.trim();
    let content = input.content.trim();
    if category.is_empty() {
        return Err("CommonPresetCategoryRequired".to_string());
    }
    if name.is_empty() {
        return Err("CommonPresetNameRequired".to_string());
    }
    if content.is_empty() {
        return Err("CommonPresetContentRequired".to_string());
    }

    let tags = normalize_list(input.tags.unwrap_or_default());
    let applicable_field_keys = normalize_list(input.applicable_field_keys.unwrap_or_default());
    validate_preset_field_keys(&applicable_field_keys)?;
    let now = now_iso();
    let enabled = input.enabled.unwrap_or(true);

    let id = if let Some(id) = input.id {
        let existing_id: Option<String> = conn
            .query_row(
                "SELECT id FROM common_presets WHERE id = ?1 AND deleted_at IS NULL",
                [&id],
                |row| row.get(0),
            )
            .optional()
            .map_err(|e| e.to_string())?;
        if existing_id.is_none() {
            return Err(format!("CommonPresetNotFound::{}", id));
        }
        conn.execute(
            "UPDATE common_presets
             SET scope = ?1, kind = ?2, category = ?3, name = ?4, content = ?5,
                tags_json = ?6, applicable_field_keys_json = ?7, enabled = ?8, updated_at = ?9
             WHERE id = ?10 AND deleted_at IS NULL",
            params![
                scope,
                input.kind,
                category,
                name,
                content,
                serde_json::to_string(&tags).map_err(|e| e.to_string())?,
                serde_json::to_string(&applicable_field_keys).map_err(|e| e.to_string())?,
                if enabled { 1 } else { 0 },
                now,
                id,
            ],
        )
        .map_err(|e| e.to_string())?;
        id
    } else {
        let id = generate_id();
        conn.execute(
            "INSERT INTO common_presets (
                id, scope, kind, category, name, content, tags_json, applicable_field_keys_json,
                usage_count, enabled, created_at, updated_at
             ) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, 0, ?9, ?10, ?11)",
            params![
                id,
                scope,
                input.kind,
                category,
                name,
                content,
                serde_json::to_string(&tags).map_err(|e| e.to_string())?,
                serde_json::to_string(&applicable_field_keys).map_err(|e| e.to_string())?,
                if enabled { 1 } else { 0 },
                now,
                now,
            ],
        )
        .map_err(|e| e.to_string())?;
        id
    };

    get_preset_locked(conn, &id)
}

fn list_field_settings_locked(conn: &Connection) -> Result<Vec<PresetFieldSetting>, String> {
    ensure_schema(conn).map_err(|e| e.to_string())?;
    let mut stmt = conn
        .prepare(
            "SELECT field_key, enabled, updated_at
             FROM preset_field_settings
             ORDER BY field_key ASC",
        )
        .map_err(|e| e.to_string())?;
    let rows = stmt
        .query_map([], |row| {
            Ok(PresetFieldSetting {
                field_key: row.get(0)?,
                enabled: row.get::<_, i64>(1)? != 0,
                updated_at: row.get(2)?,
            })
        })
        .map_err(|e| e.to_string())?;

    let mut settings = Vec::new();
    for row in rows {
        settings.push(row.map_err(|e| e.to_string())?);
    }
    Ok(settings)
}

fn set_field_setting_locked(
    conn: &Connection,
    field_key: &str,
    enabled: bool,
) -> Result<PresetFieldSetting, String> {
    ensure_schema(conn).map_err(|e| e.to_string())?;
    let field_key = field_key.trim();
    validate_preset_field_key(field_key)?;
    let now = now_iso();
    conn.execute(
        "INSERT INTO preset_field_settings (field_key, enabled, updated_at)
         VALUES (?1, ?2, ?3)
         ON CONFLICT(field_key) DO UPDATE SET
            enabled = excluded.enabled,
            updated_at = excluded.updated_at",
        params![field_key, if enabled { 1 } else { 0 }, now],
    )
    .map_err(|e| e.to_string())?;
    Ok(PresetFieldSetting {
        field_key: field_key.to_string(),
        enabled,
        updated_at: now,
    })
}

#[tauri::command]
pub async fn list_common_presets(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    filter: Option<CommonPresetFilter>,
) -> Result<Vec<CommonPreset>, String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    list_presets_locked(
        &conn,
        filter.unwrap_or(CommonPresetFilter {
            kind: None,
            category: None,
            field_key: None,
            include_disabled: None,
            sort_by: None,
        }),
    )
}

#[tauri::command]
pub async fn list_preset_field_settings(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
) -> Result<Vec<PresetFieldSetting>, String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    list_field_settings_locked(&conn)
}

#[tauri::command]
pub async fn set_preset_field_enabled(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    field_key: String,
    enabled: bool,
) -> Result<PresetFieldSetting, String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    set_field_setting_locked(&conn, &field_key, enabled)
}

#[tauri::command]
pub async fn save_common_preset(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    preset: CommonPresetInput,
) -> Result<CommonPreset, String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    save_preset_locked(&conn, preset)
}

#[tauri::command]
pub async fn set_common_preset_enabled(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    id: String,
    enabled: bool,
) -> Result<CommonPreset, String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    ensure_schema(&conn).map_err(|e| e.to_string())?;
    let now = now_iso();
    let changed = conn
        .execute(
            "UPDATE common_presets SET enabled = ?1, updated_at = ?2 WHERE id = ?3 AND deleted_at IS NULL",
            params![if enabled { 1 } else { 0 }, now, id],
        )
        .map_err(|e| e.to_string())?;
    if changed == 0 {
        return Err(format!("CommonPresetNotFound::{}", id));
    }
    get_preset_locked(&conn, &id)
}

#[tauri::command]
pub async fn delete_common_preset(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    id: String,
) -> Result<(), String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    ensure_schema(&conn).map_err(|e| e.to_string())?;
    let now = now_iso();
    let changed = conn
        .execute(
            "UPDATE common_presets
             SET deleted_at = ?1, enabled = 0, updated_at = ?1
             WHERE id = ?2 AND deleted_at IS NULL",
            params![now, id],
        )
        .map_err(|e| e.to_string())?;
    if changed == 0 {
        return Err(format!("CommonPresetNotFound::{}", id));
    }
    Ok(())
}

#[tauri::command]
pub async fn mark_common_preset_used(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    id: String,
) -> Result<CommonPreset, String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    ensure_schema(&conn).map_err(|e| e.to_string())?;
    let now = now_iso();
    let changed = conn
        .execute(
            "UPDATE common_presets
             SET usage_count = usage_count + 1, last_used_at = ?1, updated_at = ?1
             WHERE id = ?2 AND deleted_at IS NULL",
            params![now, id],
        )
        .map_err(|e| e.to_string())?;
    if changed == 0 {
        return Err(format!("CommonPresetNotFound::{}", id));
    }
    get_preset_locked(&conn, &id)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn common_preset_crud_and_field_filter_work() {
        let conn = Connection::open_in_memory().unwrap();
        ensure_schema(&conn).unwrap();

        let created = save_preset_locked(
            &conn,
            CommonPresetInput {
                id: None,
                scope: None,
                kind: "short_value".to_string(),
                category: "审核人员".to_string(),
                name: "默认审核人员".to_string(),
                content: "张三、李四".to_string(),
                tags: Some(vec!["常用".to_string(), "常用".to_string()]),
                applicable_field_keys: Some(vec!["approval.reviewers".to_string()]),
                enabled: Some(true),
            },
        )
        .unwrap();

        assert_eq!(created.scope, "workspace");
        assert_eq!(created.tags, vec!["常用".to_string()]);

        let by_field = list_presets_locked(
            &conn,
            CommonPresetFilter {
                kind: Some("short_value".to_string()),
                category: None,
                field_key: Some("approval.reviewers".to_string()),
                include_disabled: None,
                sort_by: None,
            },
        )
        .unwrap();
        assert_eq!(by_field.len(), 1);

        let unrelated = list_presets_locked(
            &conn,
            CommonPresetFilter {
                kind: Some("short_value".to_string()),
                category: None,
                field_key: Some("project_basic.customer_name".to_string()),
                include_disabled: None,
                sort_by: None,
            },
        )
        .unwrap();
        assert!(unrelated.is_empty());

        let used = {
            let now = now_iso();
            conn.execute(
                "UPDATE common_presets
                 SET usage_count = usage_count + 1, last_used_at = ?1, updated_at = ?1
                 WHERE id = ?2",
                params![now, &created.id],
            )
            .unwrap();
            get_preset_locked(&conn, &created.id).unwrap()
        };
        assert_eq!(used.usage_count, 1);
        assert!(used.last_used_at.is_some());

        conn.execute(
            "UPDATE common_presets SET enabled = 0 WHERE id = ?1",
            [&created.id],
        )
        .unwrap();
        let enabled_only = list_presets_locked(
            &conn,
            CommonPresetFilter {
                kind: None,
                category: None,
                field_key: None,
                include_disabled: None,
                sort_by: None,
            },
        )
        .unwrap();
        assert!(enabled_only.is_empty());
    }

    #[test]
    fn field_settings_persist_and_financial_fields_are_rejected() {
        let conn = Connection::open_in_memory().unwrap();
        ensure_schema(&conn).unwrap();

        let preset = save_preset_locked(
            &conn,
            CommonPresetInput {
                id: None,
                scope: None,
                kind: "short_value".to_string(),
                category: "产权归属".to_string(),
                name: "客户所有".to_string(),
                content: "客户所有".to_string(),
                tags: None,
                applicable_field_keys: Some(vec!["project_basic.property_rights".to_string()]),
                enabled: Some(true),
            },
        )
        .unwrap();

        let enabled =
            set_field_setting_locked(&conn, "project_basic.property_rights", true).unwrap();
        assert!(enabled.enabled);
        assert_eq!(list_field_settings_locked(&conn).unwrap().len(), 1);

        let disabled =
            set_field_setting_locked(&conn, "project_basic.property_rights", false).unwrap();
        assert!(!disabled.enabled);
        let persisted = list_field_settings_locked(&conn).unwrap();
        assert_eq!(persisted.len(), 1);
        assert!(!persisted[0].enabled);
        assert_eq!(
            get_preset_locked(&conn, &preset.id).unwrap().content,
            "客户所有"
        );

        let rejected = set_field_setting_locked(&conn, "finance.revenue_amount", true)
            .expect_err("financial amount fields must not be preset-enabled");
        assert!(rejected.starts_with("PresetFieldNotEligible::"));

        let rejected_preset = save_preset_locked(
            &conn,
            CommonPresetInput {
                id: None,
                scope: None,
                kind: "short_value".to_string(),
                category: "财务字段".to_string(),
                name: "收入金额".to_string(),
                content: "100000".to_string(),
                tags: None,
                applicable_field_keys: Some(vec!["finance.revenue_amount".to_string()]),
                enabled: Some(true),
            },
        )
        .expect_err("financial amount fields must not accept preset bindings");
        assert!(rejected_preset.starts_with("PresetFieldNotEligible::"));
    }
}
