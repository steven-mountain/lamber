use rusqlite::{params, OptionalExtension};
use serde::{Deserialize, Serialize};
use serde_json::Value;
use std::fs;
use std::path::Path;
use std::sync::Arc;
use tauri::{AppHandle, State};
use tauri_plugin_dialog::DialogExt;

const H200_BASELINE_ROLE: &str = "h200_baseline";
const H200_BASELINE_DEFAULT_DESCRIPTION: &str = "智算项目默认金额来源";
const H200_BASELINE_PRESET_DESCRIPTION: &str =
    "64 台 H200、5 年服务期的标准报价预设。金额口径为元、含税。";
const AMOUNT_SOURCE_PACKAGE_KIND: &str = "lamber.intelligentCompute.amountSource";
const AMOUNT_SOURCE_PACKAGE_SCHEMA_VERSION: i64 = 1;

fn now_iso() -> String {
    chrono::Utc::now().to_rfc3339()
}

fn parse_json(raw: String, fallback: Value) -> Value {
    serde_json::from_str(&raw).unwrap_or(fallback)
}

fn validate_amount_source_package(value: &Value) -> Result<(), String> {
    let object = value
        .as_object()
        .ok_or_else(|| "InvalidAmountSourcePackage::not_object".to_string())?;
    if object.get("kind").and_then(Value::as_str) != Some(AMOUNT_SOURCE_PACKAGE_KIND) {
        return Err("InvalidAmountSourcePackage::kind".to_string());
    }
    if object.get("schemaVersion").and_then(Value::as_i64)
        != Some(AMOUNT_SOURCE_PACKAGE_SCHEMA_VERSION)
    {
        return Err("InvalidAmountSourcePackage::schemaVersion".to_string());
    }
    let source = object
        .get("source")
        .and_then(Value::as_object)
        .ok_or_else(|| "InvalidAmountSourcePackage::source".to_string())?;
    if source
        .get("name")
        .and_then(Value::as_str)
        .map(str::trim)
        .filter(|name| !name.is_empty())
        .is_none()
    {
        return Err("InvalidAmountSourcePackage::source.name".to_string());
    }
    for key in [
        "parameterGroups",
        "parameters",
        "revenueItems",
        "costItems",
        "mappings",
    ] {
        if !source.get(key).is_some_and(Value::is_array) {
            return Err(format!("InvalidAmountSourcePackage::source.{}", key));
        }
    }
    if !source.get("metadata").is_some_and(Value::is_object) {
        return Err("InvalidAmountSourcePackage::source.metadata".to_string());
    }
    if !source
        .get("calculationSnapshot")
        .is_some_and(Value::is_object)
    {
        return Err("InvalidAmountSourcePackage::source.calculationSnapshot".to_string());
    }
    let project_settings = object
        .get("projectSettings")
        .and_then(Value::as_object)
        .ok_or_else(|| "InvalidAmountSourcePackage::projectSettings".to_string())?;
    let project_years = project_settings
        .get("projectYears")
        .and_then(Value::as_i64)
        .ok_or_else(|| "InvalidAmountSourcePackage::projectSettings.projectYears".to_string())?;
    if !(1..=10).contains(&project_years) {
        return Err("InvalidAmountSourcePackage::projectSettings.projectYears".to_string());
    }
    let discount_rate = project_settings
        .get("discountRate")
        .and_then(Value::as_f64)
        .ok_or_else(|| "InvalidAmountSourcePackage::projectSettings.discountRate".to_string())?;
    if !discount_rate.is_finite() || !(0.0..=1.0).contains(&discount_rate) {
        return Err("InvalidAmountSourcePackage::projectSettings.discountRate".to_string());
    }
    Ok(())
}

fn read_amount_source_package(path: &Path) -> Result<Value, String> {
    let raw =
        fs::read_to_string(path).map_err(|e| format!("ReadAmountSourcePackageFailed::{}", e))?;
    let value: Value =
        serde_json::from_str(&raw).map_err(|e| format!("InvalidAmountSourcePackageJson::{}", e))?;
    validate_amount_source_package(&value)?;
    Ok(value)
}

fn write_amount_source_package(path: &Path, value: &Value) -> Result<(), String> {
    validate_amount_source_package(value)?;
    let raw = serde_json::to_string_pretty(value).map_err(|e| e.to_string())?;
    fs::write(path, raw).map_err(|e| format!("WriteAmountSourcePackageFailed::{}", e))
}

fn ensure_intelligent_project(conn: &rusqlite::Connection, project_id: &str) -> Result<(), String> {
    let project_type: Option<String> = conn
        .query_row(
            "SELECT project_type FROM projects WHERE id = ?1",
            [project_id],
            |row| row.get(0),
        )
        .optional()
        .map_err(|e| e.to_string())?;
    match project_type.as_deref() {
        Some("intelligent_compute") => Ok(()),
        Some(_) => Err(format!("ProjectTypeMismatch::{}::ict", project_id)),
        None => Err(format!("ProjectNotFoundInCurrentWorkspace::{}", project_id)),
    }
}

#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct IntelligentComputeProjectState {
    pub project_id: String,
    pub state_version: i64,
    pub active_amount_source_id: Option<String>,
    pub project_years: i64,
    pub discount_rate: f64,
    pub sync_revision: i64,
    pub controlled_subjects: Value,
    pub last_result: Value,
    pub created_at: String,
    pub updated_at: String,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct IntelligentAmountSource {
    pub id: String,
    pub project_id: String,
    pub name: String,
    pub description: Option<String>,
    pub enabled: bool,
    pub source_version: i64,
    pub metadata: Value,
    pub parameter_groups: Value,
    pub parameters: Value,
    pub revenue_items: Value,
    pub cost_items: Value,
    pub mappings: Value,
    pub calculation_snapshot: Value,
    pub created_at: String,
    pub updated_at: String,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct IntelligentComputeProjectData {
    pub state: IntelligentComputeProjectState,
    pub amount_sources: Vec<IntelligentAmountSource>,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SaveIntelligentProjectStateRequest {
    pub expected_version: i64,
    pub active_amount_source_id: Option<String>,
    pub project_years: i64,
    pub discount_rate: f64,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SaveIntelligentAmountSourceRequest {
    pub source: IntelligentAmountSource,
    pub expected_version: i64,
}

fn row_to_state(row: &rusqlite::Row<'_>) -> rusqlite::Result<IntelligentComputeProjectState> {
    Ok(IntelligentComputeProjectState {
        project_id: row.get(0)?,
        state_version: row.get(1)?,
        active_amount_source_id: row.get(2)?,
        project_years: row.get(3)?,
        discount_rate: row.get(4)?,
        sync_revision: row.get(5)?,
        controlled_subjects: parse_json(row.get(6)?, serde_json::json!({})),
        last_result: parse_json(row.get(7)?, serde_json::json!({})),
        created_at: row.get(8)?,
        updated_at: row.get(9)?,
    })
}

fn row_to_source(row: &rusqlite::Row<'_>) -> rusqlite::Result<IntelligentAmountSource> {
    Ok(IntelligentAmountSource {
        id: row.get(0)?,
        project_id: row.get(1)?,
        name: row.get(2)?,
        description: row.get(3)?,
        enabled: row.get::<_, i64>(4)? != 0,
        source_version: row.get(5)?,
        metadata: parse_json(row.get(6)?, serde_json::json!({})),
        parameter_groups: parse_json(row.get(7)?, serde_json::json!([])),
        parameters: parse_json(row.get(8)?, serde_json::json!([])),
        revenue_items: parse_json(row.get(9)?, serde_json::json!([])),
        cost_items: parse_json(row.get(10)?, serde_json::json!([])),
        mappings: parse_json(row.get(11)?, serde_json::json!([])),
        calculation_snapshot: parse_json(row.get(12)?, serde_json::json!({})),
        created_at: row.get(13)?,
        updated_at: row.get(14)?,
    })
}

pub(crate) fn get_project_state(
    conn: &rusqlite::Connection,
    project_id: &str,
) -> Result<IntelligentComputeProjectState, String> {
    conn.query_row(
        "SELECT project_id, state_version, active_amount_source_id, project_years,
            discount_rate, sync_revision, controlled_subjects_json, last_result_json,
            created_at, updated_at
         FROM project_intelligent_compute_states WHERE project_id = ?1",
        [project_id],
        row_to_state,
    )
    .map_err(|e| e.to_string())
}

fn is_h200_baseline_source(
    conn: &rusqlite::Connection,
    project_id: &str,
    source: &IntelligentAmountSource,
) -> Result<bool, String> {
    if source.metadata.get("sourceRole").and_then(Value::as_str) == Some(H200_BASELINE_ROLE) {
        return Ok(true);
    }
    let has_legacy_baseline_description = source.description.as_deref()
        == Some(H200_BASELINE_DEFAULT_DESCRIPTION)
        || source.description.as_deref() == Some(H200_BASELINE_PRESET_DESCRIPTION);
    if !has_legacy_baseline_description {
        return Ok(false);
    }
    let first_source_id: Option<String> = conn
        .query_row(
            "SELECT id FROM intelligent_compute_amount_sources
             WHERE project_id = ?1 ORDER BY created_at ASC LIMIT 1",
            [project_id],
            |row| row.get(0),
        )
        .optional()
        .map_err(|e| e.to_string())?;
    Ok(first_source_id.as_deref() == Some(source.id.as_str()))
}

pub(crate) fn ensure_project_state(
    conn: &rusqlite::Connection,
    project_id: &str,
) -> Result<(), String> {
    let project: (i64, f64) = conn
        .query_row(
            "SELECT project_years, discount_rate FROM projects WHERE id = ?1",
            [project_id],
            |row| Ok((row.get(0)?, row.get(1)?)),
        )
        .map_err(|e| e.to_string())?;
    let now = now_iso();
    conn.execute(
        "INSERT OR IGNORE INTO project_intelligent_compute_states (
            project_id, state_version, active_amount_source_id, project_years, discount_rate,
            sync_revision, controlled_subjects_json, last_result_json, created_at, updated_at
         ) VALUES (?1, 1, NULL, ?2, ?3, 0, '{}', '{}', ?4, ?4)",
        params![project_id, project.0, project.1, now],
    )
    .map_err(|e| e.to_string())?;
    Ok(())
}

pub(crate) fn ensure_default_amount_source(
    conn: &rusqlite::Connection,
    project_id: &str,
) -> Result<String, String> {
    ensure_project_state(conn, project_id)?;
    let existing: Option<String> = conn
        .query_row(
            "SELECT id FROM intelligent_compute_amount_sources
             WHERE project_id = ?1 ORDER BY created_at ASC LIMIT 1",
            [project_id],
            |row| row.get(0),
        )
        .optional()
        .map_err(|e| e.to_string())?;
    if let Some(id) = existing {
        return Ok(id);
    }
    let source_id = format!("amount_source_{}", uuid::Uuid::new_v4().simple());
    let now = now_iso();
    let metadata = serde_json::json!({
        "sourceRole": H200_BASELINE_ROLE,
        "basePreset": "h200",
    });
    conn.execute(
        "INSERT INTO intelligent_compute_amount_sources (
            id, project_id, name, description, enabled, source_version, metadata_json,
            parameter_groups_json, parameters_json, revenue_items_json, cost_items_json,
            mappings_json, calculation_snapshot_json, created_at, updated_at
         ) VALUES (?1, ?2, 'H200 标准智算金额来源', '智算项目默认金额来源', 1, 1,
            ?3, '[]', '[]', '[]', '[]', '[]', '{}', ?4, ?4)",
        params![source_id, project_id, metadata.to_string(), now],
    )
    .map_err(|e| e.to_string())?;
    conn.execute(
        "UPDATE project_intelligent_compute_states
         SET active_amount_source_id = ?1, updated_at = ?2
         WHERE project_id = ?3",
        params![source_id, now, project_id],
    )
    .map_err(|e| e.to_string())?;
    Ok(source_id)
}

#[tauri::command]
pub async fn get_intelligent_compute_project(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
) -> Result<IntelligentComputeProjectData, String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    ensure_intelligent_project(&conn, &project_id)?;
    ensure_project_state(&conn, &project_id)?;
    ensure_default_amount_source(&conn, &project_id)?;
    let mut stmt = conn
        .prepare(
            "SELECT id, project_id, name, description, enabled, source_version, metadata_json,
                parameter_groups_json, parameters_json, revenue_items_json, cost_items_json,
                mappings_json, calculation_snapshot_json, created_at, updated_at
             FROM intelligent_compute_amount_sources
             WHERE project_id = ?1 ORDER BY created_at ASC",
        )
        .map_err(|e| e.to_string())?;
    let rows = stmt
        .query_map([&project_id], row_to_source)
        .map_err(|e| e.to_string())?;
    let mut amount_sources = Vec::new();
    for row in rows {
        amount_sources.push(row.map_err(|e| e.to_string())?);
    }
    Ok(IntelligentComputeProjectData {
        state: get_project_state(&conn, &project_id)?,
        amount_sources,
    })
}

#[tauri::command]
pub async fn save_intelligent_compute_project_state(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
    request: SaveIntelligentProjectStateRequest,
) -> Result<IntelligentComputeProjectState, String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let mut conn = db.lock().map_err(|e| e.to_string())?;
    ensure_intelligent_project(&conn, &project_id)?;
    ensure_project_state(&conn, &project_id)?;
    if let Some(source_id) = request.active_amount_source_id.as_deref() {
        let owned: bool = conn
            .query_row(
                "SELECT EXISTS(
                    SELECT 1 FROM intelligent_compute_amount_sources
                    WHERE id = ?1 AND project_id = ?2
                 )",
                params![source_id, project_id],
                |row| row.get(0),
            )
            .map_err(|e| e.to_string())?;
        if !owned {
            return Err("AmountSourceNotFound".to_string());
        }
    }
    let now = now_iso();
    let tx = conn.transaction().map_err(|e| e.to_string())?;
    let affected = tx
        .execute(
            "UPDATE project_intelligent_compute_states
             SET state_version = state_version + 1,
                 active_amount_source_id = ?1,
                 project_years = ?2,
                 discount_rate = ?3,
                 updated_at = ?4
             WHERE project_id = ?5 AND state_version = ?6",
            params![
                request.active_amount_source_id,
                request.project_years.clamp(1, 10),
                request.discount_rate.clamp(0.0, 1.0),
                now,
                project_id,
                request.expected_version,
            ],
        )
        .map_err(|e| e.to_string())?;
    if affected == 0 {
        return Err("IntelligentComputeStateVersionConflict".to_string());
    }
    tx.execute(
        "UPDATE projects SET project_years = ?1, discount_rate = ?2, updated_at = ?3 WHERE id = ?4",
        params![
            request.project_years.clamp(1, 10),
            request.discount_rate.clamp(0.0, 1.0),
            now,
            project_id,
        ],
    )
    .map_err(|e| e.to_string())?;
    if let Some(source_id) = request.active_amount_source_id.as_deref() {
        tx.execute(
            "UPDATE intelligent_compute_amount_sources
             SET enabled = CASE WHEN id = ?1 THEN 1 ELSE 0 END,
                 updated_at = ?2
             WHERE project_id = ?3",
            params![source_id, now, project_id],
        )
        .map_err(|e| e.to_string())?;
    }
    tx.commit().map_err(|e| e.to_string())?;
    conn.query_row(
        "SELECT project_id, state_version, active_amount_source_id, project_years,
            discount_rate, sync_revision, controlled_subjects_json, last_result_json,
            created_at, updated_at
         FROM project_intelligent_compute_states WHERE project_id = ?1",
        [&project_id],
        row_to_state,
    )
    .map_err(|e| e.to_string())
}

#[tauri::command]
pub async fn save_intelligent_amount_source(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
    request: SaveIntelligentAmountSourceRequest,
) -> Result<IntelligentAmountSource, String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let mut conn = db.lock().map_err(|e| e.to_string())?;
    ensure_intelligent_project(&conn, &project_id)?;
    if request.source.project_id != project_id {
        return Err("AmountSourceProjectMismatch".to_string());
    }
    let now = now_iso();
    let tx = conn.transaction().map_err(|e| e.to_string())?;
    let exists: Option<i64> = tx
        .query_row(
            "SELECT source_version FROM intelligent_compute_amount_sources
             WHERE id = ?1 AND project_id = ?2",
            params![request.source.id, project_id],
            |row| row.get(0),
        )
        .optional()
        .map_err(|e| e.to_string())?;
    if request.source.enabled {
        tx.execute(
            "UPDATE intelligent_compute_amount_sources
             SET enabled = 0, updated_at = ?1
             WHERE project_id = ?2 AND id <> ?3",
            params![now, project_id, request.source.id],
        )
        .map_err(|e| e.to_string())?;
    }
    if let Some(current_version) = exists {
        if current_version != request.expected_version {
            return Err(format!(
                "IntelligentAmountSourceVersionConflict::expected={}::current={}",
                request.expected_version, current_version
            ));
        }
        tx.execute(
            "UPDATE intelligent_compute_amount_sources
             SET name = ?1, description = ?2, enabled = ?3,
                 source_version = source_version + 1, metadata_json = ?4,
                 parameter_groups_json = ?5, parameters_json = ?6,
                 revenue_items_json = ?7, cost_items_json = ?8, mappings_json = ?9,
                 calculation_snapshot_json = ?10, updated_at = ?11
             WHERE id = ?12 AND project_id = ?13",
            params![
                request.source.name.trim(),
                request.source.description,
                request.source.enabled as i64,
                request.source.metadata.to_string(),
                request.source.parameter_groups.to_string(),
                request.source.parameters.to_string(),
                request.source.revenue_items.to_string(),
                request.source.cost_items.to_string(),
                request.source.mappings.to_string(),
                request.source.calculation_snapshot.to_string(),
                now,
                request.source.id,
                project_id,
            ],
        )
        .map_err(|e| e.to_string())?;
    } else {
        if request.expected_version != 0 {
            return Err("IntelligentAmountSourceVersionConflict::new".to_string());
        }
        tx.execute(
            "INSERT INTO intelligent_compute_amount_sources (
                id, project_id, name, description, enabled, source_version, metadata_json,
                parameter_groups_json, parameters_json, revenue_items_json, cost_items_json,
                mappings_json, calculation_snapshot_json, created_at, updated_at
             ) VALUES (?1, ?2, ?3, ?4, ?5, 1, ?6, ?7, ?8, ?9, ?10, ?11, ?12, ?13, ?13)",
            params![
                request.source.id,
                project_id,
                request.source.name.trim(),
                request.source.description,
                request.source.enabled as i64,
                request.source.metadata.to_string(),
                request.source.parameter_groups.to_string(),
                request.source.parameters.to_string(),
                request.source.revenue_items.to_string(),
                request.source.cost_items.to_string(),
                request.source.mappings.to_string(),
                request.source.calculation_snapshot.to_string(),
                now,
            ],
        )
        .map_err(|e| e.to_string())?;
    }
    tx.commit().map_err(|e| e.to_string())?;
    conn.query_row(
        "SELECT id, project_id, name, description, enabled, source_version, metadata_json,
            parameter_groups_json, parameters_json, revenue_items_json, cost_items_json,
            mappings_json, calculation_snapshot_json, created_at, updated_at
         FROM intelligent_compute_amount_sources WHERE id = ?1 AND project_id = ?2",
        params![request.source.id, project_id],
        row_to_source,
    )
    .map_err(|e| e.to_string())
}

#[tauri::command]
pub async fn delete_intelligent_amount_source(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
    amount_source_id: String,
) -> Result<(), String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let mut conn = db.lock().map_err(|e| e.to_string())?;
    ensure_intelligent_project(&conn, &project_id)?;
    delete_intelligent_amount_source_locked(&mut conn, &project_id, &amount_source_id)
}

#[tauri::command]
pub async fn export_intelligent_amount_source_package(
    app: AppHandle,
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
    package_payload: Value,
    default_file_name: String,
) -> Result<Option<String>, String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    ensure_intelligent_project(&conn, &project_id)?;
    validate_amount_source_package(&package_payload)?;
    drop(conn);

    let file = app
        .dialog()
        .file()
        .set_title("导出智算金额来源")
        .set_file_name(default_file_name)
        .add_filter("智算金额来源 JSON", &["json"])
        .blocking_save_file();
    let Some(path) = file else {
        return Ok(None);
    };
    let path_string = path.to_string();
    write_amount_source_package(Path::new(&path_string), &package_payload)?;
    Ok(Some(path_string))
}

#[tauri::command]
pub async fn select_and_read_intelligent_amount_source_package(
    app: AppHandle,
) -> Result<Option<Value>, String> {
    let file = app
        .dialog()
        .file()
        .set_title("选择智算金额来源 JSON")
        .add_filter("智算金额来源 JSON", &["json"])
        .blocking_pick_file();
    let Some(path) = file else {
        return Ok(None);
    };
    let path_string = path.to_string();
    read_amount_source_package(Path::new(&path_string)).map(Some)
}

pub(crate) fn delete_intelligent_amount_source_locked(
    conn: &mut rusqlite::Connection,
    project_id: &str,
    amount_source_id: &str,
) -> Result<(), String> {
    let count: i64 = conn
        .query_row(
            "SELECT COUNT(*) FROM intelligent_compute_amount_sources WHERE project_id = ?1",
            [project_id],
            |row| row.get(0),
        )
        .map_err(|e| e.to_string())?;
    if count <= 1 {
        return Err("CannotDeleteLastAmountSource".to_string());
    }
    let target = conn
        .query_row(
            "SELECT id, project_id, name, description, enabled, source_version, metadata_json,
                parameter_groups_json, parameters_json, revenue_items_json, cost_items_json,
                mappings_json, calculation_snapshot_json, created_at, updated_at
             FROM intelligent_compute_amount_sources WHERE id = ?1 AND project_id = ?2",
            params![amount_source_id, project_id],
            row_to_source,
        )
        .optional()
        .map_err(|e| e.to_string())?
        .ok_or_else(|| "AmountSourceNotFound".to_string())?;
    if is_h200_baseline_source(conn, project_id, &target)? {
        return Err("CannotDeleteH200BaselineAmountSource".to_string());
    }
    let tx = conn.transaction().map_err(|e| e.to_string())?;
    let affected = tx
        .execute(
            "DELETE FROM intelligent_compute_amount_sources WHERE id = ?1 AND project_id = ?2",
            params![amount_source_id, project_id],
        )
        .map_err(|e| e.to_string())?;
    if affected == 0 {
        return Err("AmountSourceNotFound".to_string());
    }
    let next_id: Option<String> = tx
        .query_row(
            "SELECT id FROM intelligent_compute_amount_sources
             WHERE project_id = ?1 ORDER BY created_at ASC LIMIT 1",
            [project_id],
            |row| row.get(0),
        )
        .optional()
        .map_err(|e| e.to_string())?;
    tx.execute(
        "UPDATE project_intelligent_compute_states
         SET active_amount_source_id = CASE
             WHEN active_amount_source_id = ?1 THEN ?2
             ELSE active_amount_source_id
         END,
         state_version = state_version + 1,
         updated_at = ?3
         WHERE project_id = ?4",
        params![amount_source_id, next_id, now_iso(), project_id],
    )
    .map_err(|e| e.to_string())?;
    tx.commit().map_err(|e| e.to_string())?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sample_amount_source_package() -> Value {
        serde_json::json!({
            "kind": AMOUNT_SOURCE_PACKAGE_KIND,
            "schemaVersion": AMOUNT_SOURCE_PACKAGE_SCHEMA_VERSION,
            "exportedAt": "2026-06-16T00:00:00Z",
            "projectSettings": {
                "projectYears": 5,
                "discountRate": 0.05
            },
            "source": {
                "name": "测试金额来源",
                "description": "测试",
                "metadata": {},
                "parameterGroups": [],
                "parameters": [],
                "revenueItems": [],
                "costItems": [],
                "mappings": [],
                "calculationSnapshot": {}
            }
        })
    }

    #[test]
    fn amount_source_package_validation_rejects_invalid_kind_and_schema() {
        let mut invalid_kind = sample_amount_source_package();
        invalid_kind["kind"] = serde_json::json!("other");
        assert_eq!(
            validate_amount_source_package(&invalid_kind).unwrap_err(),
            "InvalidAmountSourcePackage::kind"
        );

        let mut invalid_schema = sample_amount_source_package();
        invalid_schema["schemaVersion"] = serde_json::json!(99);
        assert_eq!(
            validate_amount_source_package(&invalid_schema).unwrap_err(),
            "InvalidAmountSourcePackage::schemaVersion"
        );
    }

    #[test]
    fn amount_source_package_validation_rejects_missing_core_arrays() {
        let mut package = sample_amount_source_package();
        package["source"]["revenueItems"] = Value::Null;
        assert_eq!(
            validate_amount_source_package(&package).unwrap_err(),
            "InvalidAmountSourcePackage::source.revenueItems"
        );
    }

    #[test]
    fn amount_source_package_can_be_written_and_read_back() {
        let package = sample_amount_source_package();
        let path = std::env::temp_dir().join(format!(
            "lamber-amount-source-package-{}.json",
            uuid::Uuid::new_v4().simple()
        ));
        write_amount_source_package(&path, &package).expect("write package");
        let read_back = read_amount_source_package(&path).expect("read package");
        let _ = fs::remove_file(&path);
        assert_eq!(read_back, package);
    }

    fn create_amount_source_delete_test_db() -> rusqlite::Connection {
        let conn = rusqlite::Connection::open_in_memory().expect("open in-memory database");
        conn.execute_batch(
            "
            CREATE TABLE projects (
                id TEXT PRIMARY KEY,
                project_type TEXT NOT NULL,
                project_years INTEGER NOT NULL,
                discount_rate REAL NOT NULL
            );
            CREATE TABLE project_intelligent_compute_states (
                project_id TEXT PRIMARY KEY,
                state_version INTEGER NOT NULL DEFAULT 1,
                active_amount_source_id TEXT,
                project_years INTEGER NOT NULL DEFAULT 1,
                discount_rate REAL NOT NULL DEFAULT 0.055,
                sync_revision INTEGER NOT NULL DEFAULT 0,
                controlled_subjects_json TEXT NOT NULL DEFAULT '{}',
                last_result_json TEXT NOT NULL DEFAULT '{}',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );
            CREATE TABLE intelligent_compute_amount_sources (
                id TEXT PRIMARY KEY,
                project_id TEXT NOT NULL,
                name TEXT NOT NULL,
                description TEXT,
                enabled INTEGER NOT NULL DEFAULT 1,
                source_version INTEGER NOT NULL DEFAULT 1,
                metadata_json TEXT NOT NULL DEFAULT '{}',
                parameter_groups_json TEXT NOT NULL DEFAULT '[]',
                parameters_json TEXT NOT NULL DEFAULT '[]',
                revenue_items_json TEXT NOT NULL DEFAULT '[]',
                cost_items_json TEXT NOT NULL DEFAULT '[]',
                mappings_json TEXT NOT NULL DEFAULT '[]',
                calculation_snapshot_json TEXT NOT NULL DEFAULT '{}',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );
            INSERT INTO projects (id, project_type, project_years, discount_rate)
            VALUES ('project-1', 'intelligent_compute', 5, 0.05);
            INSERT INTO project_intelligent_compute_states (
                project_id, state_version, active_amount_source_id, project_years, discount_rate,
                sync_revision, controlled_subjects_json, last_result_json, created_at, updated_at
            ) VALUES ('project-1', 1, 'source-h200', 5, 0.05, 0, '{}', '{}', 'before', 'before');
            ",
        )
        .expect("create amount source delete schema");
        conn
    }

    fn insert_amount_source(
        conn: &rusqlite::Connection,
        id: &str,
        description: &str,
        metadata: Value,
        created_at: &str,
    ) {
        conn.execute(
            "INSERT INTO intelligent_compute_amount_sources (
                id, project_id, name, description, enabled, source_version, metadata_json,
                parameter_groups_json, parameters_json, revenue_items_json, cost_items_json,
                mappings_json, calculation_snapshot_json, created_at, updated_at
             ) VALUES (?1, 'project-1', ?2, ?3, 1, 1, ?4, '[]', '[]', '[]', '[]', '[]', '{}', ?5, ?5)",
            params![id, id, description, metadata.to_string(), created_at],
        )
        .expect("insert amount source");
    }

    #[test]
    fn delete_amount_source_rejects_last_source() {
        let mut conn = create_amount_source_delete_test_db();
        insert_amount_source(&conn, "source-only", "普通来源", serde_json::json!({}), "1");
        conn.execute(
            "UPDATE project_intelligent_compute_states SET active_amount_source_id = 'source-only'
             WHERE project_id = 'project-1'",
            [],
        )
        .unwrap();

        let error = delete_intelligent_amount_source_locked(&mut conn, "project-1", "source-only")
            .unwrap_err();
        assert_eq!(error, "CannotDeleteLastAmountSource");
    }

    #[test]
    fn delete_amount_source_rejects_h200_baseline() {
        let mut conn = create_amount_source_delete_test_db();
        insert_amount_source(
            &conn,
            "source-h200",
            H200_BASELINE_DEFAULT_DESCRIPTION,
            serde_json::json!({"sourceRole": H200_BASELINE_ROLE}),
            "1",
        );
        insert_amount_source(
            &conn,
            "source-quote",
            H200_BASELINE_PRESET_DESCRIPTION,
            serde_json::json!({}),
            "2",
        );

        let error = delete_intelligent_amount_source_locked(&mut conn, "project-1", "source-h200")
            .unwrap_err();
        assert_eq!(error, "CannotDeleteH200BaselineAmountSource");
        let count: i64 = conn
            .query_row(
                "SELECT COUNT(*) FROM intelligent_compute_amount_sources WHERE project_id = 'project-1'",
                [],
                |row| row.get(0),
            )
            .unwrap();
        assert_eq!(count, 2);
    }

    #[test]
    fn delete_amount_source_removes_normal_source_and_switches_active_source() {
        let mut conn = create_amount_source_delete_test_db();
        insert_amount_source(
            &conn,
            "source-h200",
            H200_BASELINE_DEFAULT_DESCRIPTION,
            serde_json::json!({"sourceRole": H200_BASELINE_ROLE}),
            "1",
        );
        insert_amount_source(
            &conn,
            "source-quote",
            H200_BASELINE_PRESET_DESCRIPTION,
            serde_json::json!({}),
            "2",
        );
        conn.execute(
            "UPDATE project_intelligent_compute_states SET active_amount_source_id = 'source-quote'
             WHERE project_id = 'project-1'",
            [],
        )
        .unwrap();

        delete_intelligent_amount_source_locked(&mut conn, "project-1", "source-quote")
            .expect("delete normal amount source");
        let (active_source, state_version, count): (String, i64, i64) = conn
            .query_row(
                "SELECT s.active_amount_source_id, s.state_version,
                    (SELECT COUNT(*) FROM intelligent_compute_amount_sources WHERE project_id = 'project-1')
                 FROM project_intelligent_compute_states s WHERE s.project_id = 'project-1'",
                [],
                |row| Ok((row.get(0)?, row.get(1)?, row.get(2)?)),
            )
            .unwrap();
        assert_eq!(active_source, "source-h200");
        assert_eq!(state_version, 2);
        assert_eq!(count, 1);
    }
}
