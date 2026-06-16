use rusqlite::{params, OptionalExtension};
use serde::{Deserialize, Serialize};
use serde_json::Value;
use std::sync::Arc;
use tauri::State;

fn now_iso() -> String {
    chrono::Utc::now().to_rfc3339()
}

fn parse_json(raw: String, fallback: Value) -> Value {
    serde_json::from_str(&raw).unwrap_or(fallback)
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
    conn.execute(
        "INSERT INTO intelligent_compute_amount_sources (
            id, project_id, name, description, enabled, source_version, metadata_json,
            parameter_groups_json, parameters_json, revenue_items_json, cost_items_json,
            mappings_json, calculation_snapshot_json, created_at, updated_at
         ) VALUES (?1, ?2, 'H200 标准智算金额来源', '智算项目默认金额来源', 1, 1,
            '{}', '[]', '[]', '[]', '[]', '[]', '{}', ?3, ?3)",
        params![source_id, project_id, now],
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
    let state = conn
        .query_row(
            "SELECT project_id, state_version, active_amount_source_id, project_years,
                discount_rate, sync_revision, controlled_subjects_json, last_result_json,
                created_at, updated_at
             FROM project_intelligent_compute_states WHERE project_id = ?1",
            [&project_id],
            row_to_state,
        )
        .map_err(|e| e.to_string())?;
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
        state,
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
    let count: i64 = conn
        .query_row(
            "SELECT COUNT(*) FROM intelligent_compute_amount_sources WHERE project_id = ?1",
            [&project_id],
            |row| row.get(0),
        )
        .map_err(|e| e.to_string())?;
    if count <= 1 {
        return Err("CannotDeleteLastAmountSource".to_string());
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
            [&project_id],
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
