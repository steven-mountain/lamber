use crate::benefit::models::{
    BenefitAnalysisScheme, BenefitAnalysisSnapshot, IctInput, IctResult, Project, ProjectLog,
    SummaryMetrics,
};
use crate::benefit::service::compute_fingerprint;
use rusqlite::{params, OptionalExtension};
use serde::{Deserialize, Serialize};
use serde_json::Value;
use std::sync::Arc;
use std::time::{SystemTime, UNIX_EPOCH};
use tauri::State;

fn generate_id(prefix: &str) -> String {
    let now = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_else(|_| std::time::Duration::from_secs(0));
    format!("{}_{}_{}", prefix, now.as_millis(), now.subsec_nanos())
}

fn now_iso() -> String {
    chrono::Utc::now().to_rfc3339()
}

fn json_string(value: &Value) -> Result<String, String> {
    serde_json::to_string(value).map_err(|e| e.to_string())
}

fn ensure_project_exists(conn: &rusqlite::Connection, project_id: &str) -> Result<(), String> {
    let exists: bool = conn
        .query_row(
            "SELECT EXISTS(SELECT 1 FROM projects WHERE id = ?1)",
            [project_id],
            |row| row.get(0),
        )
        .map_err(|e| e.to_string())?;
    if exists {
        Ok(())
    } else {
        Err(format!("ProjectNotFoundInCurrentWorkspace::{}", project_id))
    }
}

fn row_to_project(row: &rusqlite::Row<'_>) -> rusqlite::Result<Project> {
    let summary_metrics: Option<String> = row.get(13)?;
    let summary_metrics = summary_metrics.and_then(|s| serde_json::from_str(&s).ok());
    let logs_str: Option<String> = row.get(18).ok();
    let logs: Vec<ProjectLog> = logs_str
        .and_then(|s| serde_json::from_str(&s).ok())
        .unwrap_or_default();

    Ok(Project {
        id: row.get(0)?,
        name: row.get(1)?,
        customer_name: row.get(2)?,
        status: row.get(3)?,
        benefit_status: row.get(4)?,
        default_scheme_id: row.get(5)?,
        created_at: row.get(6)?,
        updated_at: row.get(7)?,
        total_revenue_incl: row.get(8)?,
        total_cost_incl: row.get(9)?,
        project_years: row.get(10)?,
        discount_rate: row.get(11)?,
        cashflow_model: row.get(12)?,
        summary_metrics,
        folder_path: row.get(14)?,
        main_document_path: row.get(15)?,
        main_budget_file_path: row.get(16)?,
        note: row.get(17)?,
        logs,
        folder_name: row.get(19)?,
        relative_path: row.get(20)?,
        progress: row.get(21).unwrap_or(0.0),
        deadline: row.get(22)?,
        linked_folder_type: row.get(23)?,
        linked_folder_relative_path: row.get(24)?,
        linked_folder_external_path: row.get(25)?,
    })
}

fn get_project_locked(
    conn: &rusqlite::Connection,
    project_id: &str,
) -> Result<Option<Project>, String> {
    conn.query_row(
        "SELECT id, name, customer_name, status, benefit_status, default_scheme_id, created_at, updated_at,
            total_revenue_incl, total_cost_incl, project_years, discount_rate, cashflow_model, summary_metrics,
            folder_path, main_document_path, main_budget_file_path, note, logs, folder_name, relative_path,
            progress, deadline, linked_folder_type, linked_folder_relative_path, linked_folder_external_path
         FROM projects WHERE id = ?1",
        [project_id],
        row_to_project,
    )
    .optional()
    .map_err(|e| e.to_string())
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProjectDetailPatch {
    pub name: Option<String>,
    pub customer_name: Option<String>,
    pub status: Option<String>,
    pub progress: Option<f64>,
    pub deadline: Option<String>,
    pub note: Option<String>,
    pub linked_folder_type: Option<String>,
    pub linked_folder_relative_path: Option<String>,
    pub linked_folder_external_path: Option<String>,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct LifecycleStatePayload {
    pub profile_json: Value,
    pub parameters_json: Value,
    pub background_json: Value,
    pub input_payload_json: Value,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct CashflowStatePayload {
    pub cashflow_model: Option<String>,
    pub payment_model_json: Value,
    pub yearly_cashflow_json: Value,
    pub sector_cashflow_json: Value,
    pub assumptions_json: Value,
    pub metrics_json: Value,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct TemplateStatePayload {
    pub template_name: Option<String>,
    pub template_type: Option<String>,
    pub template_path: Option<String>,
    pub template_path_type: Option<String>,
    pub filled_data_json: Value,
    pub field_mapping_json: Value,
    pub output_config_json: Value,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct StoredLifecycleState {
    pub id: String,
    pub project_id: String,
    pub lifecycle_version: i64,
    pub profile_json: Value,
    pub parameters_json: Value,
    pub background_json: Value,
    pub input_payload_json: Value,
    pub created_at: String,
    pub updated_at: String,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct StoredCashflowState {
    pub id: String,
    pub project_id: String,
    pub cashflow_version: i64,
    pub cashflow_model: Option<String>,
    pub payment_model_json: Value,
    pub yearly_cashflow_json: Value,
    pub sector_cashflow_json: Value,
    pub assumptions_json: Value,
    pub metrics_json: Value,
    pub created_at: String,
    pub updated_at: String,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct StoredTemplateState {
    pub id: String,
    pub project_id: String,
    pub template_id: String,
    pub template_name: Option<String>,
    pub template_type: Option<String>,
    pub template_version: i64,
    pub template_path: Option<String>,
    pub template_path_type: Option<String>,
    pub filled_data_json: Value,
    pub field_mapping_json: Value,
    pub output_config_json: Value,
    pub created_at: String,
    pub updated_at: String,
    pub source: String,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TemplateAssetInfo {
    pub id: String,
    pub project_id: String,
    pub template_name: String,
    pub template_id: Option<String>,
    pub asset_type: String,
    pub usage: Option<String>,
    pub original_file_name: Option<String>,
    pub relative_path: String,
    pub mime_type: Option<String>,
    pub file_size: i64,
    pub width: Option<i32>,
    pub height: Option<i32>,
    pub file_hash: Option<String>,
    pub created_at: String,
    pub updated_at: String,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ProjectFullState {
    pub project: Project,
    pub lifecycle_state: Option<StoredLifecycleState>,
    pub cashflow_state: Option<StoredCashflowState>,
    pub schemes: Vec<BenefitAnalysisScheme>,
    pub latest_snapshot: Option<BenefitAnalysisSnapshot>,
    pub template_states: Vec<StoredTemplateState>,
    pub template_assets: Vec<TemplateAssetInfo>,
    pub legacy_lifecycle_input: Option<IctInput>,
    pub legacy_cashflow_metrics: Option<IctResult>,
}

fn parse_json_column(raw: String) -> Value {
    serde_json::from_str(&raw).unwrap_or(Value::Object(Default::default()))
}

fn get_lifecycle_state_locked(
    conn: &rusqlite::Connection,
    project_id: &str,
) -> Result<Option<StoredLifecycleState>, String> {
    conn.query_row(
        "SELECT id, project_id, lifecycle_version, profile_json, parameters_json, background_json, input_payload_json, created_at, updated_at
         FROM project_lifecycle_states WHERE project_id = ?1",
        [project_id],
        |row| {
            Ok(StoredLifecycleState {
                id: row.get(0)?,
                project_id: row.get(1)?,
                lifecycle_version: row.get(2)?,
                profile_json: parse_json_column(row.get(3)?),
                parameters_json: parse_json_column(row.get(4)?),
                background_json: parse_json_column(row.get(5)?),
                input_payload_json: parse_json_column(row.get(6)?),
                created_at: row.get(7)?,
                updated_at: row.get(8)?,
            })
        },
    )
    .optional()
    .map_err(|e| e.to_string())
}

fn get_cashflow_state_locked(
    conn: &rusqlite::Connection,
    project_id: &str,
) -> Result<Option<StoredCashflowState>, String> {
    conn.query_row(
        "SELECT id, project_id, cashflow_version, cashflow_model, payment_model_json, yearly_cashflow_json,
            sector_cashflow_json, assumptions_json, metrics_json, created_at, updated_at
         FROM project_cashflow_states WHERE project_id = ?1",
        [project_id],
        |row| {
            Ok(StoredCashflowState {
                id: row.get(0)?,
                project_id: row.get(1)?,
                cashflow_version: row.get(2)?,
                cashflow_model: row.get(3)?,
                payment_model_json: parse_json_column(row.get(4)?),
                yearly_cashflow_json: parse_json_column(row.get(5)?),
                sector_cashflow_json: parse_json_column(row.get(6)?),
                assumptions_json: parse_json_column(row.get(7)?),
                metrics_json: parse_json_column(row.get(8)?),
                created_at: row.get(9)?,
                updated_at: row.get(10)?,
            })
        },
    )
    .optional()
    .map_err(|e| e.to_string())
}

fn get_template_state_locked(
    conn: &rusqlite::Connection,
    project_id: &str,
    template_id: &str,
) -> Result<Option<StoredTemplateState>, String> {
    let state = conn
        .query_row(
            "SELECT id, project_id, template_id, template_name, template_type, template_version, template_path,
                template_path_type, filled_data_json, field_mapping_json, output_config_json, created_at, updated_at
             FROM project_template_states WHERE project_id = ?1 AND template_id = ?2",
            params![project_id, template_id],
            |row| {
                Ok(StoredTemplateState {
                    id: row.get(0)?,
                    project_id: row.get(1)?,
                    template_id: row.get(2)?,
                    template_name: row.get(3)?,
                    template_type: row.get(4)?,
                    template_version: row.get(5)?,
                    template_path: row.get(6)?,
                    template_path_type: row.get(7)?,
                    filled_data_json: parse_json_column(row.get(8)?),
                    field_mapping_json: parse_json_column(row.get(9)?),
                    output_config_json: parse_json_column(row.get(10)?),
                    created_at: row.get(11)?,
                    updated_at: row.get(12)?,
                    source: "project_template_states".to_string(),
                })
            },
        )
        .optional()
        .map_err(|e| e.to_string())?;

    if state.is_some() {
        return Ok(state);
    }

    let legacy_key = format!("template_form_data::{}", template_id);
    let legacy_value: Option<String> = conn
        .query_row(
            "SELECT value FROM project_settings WHERE project_id = ?1 AND key = ?2",
            params![project_id, legacy_key],
            |row| row.get(0),
        )
        .optional()
        .map_err(|e| e.to_string())?;

    Ok(legacy_value.map(|value| {
        let now = now_iso();
        StoredTemplateState {
            id: String::new(),
            project_id: project_id.to_string(),
            template_id: template_id.to_string(),
            template_name: Some(template_id.to_string()),
            template_type: None,
            template_version: 1,
            template_path: None,
            template_path_type: None,
            filled_data_json: parse_json_column(value),
            field_mapping_json: Value::Object(Default::default()),
            output_config_json: Value::Object(Default::default()),
            created_at: now.clone(),
            updated_at: now,
            source: "project_settings".to_string(),
        }
    }))
}

fn list_template_states_locked(
    conn: &rusqlite::Connection,
    project_id: &str,
) -> Result<Vec<StoredTemplateState>, String> {
    let mut stmt = conn
        .prepare(
            "SELECT id, project_id, template_id, template_name, template_type, template_version, template_path,
                template_path_type, filled_data_json, field_mapping_json, output_config_json, created_at, updated_at
             FROM project_template_states WHERE project_id = ?1 ORDER BY updated_at DESC",
        )
        .map_err(|e| e.to_string())?;
    let rows = stmt
        .query_map([project_id], |row| {
            Ok(StoredTemplateState {
                id: row.get(0)?,
                project_id: row.get(1)?,
                template_id: row.get(2)?,
                template_name: row.get(3)?,
                template_type: row.get(4)?,
                template_version: row.get(5)?,
                template_path: row.get(6)?,
                template_path_type: row.get(7)?,
                filled_data_json: parse_json_column(row.get(8)?),
                field_mapping_json: parse_json_column(row.get(9)?),
                output_config_json: parse_json_column(row.get(10)?),
                created_at: row.get(11)?,
                updated_at: row.get(12)?,
                source: "project_template_states".to_string(),
            })
        })
        .map_err(|e| e.to_string())?;

    let mut list = Vec::new();
    for row in rows {
        list.push(row.map_err(|e| e.to_string())?);
    }
    Ok(list)
}

fn list_template_assets_locked(
    conn: &rusqlite::Connection,
    project_id: &str,
    template_id: Option<&str>,
) -> Result<Vec<TemplateAssetInfo>, String> {
    let sql = if template_id.is_some() {
        "SELECT id, project_id, template_name, template_id, asset_type, usage, original_file_name, relative_path,
            mime_type, file_size, width, height, file_hash, created_at, updated_at
         FROM project_template_assets
         WHERE project_id = ?1 AND deleted_at IS NULL AND (template_id = ?2 OR template_name = ?2)
         ORDER BY updated_at DESC"
    } else {
        "SELECT id, project_id, template_name, template_id, asset_type, usage, original_file_name, relative_path,
            mime_type, file_size, width, height, file_hash, created_at, updated_at
         FROM project_template_assets
         WHERE project_id = ?1 AND deleted_at IS NULL
         ORDER BY updated_at DESC"
    };
    let mut stmt = conn.prepare(sql).map_err(|e| e.to_string())?;
    let mut list = Vec::new();
    if let Some(template_id) = template_id {
        let rows = stmt
            .query_map(params![project_id, template_id], |row| {
                Ok(TemplateAssetInfo {
                    id: row.get(0)?,
                    project_id: row.get(1)?,
                    template_name: row.get(2)?,
                    template_id: row.get(3)?,
                    asset_type: row.get(4)?,
                    usage: row.get(5)?,
                    original_file_name: row.get(6)?,
                    relative_path: row.get(7)?,
                    mime_type: row.get(8)?,
                    file_size: row.get(9)?,
                    width: row.get(10)?,
                    height: row.get(11)?,
                    file_hash: row.get(12)?,
                    created_at: row.get(13)?,
                    updated_at: row.get(14)?,
                })
            })
            .map_err(|e| e.to_string())?;
        for row in rows {
            list.push(row.map_err(|e| e.to_string())?);
        }
    } else {
        let rows = stmt
            .query_map([project_id], |row| {
                Ok(TemplateAssetInfo {
                    id: row.get(0)?,
                    project_id: row.get(1)?,
                    template_name: row.get(2)?,
                    template_id: row.get(3)?,
                    asset_type: row.get(4)?,
                    usage: row.get(5)?,
                    original_file_name: row.get(6)?,
                    relative_path: row.get(7)?,
                    mime_type: row.get(8)?,
                    file_size: row.get(9)?,
                    width: row.get(10)?,
                    height: row.get(11)?,
                    file_hash: row.get(12)?,
                    created_at: row.get(13)?,
                    updated_at: row.get(14)?,
                })
            })
            .map_err(|e| e.to_string())?;
        for row in rows {
            list.push(row.map_err(|e| e.to_string())?);
        }
    }
    Ok(list)
}

fn latest_snapshot_locked(
    conn: &rusqlite::Connection,
    project_id: &str,
) -> Result<Option<BenefitAnalysisSnapshot>, String> {
    conn.query_row(
        "SELECT id, scheme_id, project_id, version, input_params, output_metrics, fingerprint, created_at
         FROM benefit_snapshots WHERE project_id = ?1 ORDER BY created_at DESC, version DESC LIMIT 1",
        [project_id],
        |row| {
            let input_raw: String = row.get(4)?;
            let output_raw: String = row.get(5)?;
            Ok(BenefitAnalysisSnapshot {
                id: row.get(0)?,
                scheme_id: row.get(1)?,
                project_id: row.get(2)?,
                version: row.get(3)?,
                input_params: serde_json::from_str(&input_raw).map_err(|e| {
                    rusqlite::Error::FromSqlConversionFailure(4, rusqlite::types::Type::Text, Box::new(e))
                })?,
                output_metrics: serde_json::from_str(&output_raw).map_err(|e| {
                    rusqlite::Error::FromSqlConversionFailure(5, rusqlite::types::Type::Text, Box::new(e))
                })?,
                fingerprint: row.get(6)?,
                created_at: row.get(7)?,
            })
        },
    )
    .optional()
    .map_err(|e| e.to_string())
}

#[tauri::command]
pub async fn save_project_detail(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
    patch: ProjectDetailPatch,
) -> Result<Project, String> {
    let workspace = runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let mut conn = db.lock().map_err(|e| e.to_string())?;
    ensure_project_exists(&conn, &project_id)?;
    let mut project = get_project_locked(&conn, &project_id)?
        .ok_or_else(|| format!("ProjectNotFoundInCurrentWorkspace::{}", project_id))?;

    if let Some(name) = patch.name {
        let trimmed = name.trim();
        if trimmed.is_empty() {
            return Err("ProjectNameRequired".to_string());
        }
        let duplicate: bool = conn
            .query_row(
                "SELECT EXISTS(SELECT 1 FROM projects WHERE id <> ?1 AND lower(name) = lower(?2))",
                params![project_id, trimmed],
                |row| row.get(0),
            )
            .map_err(|e| e.to_string())?;
        if duplicate {
            return Err(format!("ProjectNameDuplicated::{}", trimmed));
        }
        project.name = trimmed.to_string();
    }
    if let Some(customer_name) = patch.customer_name {
        let trimmed = customer_name.trim();
        project.customer_name = if trimmed.is_empty() {
            "未知客户".to_string()
        } else {
            trimmed.to_string()
        };
    }
    if let Some(status) = patch.status {
        let trimmed = status.trim();
        if !trimmed.is_empty() {
            project.status = trimmed.to_string();
        }
    }
    if let Some(progress) = patch.progress {
        project.progress = progress.clamp(0.0, 100.0);
    }
    if patch.deadline.is_some() {
        project.deadline = patch.deadline;
    }
    if patch.note.is_some() {
        project.note = patch.note;
    }
    if patch.linked_folder_type.is_some() {
        project.linked_folder_type = patch.linked_folder_type;
    }
    if patch.linked_folder_relative_path.is_some() {
        project.linked_folder_relative_path = patch.linked_folder_relative_path;
    }
    if patch.linked_folder_external_path.is_some() {
        project.linked_folder_external_path = patch.linked_folder_external_path;
    }

    let now = now_iso();
    let logs_str = serde_json::to_string(&project.logs).map_err(|e| e.to_string())?;
    let affected = conn
        .execute(
            "UPDATE projects SET name = ?1, customer_name = ?2, status = ?3, progress = ?4, deadline = ?5,
                note = ?6, linked_folder_type = ?7, linked_folder_relative_path = ?8, linked_folder_external_path = ?9,
                logs = ?10, updated_at = ?11
             WHERE id = ?12",
            params![
                project.name,
                project.customer_name,
                project.status,
                project.progress,
                project.deadline,
                project.note,
                project.linked_folder_type,
                project.linked_folder_relative_path,
                project.linked_folder_external_path,
                logs_str,
                now,
                project_id,
            ],
        )
        .map_err(|e| e.to_string())?;
    if affected == 0 {
        return Err(format!("ProjectNotFoundInCurrentWorkspace::{}", project_id));
    }

    project.updated_at = now;
    if let Some(ref rel_path) = project.relative_path {
        let ws_root = std::path::Path::new(&workspace.workspace_root);
        let project_dir = crate::workspace::resolve_workspace_path(ws_root, rel_path);
        let project_json_path = project_dir.join("project.json");
        if project_json_path.exists() {
            if let Ok(content) = std::fs::read_to_string(&project_json_path) {
                if let Ok(mut json_val) = serde_json::from_str::<Value>(&content) {
                    json_val["name"] = Value::String(project.name.clone());
                    json_val["updatedAt"] = Value::String(project.updated_at.clone());
                    if let Ok(updated_content) = serde_json::to_string_pretty(&json_val) {
                        let _ = std::fs::write(project_json_path, updated_content);
                    }
                }
            }
        }
    }

    Ok(project)
}

#[tauri::command]
pub async fn get_project_detail(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
) -> Result<Project, String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    get_project_locked(&conn, &project_id)?
        .ok_or_else(|| format!("ProjectNotFoundInCurrentWorkspace::{}", project_id))
}

#[tauri::command]
pub async fn save_lifecycle_state(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
    lifecycle_state: LifecycleStatePayload,
) -> Result<StoredLifecycleState, String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    ensure_project_exists(&conn, &project_id)?;
    let now = now_iso();
    let existing_id: Option<String> = conn
        .query_row(
            "SELECT id FROM project_lifecycle_states WHERE project_id = ?1",
            [&project_id],
            |row| row.get(0),
        )
        .optional()
        .map_err(|e| e.to_string())?;
    let id = existing_id.unwrap_or_else(|| generate_id("lifecycle"));
    conn.execute(
        "INSERT INTO project_lifecycle_states (
            id, project_id, lifecycle_version, profile_json, parameters_json, background_json,
            input_payload_json, created_at, updated_at
         ) VALUES (?1, ?2, 1, ?3, ?4, ?5, ?6, ?7, ?8)
         ON CONFLICT(project_id) DO UPDATE SET
            lifecycle_version = lifecycle_version + 1,
            profile_json = excluded.profile_json,
            parameters_json = excluded.parameters_json,
            background_json = excluded.background_json,
            input_payload_json = excluded.input_payload_json,
            updated_at = excluded.updated_at",
        params![
            id,
            project_id,
            json_string(&lifecycle_state.profile_json)?,
            json_string(&lifecycle_state.parameters_json)?,
            json_string(&lifecycle_state.background_json)?,
            json_string(&lifecycle_state.input_payload_json)?,
            now,
            now,
        ],
    )
    .map_err(|e| e.to_string())?;
    get_lifecycle_state_locked(&conn, &project_id)?
        .ok_or_else(|| "LifecycleStateSaveFailed".to_string())
}

#[tauri::command]
pub async fn get_lifecycle_state(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
) -> Result<Option<StoredLifecycleState>, String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    ensure_project_exists(&conn, &project_id)?;
    get_lifecycle_state_locked(&conn, &project_id)
}

#[tauri::command]
pub async fn save_cashflow_state(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
    cashflow_state: CashflowStatePayload,
) -> Result<StoredCashflowState, String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    ensure_project_exists(&conn, &project_id)?;
    let now = now_iso();
    let existing_id: Option<String> = conn
        .query_row(
            "SELECT id FROM project_cashflow_states WHERE project_id = ?1",
            [&project_id],
            |row| row.get(0),
        )
        .optional()
        .map_err(|e| e.to_string())?;
    let id = existing_id.unwrap_or_else(|| generate_id("cashflow"));
    conn.execute(
        "INSERT INTO project_cashflow_states (
            id, project_id, cashflow_version, cashflow_model, payment_model_json, yearly_cashflow_json,
            sector_cashflow_json, assumptions_json, metrics_json, created_at, updated_at
         ) VALUES (?1, ?2, 1, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10)
         ON CONFLICT(project_id) DO UPDATE SET
            cashflow_version = cashflow_version + 1,
            cashflow_model = excluded.cashflow_model,
            payment_model_json = excluded.payment_model_json,
            yearly_cashflow_json = excluded.yearly_cashflow_json,
            sector_cashflow_json = excluded.sector_cashflow_json,
            assumptions_json = excluded.assumptions_json,
            metrics_json = excluded.metrics_json,
            updated_at = excluded.updated_at",
        params![
            id,
            project_id,
            cashflow_state.cashflow_model,
            json_string(&cashflow_state.payment_model_json)?,
            json_string(&cashflow_state.yearly_cashflow_json)?,
            json_string(&cashflow_state.sector_cashflow_json)?,
            json_string(&cashflow_state.assumptions_json)?,
            json_string(&cashflow_state.metrics_json)?,
            now,
            now,
        ],
    )
    .map_err(|e| e.to_string())?;
    get_cashflow_state_locked(&conn, &project_id)?
        .ok_or_else(|| "CashflowStateSaveFailed".to_string())
}

#[tauri::command]
pub async fn get_cashflow_state(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
) -> Result<Option<StoredCashflowState>, String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    ensure_project_exists(&conn, &project_id)?;
    get_cashflow_state_locked(&conn, &project_id)
}

#[tauri::command]
pub async fn save_template_state(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
    template_id: String,
    template_state: TemplateStatePayload,
) -> Result<StoredTemplateState, String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let mut conn = db.lock().map_err(|e| e.to_string())?;
    ensure_project_exists(&conn, &project_id)?;
    let tx = conn.transaction().map_err(|e| e.to_string())?;
    let now = now_iso();
    let existing_id: Option<String> = tx
        .query_row(
            "SELECT id FROM project_template_states WHERE project_id = ?1 AND template_id = ?2",
            params![project_id, template_id],
            |row| row.get(0),
        )
        .optional()
        .map_err(|e| e.to_string())?;
    let id = existing_id.unwrap_or_else(|| generate_id("template"));
    let template_name = template_state
        .template_name
        .clone()
        .unwrap_or_else(|| template_id.clone());
    tx.execute(
        "INSERT INTO project_template_states (
            id, project_id, template_id, template_name, template_type, template_version, template_path,
            template_path_type, filled_data_json, field_mapping_json, output_config_json, created_at, updated_at
         ) VALUES (?1, ?2, ?3, ?4, ?5, 1, ?6, ?7, ?8, ?9, ?10, ?11, ?12)
         ON CONFLICT(project_id, template_id) DO UPDATE SET
            template_name = excluded.template_name,
            template_type = excluded.template_type,
            template_version = template_version + 1,
            template_path = excluded.template_path,
            template_path_type = excluded.template_path_type,
            filled_data_json = excluded.filled_data_json,
            field_mapping_json = excluded.field_mapping_json,
            output_config_json = excluded.output_config_json,
            updated_at = excluded.updated_at",
        params![
            id,
            project_id,
            template_id,
            template_name,
            template_state.template_type,
            template_state.template_path,
            template_state.template_path_type,
            json_string(&template_state.filled_data_json)?,
            json_string(&template_state.field_mapping_json)?,
            json_string(&template_state.output_config_json)?,
            now,
            now,
        ],
    )
    .map_err(|e| e.to_string())?;

    // Keep the legacy key readable during the transition and for orphan-asset cleanup.
    tx.execute(
        "INSERT OR REPLACE INTO project_settings (project_id, key, value, updated_at) VALUES (?1, ?2, ?3, ?4)",
        params![
            project_id,
            format!("template_form_data::{}", template_id),
            json_string(&template_state.filled_data_json)?,
            now,
        ],
    )
    .map_err(|e| e.to_string())?;
    tx.commit().map_err(|e| e.to_string())?;

    get_template_state_locked(&conn, &project_id, &template_id)?
        .ok_or_else(|| "TemplateStateSaveFailed".to_string())
}

#[tauri::command]
pub async fn get_template_state(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
    template_id: String,
) -> Result<Option<StoredTemplateState>, String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    ensure_project_exists(&conn, &project_id)?;
    get_template_state_locked(&conn, &project_id, &template_id)
}

#[tauri::command]
pub async fn list_template_states(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
) -> Result<Vec<StoredTemplateState>, String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    ensure_project_exists(&conn, &project_id)?;
    list_template_states_locked(&conn, &project_id)
}

#[tauri::command]
pub async fn list_template_assets(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
    template_id: Option<String>,
) -> Result<Vec<TemplateAssetInfo>, String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    ensure_project_exists(&conn, &project_id)?;
    list_template_assets_locked(&conn, &project_id, template_id.as_deref())
}

#[tauri::command]
pub async fn save_benefit_analysis(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
    scheme_id_opt: Option<String>,
    scheme_name: String,
    input_params: IctInput,
    output_metrics: IctResult,
    is_save_as_new: bool,
) -> Result<Project, String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let mut conn = db.lock().map_err(|e| e.to_string())?;
    ensure_project_exists(&conn, &project_id)?;
    let tx = conn.transaction().map_err(|e| e.to_string())?;
    let mut project = get_project_locked(&tx, &project_id)?
        .ok_or_else(|| format!("ProjectNotFoundInCurrentWorkspace::{}", project_id))?;
    let timestamp = now_iso();

    let mut stmt = tx
        .prepare("SELECT id, project_id, name, created_at, updated_at FROM benefit_schemes WHERE project_id = ?1")
        .map_err(|e| e.to_string())?;
    let scheme_iter = stmt
        .query_map([&project_id], |row| {
            Ok(BenefitAnalysisScheme {
                id: row.get(0)?,
                project_id: row.get(1)?,
                name: row.get(2)?,
                created_at: row.get(3)?,
                updated_at: row.get(4)?,
            })
        })
        .map_err(|e| e.to_string())?;
    let mut existing_schemes = Vec::new();
    for scheme in scheme_iter {
        existing_schemes.push(scheme.map_err(|e| e.to_string())?);
    }
    drop(stmt);

    let (scheme_id, is_new_scheme) = if is_save_as_new {
        (generate_id("scheme"), true)
    } else if let Some(scheme_id) = scheme_id_opt {
        (scheme_id, false)
    } else if let Some(existing) = existing_schemes.iter().find(|s| s.name == scheme_name) {
        (existing.id.clone(), false)
    } else {
        (generate_id("scheme"), true)
    };

    if is_new_scheme || !existing_schemes.iter().any(|s| s.id == scheme_id) {
        tx.execute(
            "INSERT INTO benefit_schemes (id, project_id, name, created_at, updated_at) VALUES (?1, ?2, ?3, ?4, ?5)",
            params![scheme_id, project_id, scheme_name, timestamp, timestamp],
        )
        .map_err(|e| e.to_string())?;
    } else {
        tx.execute(
            "UPDATE benefit_schemes SET name = ?1, updated_at = ?2 WHERE id = ?3 AND project_id = ?4",
            params![scheme_name, timestamp, scheme_id, project_id],
        )
        .map_err(|e| e.to_string())?;
    }

    let version: i64 = tx
        .query_row(
            "SELECT COALESCE(MAX(version), 0) + 1 FROM benefit_snapshots WHERE scheme_id = ?1",
            [&scheme_id],
            |row| row.get(0),
        )
        .map_err(|e| e.to_string())?;
    let fingerprint = compute_fingerprint(&input_params);
    tx.execute(
        "INSERT INTO benefit_snapshots (id, scheme_id, project_id, version, input_params, output_metrics, fingerprint, created_at)
         VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8)",
        params![
            generate_id("snapshot"),
            scheme_id,
            project_id,
            version,
            serde_json::to_string(&input_params).map_err(|e| e.to_string())?,
            serde_json::to_string(&output_metrics).map_err(|e| e.to_string())?,
            fingerprint,
            timestamp,
        ],
    )
    .map_err(|e| e.to_string())?;

    let risk_level = {
        let npv_val = output_metrics.npv.parse::<f64>().unwrap_or(0.0);
        let margin_val = output_metrics
            .margin_rate
            .trim_end_matches('%')
            .parse::<f64>()
            .unwrap_or(0.0)
            / if output_metrics.margin_rate.contains('%') {
                100.0
            } else {
                1.0
            };
        let npv_rate_val = output_metrics
            .npv_rate
            .trim_end_matches('%')
            .parse::<f64>()
            .unwrap_or(0.0)
            / if output_metrics.npv_rate.contains('%') {
                100.0
            } else {
                1.0
            };
        if npv_val < 0.0 || margin_val < 0.0 {
            "高风险".to_string()
        } else if margin_val < 0.08 || npv_rate_val < 0.04 {
            "中风险".to_string()
        } else {
            "低风险".to_string()
        }
    };
    project.default_scheme_id = Some(scheme_id);
    project.benefit_status = "normal".to_string();
    project.summary_metrics = Some(SummaryMetrics {
        margin_rate: output_metrics.margin_rate.clone(),
        npv: output_metrics.npv.clone(),
        npv_rate: output_metrics.npv_rate.clone(),
        irr: output_metrics.irr.clone(),
        dynamic_payback: output_metrics.dynamic_payback.clone(),
        risk_level,
    });
    project.total_revenue_incl = [
        &input_params.rev_it_integration,
        &input_params.rev_it_maintenance,
        &input_params.rev_it_device_sales,
        &input_params.rev_it_device_lease,
        &input_params.rev_it_other,
        &input_params.rev_it_cloud,
        &input_params.rev_ct_line,
        &input_params.rev_ct_product,
        &input_params.rev_non_it_ct,
    ]
    .iter()
    .map(|item| item.incl_tax.parse::<f64>().unwrap_or(0.0))
    .sum();
    project.total_cost_incl = [
        &input_params.cost_it_device,
        &input_params.cost_it_construction,
        &input_params.cost_it_survey,
        &input_params.cost_it_integration,
        &input_params.cost_it_other,
        &input_params.cost_it_maintenance,
        &input_params.cost_it_running,
        &input_params.cost_it_bidding,
        &input_params.cost_it_design_eval,
        &input_params.cost_it_audit,
        &input_params.cost_ct_construction,
        &input_params.cost_ct_maintenance,
        &input_params.cost_ct_other,
        &input_params.cost_ct_bandwidth,
        &input_params.cost_ct_renewal,
        &input_params.cost_non_it_ct,
        &input_params.cost_mix_marketing,
        &input_params.cost_mix_channel,
        &input_params.cost_mix_other,
    ]
    .iter()
    .map(|item| item.incl_tax.parse::<f64>().unwrap_or(0.0))
    .sum();
    project.project_years = input_params.project_years.unwrap_or(1);
    project.discount_rate = input_params.discount_rate.parse::<f64>().unwrap_or(0.055);
    project.cashflow_model = input_params
        .cashflow_model
        .clone()
        .unwrap_or_else(|| "model_a".to_string());
    project.updated_at = timestamp.clone();
    project.logs.push(ProjectLog {
        id: generate_id("log"),
        timestamp: timestamp.clone(),
        description: "保存效益分析方案".to_string(),
    });

    tx.execute(
        "UPDATE projects SET default_scheme_id = ?1, benefit_status = ?2, updated_at = ?3,
            total_revenue_incl = ?4, total_cost_incl = ?5, project_years = ?6, discount_rate = ?7,
            cashflow_model = ?8, summary_metrics = ?9, logs = ?10
         WHERE id = ?11",
        params![
            project.default_scheme_id,
            project.benefit_status,
            project.updated_at,
            project.total_revenue_incl,
            project.total_cost_incl,
            project.project_years,
            project.discount_rate,
            project.cashflow_model,
            serde_json::to_string(&project.summary_metrics).map_err(|e| e.to_string())?,
            serde_json::to_string(&project.logs).map_err(|e| e.to_string())?,
            project_id,
        ],
    )
    .map_err(|e| e.to_string())?;
    tx.commit().map_err(|e| e.to_string())?;

    Ok(project)
}

#[tauri::command]
pub async fn get_benefit_schemes(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
) -> Result<Vec<BenefitAnalysisScheme>, String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    ensure_project_exists(&conn, &project_id)?;
    let mut stmt = conn
        .prepare("SELECT id, project_id, name, created_at, updated_at FROM benefit_schemes WHERE project_id = ?1 ORDER BY updated_at DESC")
        .map_err(|e| e.to_string())?;
    let rows = stmt
        .query_map([project_id], |row| {
            Ok(BenefitAnalysisScheme {
                id: row.get(0)?,
                project_id: row.get(1)?,
                name: row.get(2)?,
                created_at: row.get(3)?,
                updated_at: row.get(4)?,
            })
        })
        .map_err(|e| e.to_string())?;
    let mut list = Vec::new();
    for row in rows {
        list.push(row.map_err(|e| e.to_string())?);
    }
    Ok(list)
}

#[tauri::command]
pub async fn get_project_full_state(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
) -> Result<ProjectFullState, String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    let project = get_project_locked(&conn, &project_id)?
        .ok_or_else(|| format!("ProjectNotFoundInCurrentWorkspace::{}", project_id))?;
    let lifecycle_state = get_lifecycle_state_locked(&conn, &project_id)?;
    let cashflow_state = get_cashflow_state_locked(&conn, &project_id)?;
    let latest_snapshot = latest_snapshot_locked(&conn, &project_id)?;
    let schemes = {
        let mut stmt = conn
            .prepare("SELECT id, project_id, name, created_at, updated_at FROM benefit_schemes WHERE project_id = ?1 ORDER BY updated_at DESC")
            .map_err(|e| e.to_string())?;
        let rows = stmt
            .query_map([&project_id], |row| {
                Ok(BenefitAnalysisScheme {
                    id: row.get(0)?,
                    project_id: row.get(1)?,
                    name: row.get(2)?,
                    created_at: row.get(3)?,
                    updated_at: row.get(4)?,
                })
            })
            .map_err(|e| e.to_string())?;
        let mut list = Vec::new();
        for row in rows {
            list.push(row.map_err(|e| e.to_string())?);
        }
        list
    };
    Ok(ProjectFullState {
        project,
        lifecycle_state,
        cashflow_state,
        schemes,
        legacy_lifecycle_input: latest_snapshot
            .as_ref()
            .map(|snap| snap.input_params.clone()),
        legacy_cashflow_metrics: latest_snapshot
            .as_ref()
            .map(|snap| snap.output_metrics.clone()),
        latest_snapshot,
        template_states: list_template_states_locked(&conn, &project_id)?,
        template_assets: list_template_assets_locked(&conn, &project_id, None)?,
    })
}
