use crate::benefit::models::StoreData;
use chrono::Utc;
use std::fs;
use std::path::Path;
use std::sync::Arc;
use serde::Serialize;

#[derive(Serialize, Clone, Debug)]
#[serde(rename_all = "camelCase")]
pub struct MigrationReport {
    pub success: bool,
    pub backup_path: String,
    pub projects_count: usize,
    pub schemes_count: usize,
    pub snapshots_count: usize,
    pub files_count: usize,
    pub message: String,
}

pub fn check_migration_needed(conn: &rusqlite::Connection, json_path: &Path) -> bool {
    if !json_path.exists() {
        return false;
    }
    if let Ok(metadata) = fs::metadata(json_path) {
        if metadata.len() == 0 {
            return false;
        }
    } else {
        return false;
    }

    let mut stmt = match conn.prepare("SELECT value FROM app_settings WHERE key = 'migration_status'") {
        Ok(s) => s,
        Err(_) => return false,
    };

    let mut rows = match stmt.query([]) {
        Ok(r) => r,
        Err(_) => return false,
    };

    if let Ok(Some(row)) = rows.next() {
        let status: String = row.get(0).unwrap_or_default();
        return status != "completed" && status != "skipped";
    }

    true
}

pub fn run_migration(conn: &mut rusqlite::Connection, json_path: &Path) -> Result<MigrationReport, String> {
    if !json_path.exists() {
        return Err("旧数据文件 projects_store.json 不存在".to_string());
    }

    // 1. Create a backup file
    let timestamp = Utc::now().format("%Y%m%d_%H%M%S").to_string();
    let mut backup_path = json_path.to_path_buf();
    if let Some(parent) = json_path.parent() {
        backup_path = parent.join(format!("projects_store_backup_{}.json", timestamp));
    } else {
        backup_path.set_extension(format!("backup_{}.json", timestamp));
    }

    fs::copy(json_path, &backup_path)
        .map_err(|e| format!("备份旧 JSON 文件失败: {}", e))?;

    let backup_path_str = backup_path.to_string_lossy().to_string();

    // 2. Read and parse the JSON file
    let content = fs::read_to_string(json_path)
        .map_err(|e| format!("读取旧 JSON 文件失败: {}", e))?;
    
    let store: StoreData = serde_json::from_str(&content)
        .map_err(|e| format!("解析旧数据失败: {}", e))?;

    let projects_count = store.projects.len();
    let schemes_count = store.schemes.len();
    let snapshots_count = store.snapshots.len();
    let files_count = store.project_files.len();

    // 3. Start a database transaction
    let tx = conn.transaction().map_err(|e| format!("开启数据库事务失败: {}", e))?;

    // 4. Insert projects
    for project in &store.projects {
        let summary_metrics_str = project.summary_metrics.as_ref()
            .and_then(|m| serde_json::to_string(m).ok());
        let logs_str = serde_json::to_string(&project.logs).unwrap_or_default();

        tx.execute(
            "INSERT OR REPLACE INTO projects (
                id, name, customer_name, status, benefit_status, default_scheme_id, created_at, updated_at,
                total_revenue_incl, total_cost_incl, project_years, discount_rate, cashflow_model,
                summary_metrics, folder_path, main_document_path, main_budget_file_path, note, logs
            ) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10, ?11, ?12, ?13, ?14, ?15, ?16, ?17, ?18, ?19)",
            rusqlite::params![
                project.id,
                project.name,
                project.customer_name,
                project.status,
                project.benefit_status,
                project.default_scheme_id,
                project.created_at,
                project.updated_at,
                project.total_revenue_incl,
                project.total_cost_incl,
                project.project_years,
                project.discount_rate,
                project.cashflow_model,
                summary_metrics_str,
                project.folder_path,
                project.main_document_path,
                project.main_budget_file_path,
                project.note,
                logs_str,
            ],
        ).map_err(|e| format!("迁移项目失败 (ID: {}): {}", project.id, e))?;
    }

    // 5. Insert schemes
    for scheme in &store.schemes {
        tx.execute(
            "INSERT OR REPLACE INTO benefit_schemes (id, project_id, name, created_at, updated_at) VALUES (?1, ?2, ?3, ?4, ?5)",
            rusqlite::params![
                scheme.id,
                scheme.project_id,
                scheme.name,
                scheme.created_at,
                scheme.updated_at,
            ],
        ).map_err(|e| format!("迁移分析方案失败 (ID: {}): {}", scheme.id, e))?;
    }

    // 6. Insert snapshots
    for snapshot in &store.snapshots {
        let input_params_str = serde_json::to_string(&snapshot.input_params)
            .map_err(|e| format!("序列化测算输入参数失败: {}", e))?;
        let output_metrics_str = serde_json::to_string(&snapshot.output_metrics)
            .map_err(|e| format!("序列化测算输出参数失败: {}", e))?;

        tx.execute(
            "INSERT OR REPLACE INTO benefit_snapshots (id, scheme_id, project_id, version, input_params, output_metrics, fingerprint, created_at) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8)",
            rusqlite::params![
                snapshot.id,
                snapshot.scheme_id,
                snapshot.project_id,
                snapshot.version,
                input_params_str,
                output_metrics_str,
                snapshot.fingerprint,
                snapshot.created_at,
            ],
        ).map_err(|e| format!("迁移测算快照失败 (ID: {}): {}", snapshot.id, e))?;
    }

    // 7. Insert project files
    for file in &store.project_files {
        tx.execute(
            "INSERT OR REPLACE INTO project_files (
                id, project_id, file_name, file_path, original_path, managed_path, file_type, extension,
                size, exists, last_scanned_at, modified_at, storage_mode, is_main_document, is_main_budget_file,
                note, created_at, updated_at
            ) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10, ?11, ?12, ?13, ?14, ?15, ?16, ?17, ?18)",
            rusqlite::params![
                file.id,
                file.project_id,
                file.file_name,
                file.file_path,
                file.original_path,
                file.managed_path,
                file.file_type,
                file.extension,
                file.size,
                if file.exists { 1 } else { 0 },
                file.last_scanned_at,
                file.modified_at,
                file.storage_mode,
                if file.is_main_document { 1 } else { 0 },
                if file.is_main_budget_file { 1 } else { 0 },
                file.note,
                file.created_at,
                file.updated_at,
            ],
        ).map_err(|e| format!("迁移文件记录失败 (ID: {}): {}", file.id, e))?;
    }

    // 8. Write migration_status = "completed"
    let now = Utc::now().to_rfc3339();
    tx.execute(
        "INSERT OR REPLACE INTO app_settings (key, value, updated_at) VALUES ('migration_status', 'completed', ?1)",
        [now],
    ).map_err(|e| format!("更新迁移状态失败: {}", e))?;

    // 9. Commit transaction
    tx.commit().map_err(|e| format!("提交迁移事务失败: {}", e))?;

    Ok(MigrationReport {
        success: true,
        backup_path: backup_path_str,
        projects_count,
        schemes_count,
        snapshots_count,
        files_count,
        message: "数据迁移成功完成。".to_string(),
    })
}

pub fn skip_migration(conn: &rusqlite::Connection) -> Result<(), String> {
    let now = Utc::now().to_rfc3339();
    conn.execute(
        "INSERT OR REPLACE INTO app_settings (key, value, updated_at) VALUES ('migration_status', 'skipped', ?1)",
        [now],
    ).map_err(|e| format!("更新迁移状态为跳过失败: {}", e))?;
    Ok(())
}

use tauri::{AppHandle, State, Manager};
use std::sync::Mutex;

#[tauri::command]
pub async fn check_db_migration(
    app: AppHandle,
    conn: State<'_, Arc<Mutex<rusqlite::Connection>>>,
) -> Result<bool, String> {
    let app_data_dir = app
        .path()
        .app_data_dir()
        .map_err(|e| format!("无法获取 App 数据目录: {}", e))?;
    let store_path = app_data_dir.join("projects_store.json");
    let db_conn = conn.lock().map_err(|e| e.to_string())?;
    Ok(check_migration_needed(&db_conn, &store_path))
}

#[tauri::command]
pub async fn run_db_migration(
    app: AppHandle,
    conn: State<'_, Arc<Mutex<rusqlite::Connection>>>,
    project_repo: State<'_, Arc<crate::benefit::repository::DualProjectRepository>>,
    file_repo: State<'_, Arc<crate::project_files::repository::DualProjectFileRepository>>,
) -> Result<MigrationReport, String> {
    let app_data_dir = app
        .path()
        .app_data_dir()
        .map_err(|e| format!("无法获取 App 数据目录: {}", e))?;
    let store_path = app_data_dir.join("projects_store.json");
    
    let mut db_conn = conn.lock().map_err(|e| e.to_string())?;
    let report = run_migration(&mut db_conn, &store_path)?;

    // Switch to SQLite backend
    let sqlite_p = crate::benefit::repository::SqliteProjectRepository::new(conn.inner().clone());
    let sqlite_f = crate::project_files::repository::SqliteProjectFileRepository::new(conn.inner().clone());
    
    project_repo.switch_to_sqlite(sqlite_p);
    file_repo.switch_to_sqlite(sqlite_f);

    Ok(report)
}

#[tauri::command]
pub async fn skip_db_migration(
    conn: State<'_, Arc<Mutex<rusqlite::Connection>>>,
    project_repo: State<'_, Arc<crate::benefit::repository::DualProjectRepository>>,
    file_repo: State<'_, Arc<crate::project_files::repository::DualProjectFileRepository>>,
) -> Result<(), String> {
    let db_conn = conn.lock().map_err(|e| e.to_string())?;
    skip_migration(&db_conn)?;

    // Switch to SQLite backend
    let sqlite_p = crate::benefit::repository::SqliteProjectRepository::new(conn.inner().clone());
    let sqlite_f = crate::project_files::repository::SqliteProjectFileRepository::new(conn.inner().clone());
    
    project_repo.switch_to_sqlite(sqlite_p);
    file_repo.switch_to_sqlite(sqlite_f);

    Ok(())
}
