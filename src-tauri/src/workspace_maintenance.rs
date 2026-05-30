use crate::config_manager::{AppConfig, ConfigManager};
use crate::workspace::{self, CurrentWorkspace, WorkspaceManifest, WorkspaceRuntime};
use chrono::{Local, Utc};
use rusqlite::{params, Connection, OptionalExtension};
use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use std::collections::{BTreeMap, HashSet};
use std::fs::{self, File};
use std::io::{Read, Write};
use std::path::{Component, Path, PathBuf};
use std::sync::{Arc, Mutex};
use tauri::{AppHandle, State};
use zip::{write::SimpleFileOptions, ZipArchive, ZipWriter};

const BACKUPS_DIR: &str = ".backups";
const EXPORTS_DIR: &str = ".exports";
const PROJECT_ASSETS_DIR: &str = ".projects";
const RETIRED_MODULE_IDS: &[&str] = &["benefit_tool", "docfill_tool"];

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct HealthCheckResult {
    pub status: String,
    pub items: Vec<HealthCheckItem>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct HealthCheckItem {
    pub id: String,
    pub severity: String,
    pub category: String,
    pub message: String,
    pub detail: Option<String>,
    pub repairable: bool,
    pub repair_action: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WorkspaceBackupInfo {
    pub id: String,
    pub file_name: String,
    pub path: String,
    pub created_at: String,
    pub size_bytes: u64,
    pub is_daily: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WorkspaceExportOptions {
    pub include_backups: Option<bool>,
    pub include_exports: Option<bool>,
    pub allow_warnings: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ExportManifest {
    pub exported_at: String,
    pub app_version: String,
    pub workspace_id: String,
    pub workspace_name: String,
    pub workspace_version: i32,
    pub include_backups: bool,
    pub include_exports: bool,
    pub external_path_count: usize,
    pub warnings: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ExportWorkspaceResult {
    pub archive_path: String,
    pub database_backup_path: String,
    pub manifest: ExportManifest,
    pub warnings: Vec<HealthCheckItem>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ArchiveValidationResult {
    pub valid: bool,
    pub root_prefix: Option<String>,
    pub workspace_id: Option<String>,
    pub workspace_name: Option<String>,
    pub workspace_version: Option<i32>,
    pub errors: Vec<String>,
    pub warnings: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ImportWorkspaceOptions {
    pub open_after_import: Option<bool>,
    pub conflict_strategy: Option<String>,
    pub destination_name: Option<String>,
}

impl Default for ImportWorkspaceOptions {
    fn default() -> Self {
        Self {
            open_after_import: Some(false),
            conflict_strategy: Some("rename".to_string()),
            destination_name: None,
        }
    }
}

fn value_as_bool(value: &Value) -> Option<bool> {
    match value {
        Value::Bool(value) => Some(*value),
        Value::String(value) if value.eq_ignore_ascii_case("true") => Some(true),
        Value::String(value) if value.eq_ignore_ascii_case("false") => Some(false),
        _ => None,
    }
}

fn value_bool_arg(value: Option<&Value>, field_name: &str) -> Option<bool> {
    match value {
        None | Some(Value::Null) => None,
        Some(Value::Object(map)) => map
            .get(field_name)
            .or_else(|| map.get("open_after_import"))
            .and_then(value_as_bool),
        Some(value) => value_as_bool(value),
    }
}

fn value_string_field(value: Option<&Value>, camel_name: &str, snake_name: &str) -> Option<String> {
    value
        .and_then(Value::as_object)
        .and_then(|map| map.get(camel_name).or_else(|| map.get(snake_name)))
        .and_then(Value::as_str)
        .map(str::to_string)
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ImportWorkspaceResult {
    pub workspace_root: String,
    pub opened: bool,
    pub workspace: Option<CurrentWorkspace>,
    pub warnings: Vec<HealthCheckItem>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ExternalPathInfo {
    pub path: String,
    pub project_id: Option<String>,
    pub project_name: Option<String>,
    pub path_type: String,
    pub exists: bool,
    pub impact: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PathConversionCandidate {
    pub id: String,
    pub table_name: String,
    pub record_id: String,
    pub column_name: String,
    pub project_id: Option<String>,
    pub current_path: String,
    pub relative_path: String,
    pub reason: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PathConversionResult {
    pub dry_run: bool,
    pub candidates: Vec<PathConversionCandidate>,
    pub applied: usize,
    pub backup_path: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RepairWorkspaceResult {
    pub repaired: usize,
    pub backup_path: Option<String>,
    pub health: HealthCheckResult,
}

fn now_stamp() -> String {
    Local::now().format("%Y-%m-%d-%H-%M-%S").to_string()
}

fn today_stamp() -> String {
    Local::now().format("%Y-%m-%d").to_string()
}

fn sanitize_file_part(input: &str) -> String {
    let mut out = String::new();
    for ch in input.trim().chars() {
        if ch.is_ascii_alphanumeric() || ch == '-' || ch == '_' || ch == '.' {
            out.push(ch);
        } else if ch.is_alphanumeric() {
            out.push(ch);
        } else {
            out.push('_');
        }
    }
    let trimmed = out.trim_matches('_').to_string();
    if trimmed.is_empty() {
        "LamberWorkspace".to_string()
    } else {
        trimmed
    }
}

fn to_zip_name(path: &Path) -> String {
    path.to_string_lossy().replace('\\', "/")
}

fn sqlite_quote_path(path: &Path) -> String {
    path.to_string_lossy().replace('\'', "''")
}

fn backup_database_with_conn(conn: &Connection, dest: &Path) -> Result<(), String> {
    if let Some(parent) = dest.parent() {
        fs::create_dir_all(parent).map_err(|e| format!("创建备份目录失败: {}", e))?;
    }
    if dest.exists() {
        return Err(format!("备份文件已存在: {}", dest.display()));
    }
    let _ = conn.execute_batch("PRAGMA wal_checkpoint(FULL);");
    let sql = format!("VACUUM INTO '{}';", sqlite_quote_path(dest));
    conn.execute_batch(&sql)
        .map_err(|e| format!("创建 SQLite 备份失败: {}", e))
}

fn create_backup_with_name(
    root: &Path,
    conn: &Connection,
    file_name: &str,
) -> Result<PathBuf, String> {
    let backup_dir = root.join(BACKUPS_DIR);
    fs::create_dir_all(&backup_dir).map_err(|e| format!("创建备份目录失败: {}", e))?;
    workspace::mark_path_hidden_if_supported(&backup_dir);
    let dest = backup_dir.join(file_name);
    backup_database_with_conn(conn, &dest)?;
    Ok(dest)
}

pub(crate) fn ensure_daily_workspace_backup(
    root: &Path,
    conn: &Connection,
) -> Result<Option<String>, String> {
    let backup_dir = root.join(BACKUPS_DIR);
    fs::create_dir_all(&backup_dir).map_err(|e| format!("创建备份目录失败: {}", e))?;
    workspace::mark_path_hidden_if_supported(&backup_dir);
    let daily_name = format!("lamber-{}.sqlite", today_stamp());
    let daily_path = backup_dir.join(&daily_name);
    if daily_path.exists() {
        prune_backups(&backup_dir, 20)?;
        return Ok(None);
    }
    backup_database_with_conn(conn, &daily_path)?;
    prune_backups(&backup_dir, 20)?;
    Ok(Some(daily_path.to_string_lossy().to_string()))
}

fn prune_backups(backup_dir: &Path, keep: usize) -> Result<(), String> {
    let mut backups = list_backups_in_dir(backup_dir)?;
    backups.sort_by(|a, b| b.file_name.cmp(&a.file_name));
    for item in backups.into_iter().skip(keep) {
        let _ = fs::remove_file(item.path);
    }
    Ok(())
}

fn list_backups_in_dir(backup_dir: &Path) -> Result<Vec<WorkspaceBackupInfo>, String> {
    if !backup_dir.exists() {
        return Ok(Vec::new());
    }
    let mut list = Vec::new();
    for entry in fs::read_dir(backup_dir).map_err(|e| format!("读取备份目录失败: {}", e))? {
        let entry = match entry {
            Ok(v) => v,
            Err(_) => continue,
        };
        let path = entry.path();
        if !path.is_file() || path.extension().and_then(|s| s.to_str()) != Some("sqlite") {
            continue;
        }
        let meta = match fs::metadata(&path) {
            Ok(m) => m,
            Err(_) => continue,
        };
        let file_name = path
            .file_name()
            .map(|n| n.to_string_lossy().to_string())
            .unwrap_or_default();
        let created_at = meta
            .modified()
            .ok()
            .map(|t| chrono::DateTime::<Local>::from(t).to_rfc3339())
            .unwrap_or_default();
        list.push(WorkspaceBackupInfo {
            id: file_name.clone(),
            file_name: file_name.clone(),
            path: path.to_string_lossy().to_string(),
            created_at,
            size_bytes: meta.len(),
            is_daily: file_name.len() == "lamber-YYYY-MM-DD.sqlite".len(),
        });
    }
    list.sort_by(|a, b| b.file_name.cmp(&a.file_name));
    Ok(list)
}

#[tauri::command]
pub async fn create_workspace_backup(
    runtime: State<'_, Arc<WorkspaceRuntime>>,
) -> Result<WorkspaceBackupInfo, String> {
    let ws = runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    let root = PathBuf::from(&ws.workspace_root);
    let file_name = format!("lamber-{}.sqlite", now_stamp());
    let path = create_backup_with_name(&root, &conn, &file_name)?;
    let info = list_backups_in_dir(&root.join(BACKUPS_DIR))?
        .into_iter()
        .find(|item| item.file_name == file_name)
        .ok_or_else(|| "备份已创建，但无法读取备份元数据".to_string())?;
    Ok(WorkspaceBackupInfo {
        path: path.to_string_lossy().to_string(),
        ..info
    })
}

#[tauri::command]
pub async fn list_workspace_backups(
    runtime: State<'_, Arc<WorkspaceRuntime>>,
) -> Result<Vec<WorkspaceBackupInfo>, String> {
    let ws = runtime.require_workspace()?;
    list_backups_in_dir(&PathBuf::from(ws.workspace_root).join(BACKUPS_DIR))
}

fn safe_backup_path(root: &Path, backup_id: &str) -> Result<PathBuf, String> {
    if backup_id.contains('/') || backup_id.contains('\\') || backup_id.contains("..") {
        return Err("非法备份标识".to_string());
    }
    let path = root.join(BACKUPS_DIR).join(backup_id);
    if path.extension().and_then(|s| s.to_str()) != Some("sqlite") {
        return Err("备份文件必须是 sqlite 文件".to_string());
    }
    Ok(path)
}

#[tauri::command]
pub async fn delete_workspace_backup(
    runtime: State<'_, Arc<WorkspaceRuntime>>,
    backup_id: String,
) -> Result<(), String> {
    let ws = runtime.require_workspace()?;
    let path = safe_backup_path(Path::new(&ws.workspace_root), &backup_id)?;
    if path.exists() {
        fs::remove_file(path).map_err(|e| format!("删除备份失败: {}", e))?;
    }
    Ok(())
}

fn validate_sqlite_file(path: &Path) -> Result<(), String> {
    let conn = Connection::open(path).map_err(|e| format!("打开备份数据库失败: {}", e))?;
    let result: String = conn
        .query_row("PRAGMA integrity_check;", [], |row| row.get(0))
        .map_err(|e| format!("数据库完整性检查失败: {}", e))?;
    if result.eq_ignore_ascii_case("ok") {
        Ok(())
    } else {
        Err(format!("数据库完整性检查未通过: {}", result))
    }
}

#[tauri::command]
pub async fn restore_workspace_backup(
    app: AppHandle,
    runtime: State<'_, Arc<WorkspaceRuntime>>,
    backup_id: String,
) -> Result<CurrentWorkspace, String> {
    let ws = runtime.require_workspace()?;
    let root = PathBuf::from(&ws.workspace_root);
    let backup_path = safe_backup_path(&root, &backup_id)?;
    if !backup_path.exists() {
        return Err("指定备份不存在".to_string());
    }
    validate_sqlite_file(&backup_path)?;

    let pre_backup_path = {
        let db = runtime.require_db()?;
        let conn = db.lock().map_err(|e| e.to_string())?;
        create_backup_with_name(
            &root,
            &conn,
            &format!("lamber-before-restore-{}.sqlite", now_stamp()),
        )?
    };

    runtime.close_database()?;

    let db_file = workspace::db_path(&root);
    let restore_tmp = root.join(format!(".lamber.sqlite.restore-{}.tmp", now_stamp()));
    let old_tmp = root.join(format!(".lamber.sqlite.previous-{}.tmp", now_stamp()));

    let rollback = |runtime: &WorkspaceRuntime,
                    app: &AppHandle,
                    root: &Path,
                    old_tmp: &Path,
                    db_file: &Path|
     -> Result<(), String> {
        if old_tmp.exists() {
            let _ = fs::remove_file(db_file);
            let _ = fs::rename(old_tmp, db_file);
        }
        let _ = workspace::open_workspace_internal(app, runtime, root);
        Ok(())
    };

    fs::copy(&backup_path, &restore_tmp).map_err(|e| format!("复制备份到临时文件失败: {}", e))?;
    if let Err(err) = validate_sqlite_file(&restore_tmp) {
        let _ = fs::remove_file(&restore_tmp);
        let _ = workspace::open_workspace_internal(&app, &runtime, &root);
        return Err(err);
    }

    if db_file.exists() {
        if let Err(err) = fs::rename(&db_file, &old_tmp) {
            let _ = fs::remove_file(&restore_tmp);
            let _ = workspace::open_workspace_internal(&app, &runtime, &root);
            return Err(format!("替换数据库前释放或移动原数据库失败: {}", err));
        }
    }

    if let Err(err) = fs::rename(&restore_tmp, &db_file) {
        let _ = rollback(&runtime, &app, &root, &old_tmp, &db_file);
        return Err(format!("替换数据库失败，已尝试回滚: {}", err));
    }

    match workspace::open_workspace_internal(&app, &runtime, &root) {
        Ok(workspace) => {
            let _ = fs::remove_file(&old_tmp);
            let _ = pre_backup_path;
            Ok(workspace)
        }
        Err(err) => {
            let _ = rollback(&runtime, &app, &root, &old_tmp, &db_file);
            Err(format!("恢复后重新打开工作区失败，已尝试回滚: {}", err))
        }
    }
}

fn push_item(
    items: &mut Vec<HealthCheckItem>,
    id: impl Into<String>,
    severity: &str,
    category: &str,
    message: impl Into<String>,
    detail: Option<String>,
    repairable: bool,
    repair_action: Option<&str>,
) {
    items.push(HealthCheckItem {
        id: id.into(),
        severity: severity.to_string(),
        category: category.to_string(),
        message: message.into(),
        detail,
        repairable,
        repair_action: repair_action.map(|s| s.to_string()),
    });
}

fn table_exists(conn: &Connection, table: &str) -> bool {
    conn.query_row(
        "SELECT EXISTS(SELECT 1 FROM sqlite_master WHERE type='table' AND name=?1)",
        [table],
        |row| row.get::<_, bool>(0),
    )
    .unwrap_or(false)
}

fn index_exists(conn: &Connection, index: &str) -> bool {
    conn.query_row(
        "SELECT EXISTS(SELECT 1 FROM sqlite_master WHERE type='index' AND name=?1)",
        [index],
        |row| row.get::<_, bool>(0),
    )
    .unwrap_or(false)
}

fn is_absolute_path_str(value: &str) -> bool {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        return false;
    }
    Path::new(trimmed).is_absolute()
        || trimmed.starts_with('/')
        || trimmed.starts_with("\\\\")
        || trimmed.starts_with("//")
        || (trimmed.len() > 2
            && trimmed.as_bytes()[1] == b':'
            && (trimmed.as_bytes()[2] == b'\\' || trimmed.as_bytes()[2] == b'/'))
}

fn normalize_compare(value: &str) -> String {
    let normalized = value.replace('\\', "/");
    if cfg!(windows) {
        normalized.to_ascii_lowercase()
    } else {
        normalized
    }
}

fn path_inside_workspace_str(workspace_root: &Path, value: &str) -> bool {
    if !is_absolute_path_str(value) {
        return false;
    }
    let target = PathBuf::from(value);
    if let (Ok(ws), Ok(tgt)) = (fs::canonicalize(workspace_root), fs::canonicalize(&target)) {
        return tgt.starts_with(ws);
    }
    let ws_str = normalize_compare(&workspace_root.to_string_lossy());
    let target_str = normalize_compare(value);
    target_str == ws_str
        || target_str.starts_with(&(ws_str.trim_end_matches('/').to_string() + "/"))
}

fn relative_from_workspace(workspace_root: &Path, value: &str) -> Option<String> {
    if path_inside_workspace_str(workspace_root, value) {
        Some(workspace::to_relative_workspace_path(
            workspace_root,
            Path::new(value),
        ))
    } else {
        None
    }
}

fn resolved_path_exists(workspace_root: &Path, value: &str) -> bool {
    if value.trim().is_empty() {
        return false;
    }
    let p = Path::new(value);
    if p.is_absolute() {
        p.exists()
    } else {
        workspace_root.join(value).exists()
    }
}

fn path_exists_raw(value: &str) -> bool {
    !value.trim().is_empty() && Path::new(value).exists()
}

fn collect_external_paths(
    app: &AppHandle,
    conn: &Connection,
    workspace_root: &Path,
) -> Result<Vec<ExternalPathInfo>, String> {
    let mut list = Vec::new();
    let mut project_names = BTreeMap::new();
    if table_exists(conn, "projects") {
        let mut stmt = conn
            .prepare("SELECT id, name FROM projects")
            .map_err(|e| e.to_string())?;
        let rows = stmt
            .query_map([], |row| {
                Ok((row.get::<_, String>(0)?, row.get::<_, String>(1)?))
            })
            .map_err(|e| e.to_string())?;
        for row in rows {
            if let Ok((id, name)) = row {
                project_names.insert(id, name);
            }
        }
    }

    if table_exists(conn, "projects") {
        let mut stmt = conn
            .prepare("SELECT id, name, folder_path, linked_folder_external_path FROM projects")
            .map_err(|e| e.to_string())?;
        let rows = stmt
            .query_map([], |row| {
                Ok((
                    row.get::<_, String>(0)?,
                    row.get::<_, String>(1)?,
                    row.get::<_, Option<String>>(2)?,
                    row.get::<_, Option<String>>(3)?,
                ))
            })
            .map_err(|e| e.to_string())?;
        for row in rows {
            let (project_id, project_name, folder_path, external_path) =
                row.map_err(|e| e.to_string())?;
            for (path_type, path) in [
                ("project_folder", folder_path),
                ("linked_folder_external_path", external_path),
            ] {
                if let Some(path) = path {
                    if is_absolute_path_str(&path)
                        && !path_inside_workspace_str(workspace_root, &path)
                    {
                        list.push(ExternalPathInfo {
                            exists: path_exists_raw(&path),
                            path,
                            project_id: Some(project_id.clone()),
                            project_name: Some(project_name.clone()),
                            path_type: path_type.to_string(),
                            impact: "复制或导出 Workspace 时不会自动迁移此外部项目目录".to_string(),
                        });
                    }
                }
            }
        }
    }

    if table_exists(conn, "project_files") {
        let mut stmt = conn
            .prepare("SELECT project_id, file_name, file_path, original_path, managed_path, absolute_path_snapshot FROM project_files")
            .map_err(|e| e.to_string())?;
        let rows = stmt
            .query_map([], |row| {
                Ok((
                    row.get::<_, String>(0)?,
                    row.get::<_, String>(1)?,
                    row.get::<_, String>(2)?,
                    row.get::<_, Option<String>>(3)?,
                    row.get::<_, Option<String>>(4)?,
                    row.get::<_, Option<String>>(5)?,
                ))
            })
            .map_err(|e| e.to_string())?;
        for row in rows {
            let (project_id, file_name, file_path, original_path, managed_path, snapshot) =
                row.map_err(|e| e.to_string())?;
            for (path_type, path) in [
                ("project_file", Some(file_path)),
                ("project_file_original", original_path),
                ("project_file_managed", managed_path),
                ("project_file_snapshot", snapshot),
            ] {
                if let Some(path) = path {
                    if is_absolute_path_str(&path)
                        && !path_inside_workspace_str(workspace_root, &path)
                    {
                        list.push(ExternalPathInfo {
                            exists: path_exists_raw(&path),
                            path,
                            project_id: Some(project_id.clone()),
                            project_name: project_names.get(&project_id).cloned(),
                            path_type: path_type.to_string(),
                            impact: format!("文件 {} 位于工作区外部", file_name),
                        });
                    }
                }
            }
        }
    }

    if table_exists(conn, "project_template_states") {
        let mut stmt = conn
            .prepare("SELECT project_id, template_name, template_path FROM project_template_states WHERE template_path IS NOT NULL AND template_path <> ''")
            .map_err(|e| e.to_string())?;
        let rows = stmt
            .query_map([], |row| {
                Ok((
                    row.get::<_, String>(0)?,
                    row.get::<_, Option<String>>(1)?,
                    row.get::<_, String>(2)?,
                ))
            })
            .map_err(|e| e.to_string())?;
        for row in rows {
            let (project_id, template_name, path) = row.map_err(|e| e.to_string())?;
            if is_absolute_path_str(&path) && !path_inside_workspace_str(workspace_root, &path) {
                list.push(ExternalPathInfo {
                    exists: path_exists_raw(&path),
                    path,
                    project_id: Some(project_id.clone()),
                    project_name: project_names.get(&project_id).cloned(),
                    path_type: "template_path".to_string(),
                    impact: format!("模板 {} 位于工作区外部", template_name.unwrap_or_default()),
                });
            }
        }
    }

    if table_exists(conn, "project_template_assets") {
        let mut stmt = conn
            .prepare("SELECT project_id, original_file_name, absolute_path_snapshot FROM project_template_assets WHERE deleted_at IS NULL")
            .map_err(|e| e.to_string())?;
        let rows = stmt
            .query_map([], |row| {
                Ok((
                    row.get::<_, String>(0)?,
                    row.get::<_, Option<String>>(1)?,
                    row.get::<_, String>(2)?,
                ))
            })
            .map_err(|e| e.to_string())?;
        for row in rows {
            let (project_id, original_file_name, path) = row.map_err(|e| e.to_string())?;
            if is_absolute_path_str(&path) && !path_inside_workspace_str(workspace_root, &path) {
                list.push(ExternalPathInfo {
                    exists: path_exists_raw(&path),
                    path,
                    project_id: Some(project_id.clone()),
                    project_name: project_names.get(&project_id).cloned(),
                    path_type: "template_asset_snapshot".to_string(),
                    impact: format!(
                        "模板资产 {} 的快照路径位于工作区外部",
                        original_file_name.unwrap_or_default()
                    ),
                });
            }
        }
    }

    if table_exists(conn, "project_roots") {
        let mut stmt = conn
            .prepare("SELECT id, name, root_path FROM project_roots")
            .map_err(|e| e.to_string())?;
        let rows = stmt
            .query_map([], |row| {
                Ok((
                    row.get::<_, String>(0)?,
                    row.get::<_, String>(1)?,
                    row.get::<_, String>(2)?,
                ))
            })
            .map_err(|e| e.to_string())?;
        for row in rows {
            let (id, name, path) = row.map_err(|e| e.to_string())?;
            if !id.starts_with("workspace_root_")
                && is_absolute_path_str(&path)
                && !path_inside_workspace_str(workspace_root, &path)
            {
                list.push(ExternalPathInfo {
                    exists: path_exists_raw(&path),
                    path,
                    project_id: None,
                    project_name: None,
                    path_type: "external_project_root".to_string(),
                    impact: format!("项目根目录 {} 是外部路径，只检查存在性，不自动改写", name),
                });
            }
        }
    }

    let config = ConfigManager::new(app).load();
    for (module, path) in config.module_paths {
        if RETIRED_MODULE_IDS.contains(&module.as_str()) {
            continue;
        }
        if is_absolute_path_str(&path) && !path_inside_workspace_str(workspace_root, &path) {
            list.push(ExternalPathInfo {
                exists: path_exists_raw(&path),
                path,
                project_id: None,
                project_name: None,
                path_type: format!("module_path:{}", module),
                impact: "模块模板/输出目录位于工作区外部，导出 Workspace 不会自动包含".to_string(),
            });
        }
    }

    Ok(list)
}

fn collect_internal_absolute_path_candidates(
    conn: &Connection,
    workspace_root: &Path,
) -> Result<Vec<PathConversionCandidate>, String> {
    let mut out = Vec::new();
    let mut push_candidate = |table_name: &str,
                              record_id: String,
                              column_name: &str,
                              project_id: Option<String>,
                              current_path: String,
                              reason: &str| {
        if let Some(relative_path) = relative_from_workspace(workspace_root, &current_path) {
            out.push(PathConversionCandidate {
                id: format!("pathconv:{}:{}:{}", table_name, record_id, column_name),
                table_name: table_name.to_string(),
                record_id,
                column_name: column_name.to_string(),
                project_id,
                current_path,
                relative_path,
                reason: reason.to_string(),
            });
        }
    };

    if table_exists(conn, "projects") {
        let mut stmt = conn
            .prepare(
                "SELECT id, folder_path, relative_path, linked_folder_relative_path, linked_folder_external_path, main_document_path, main_budget_file_path FROM projects",
            )
            .map_err(|e| e.to_string())?;
        let rows = stmt
            .query_map([], |row| {
                Ok((
                    row.get::<_, String>(0)?,
                    row.get::<_, Option<String>>(1)?,
                    row.get::<_, Option<String>>(2)?,
                    row.get::<_, Option<String>>(3)?,
                    row.get::<_, Option<String>>(4)?,
                    row.get::<_, Option<String>>(5)?,
                    row.get::<_, Option<String>>(6)?,
                ))
            })
            .map_err(|e| e.to_string())?;
        for row in rows {
            let (id, folder, rel, linked_rel, linked_ext, main_doc, main_budget) =
                row.map_err(|e| e.to_string())?;
            for (column, value) in [
                ("folder_path", folder),
                ("relative_path", rel),
                ("linked_folder_relative_path", linked_rel),
                ("linked_folder_external_path", linked_ext),
                ("main_document_path", main_doc),
                ("main_budget_file_path", main_budget),
            ] {
                if let Some(value) = value {
                    push_candidate(
                        "projects",
                        id.clone(),
                        column,
                        Some(id.clone()),
                        value,
                        "项目路径位于当前 Workspace 内，可转为相对路径",
                    );
                }
            }
        }
    }

    if table_exists(conn, "project_files") {
        let mut stmt = conn
            .prepare("SELECT id, project_id, file_path, original_path, managed_path, relative_path, absolute_path_snapshot FROM project_files")
            .map_err(|e| e.to_string())?;
        let rows = stmt
            .query_map([], |row| {
                Ok((
                    row.get::<_, String>(0)?,
                    row.get::<_, String>(1)?,
                    row.get::<_, String>(2)?,
                    row.get::<_, Option<String>>(3)?,
                    row.get::<_, Option<String>>(4)?,
                    row.get::<_, Option<String>>(5)?,
                    row.get::<_, Option<String>>(6)?,
                ))
            })
            .map_err(|e| e.to_string())?;
        for row in rows {
            let (id, project_id, file_path, original, managed, rel, snapshot) =
                row.map_err(|e| e.to_string())?;
            for (column, value) in [
                ("file_path", Some(file_path)),
                ("original_path", original),
                ("managed_path", managed),
                ("relative_path", rel),
                ("absolute_path_snapshot", snapshot),
            ] {
                if let Some(value) = value {
                    push_candidate(
                        "project_files",
                        id.clone(),
                        column,
                        Some(project_id.clone()),
                        value,
                        "项目文件路径位于当前 Workspace 内，可转为相对路径",
                    );
                }
            }
        }
    }

    if table_exists(conn, "project_directories") {
        let mut stmt = conn
            .prepare("SELECT id, project_id, relative_path FROM project_directories")
            .map_err(|e| e.to_string())?;
        let rows = stmt
            .query_map([], |row| {
                Ok((
                    row.get::<_, String>(0)?,
                    row.get::<_, String>(1)?,
                    row.get::<_, String>(2)?,
                ))
            })
            .map_err(|e| e.to_string())?;
        for row in rows {
            let (id, project_id, value) = row.map_err(|e| e.to_string())?;
            push_candidate(
                "project_directories",
                id,
                "relative_path",
                Some(project_id),
                value,
                "项目目录路径位于当前 Workspace 内，可转为相对路径",
            );
        }
    }

    if table_exists(conn, "project_template_states") {
        let mut stmt = conn
            .prepare("SELECT id, project_id, template_path FROM project_template_states WHERE template_path IS NOT NULL AND template_path <> ''")
            .map_err(|e| e.to_string())?;
        let rows = stmt
            .query_map([], |row| {
                Ok((
                    row.get::<_, String>(0)?,
                    row.get::<_, String>(1)?,
                    row.get::<_, String>(2)?,
                ))
            })
            .map_err(|e| e.to_string())?;
        for row in rows {
            let (id, project_id, value) = row.map_err(|e| e.to_string())?;
            push_candidate(
                "project_template_states",
                id,
                "template_path",
                Some(project_id),
                value,
                "模板路径位于当前 Workspace 内，可转为相对路径",
            );
        }
    }

    if table_exists(conn, "project_template_assets") {
        let mut stmt = conn
            .prepare("SELECT id, project_id, relative_path, absolute_path_snapshot FROM project_template_assets WHERE deleted_at IS NULL")
            .map_err(|e| e.to_string())?;
        let rows = stmt
            .query_map([], |row| {
                Ok((
                    row.get::<_, String>(0)?,
                    row.get::<_, String>(1)?,
                    row.get::<_, String>(2)?,
                    row.get::<_, String>(3)?,
                ))
            })
            .map_err(|e| e.to_string())?;
        for row in rows {
            let (id, project_id, rel, snapshot) = row.map_err(|e| e.to_string())?;
            push_candidate(
                "project_template_assets",
                id.clone(),
                "relative_path",
                Some(project_id.clone()),
                rel,
                "模板资产路径位于当前 Workspace 内，可转为相对路径",
            );
            push_candidate(
                "project_template_assets",
                id,
                "absolute_path_snapshot",
                Some(project_id),
                snapshot,
                "模板资产快照路径位于当前 Workspace 内，可转为相对路径",
            );
        }
    }

    Ok(out)
}

fn apply_path_candidates(
    conn: &mut Connection,
    candidates: &[PathConversionCandidate],
) -> Result<usize, String> {
    let tx = conn.transaction().map_err(|e| e.to_string())?;
    let mut applied = 0;
    for candidate in candidates {
        match (
            candidate.table_name.as_str(),
            candidate.column_name.as_str(),
        ) {
            ("projects", "linked_folder_external_path") => {
                tx.execute(
                    "UPDATE projects SET linked_folder_type='internal', linked_folder_relative_path=?1, linked_folder_external_path=NULL, folder_path=?1, relative_path=?1 WHERE id=?2",
                    params![candidate.relative_path, candidate.record_id],
                )
                .map_err(|e| e.to_string())?;
            }
            ("projects", column)
                if matches!(
                    column,
                    "folder_path"
                        | "relative_path"
                        | "linked_folder_relative_path"
                        | "main_document_path"
                        | "main_budget_file_path"
                ) =>
            {
                let sql = format!("UPDATE projects SET {}=?1 WHERE id=?2", column);
                tx.execute(&sql, params![candidate.relative_path, candidate.record_id])
                    .map_err(|e| e.to_string())?;
            }
            ("project_files", column)
                if matches!(
                    column,
                    "file_path"
                        | "original_path"
                        | "managed_path"
                        | "relative_path"
                        | "absolute_path_snapshot"
                ) =>
            {
                let sql = format!("UPDATE project_files SET {}=?1 WHERE id=?2", column);
                tx.execute(&sql, params![candidate.relative_path, candidate.record_id])
                    .map_err(|e| e.to_string())?;
            }
            ("project_directories", "relative_path") => {
                tx.execute(
                    "UPDATE project_directories SET relative_path=?1 WHERE id=?2",
                    params![candidate.relative_path, candidate.record_id],
                )
                .map_err(|e| e.to_string())?;
            }
            ("project_template_states", "template_path") => {
                tx.execute(
                    "UPDATE project_template_states SET template_path=?1, template_path_type='workspace' WHERE id=?2",
                    params![candidate.relative_path, candidate.record_id],
                )
                .map_err(|e| e.to_string())?;
            }
            ("project_template_assets", column)
                if matches!(column, "relative_path" | "absolute_path_snapshot") =>
            {
                let sql = format!(
                    "UPDATE project_template_assets SET {}=?1 WHERE id=?2",
                    column
                );
                tx.execute(&sql, params![candidate.relative_path, candidate.record_id])
                    .map_err(|e| e.to_string())?;
            }
            _ => continue,
        }
        applied += 1;
    }
    tx.commit().map_err(|e| e.to_string())?;
    Ok(applied)
}

#[tauri::command]
pub async fn convert_internal_absolute_paths_to_relative(
    runtime: State<'_, Arc<WorkspaceRuntime>>,
    dry_run: Option<bool>,
) -> Result<PathConversionResult, String> {
    let ws = runtime.require_workspace()?;
    let root = PathBuf::from(&ws.workspace_root);
    let db = runtime.require_db()?;
    let mut conn = db.lock().map_err(|e| e.to_string())?;
    let candidates = collect_internal_absolute_path_candidates(&conn, &root)?;
    let dry_run = dry_run.unwrap_or(true);
    if dry_run {
        return Ok(PathConversionResult {
            dry_run,
            candidates,
            applied: 0,
            backup_path: None,
        });
    }
    let backup_path = create_backup_with_name(
        &root,
        &conn,
        &format!("lamber-before-path-repair-{}.sqlite", now_stamp()),
    )?;
    let applied = apply_path_candidates(&mut conn, &candidates)?;
    Ok(PathConversionResult {
        dry_run,
        candidates,
        applied,
        backup_path: Some(backup_path.to_string_lossy().to_string()),
    })
}

#[tauri::command]
pub async fn list_external_paths(
    app: AppHandle,
    runtime: State<'_, Arc<WorkspaceRuntime>>,
) -> Result<Vec<ExternalPathInfo>, String> {
    let ws = runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    collect_external_paths(&app, &conn, Path::new(&ws.workspace_root))
}

#[tauri::command]
pub async fn inspect_workspace_paths(
    app: AppHandle,
    runtime: State<'_, Arc<WorkspaceRuntime>>,
) -> Result<(Vec<PathConversionCandidate>, Vec<ExternalPathInfo>), String> {
    let ws = runtime.require_workspace()?;
    let root = PathBuf::from(&ws.workspace_root);
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    Ok((
        collect_internal_absolute_path_candidates(&conn, &root)?,
        collect_external_paths(&app, &conn, &root)?,
    ))
}

fn run_workspace_health_check_internal(
    app: &AppHandle,
    ws: &CurrentWorkspace,
    conn: &Connection,
) -> Result<HealthCheckResult, String> {
    let mut items = Vec::new();
    let root = PathBuf::from(&ws.workspace_root);
    let manifest_path = workspace::manifest_path(&root);
    let db_path = workspace::db_path(&root);

    for (name, path, required, repairable) in [
        (workspace::MANIFEST_FILE, manifest_path.clone(), true, false),
        (workspace::DATABASE_FILE, db_path.clone(), true, false),
        (BACKUPS_DIR, root.join(BACKUPS_DIR), false, true),
        (EXPORTS_DIR, root.join(EXPORTS_DIR), false, true),
        (
            PROJECT_ASSETS_DIR,
            root.join(PROJECT_ASSETS_DIR),
            false,
            true,
        ),
    ] {
        if !path.exists() {
            push_item(
                &mut items,
                format!("structure:missing-dir:{}", name),
                if required { "error" } else { "warning" },
                "基础结构",
                format!("缺少工作区资源 {}", name),
                Some(path.to_string_lossy().to_string()),
                repairable,
                if repairable {
                    Some("创建缺失目录")
                } else {
                    None
                },
            );
        }
    }

    if let Ok(raw) = fs::read_to_string(&manifest_path) {
        match serde_json::from_str::<WorkspaceManifest>(&raw) {
            Ok(manifest) => {
                if manifest.workspace_id.trim().is_empty() {
                    push_item(
                        &mut items,
                        "structure:workspace-id-missing",
                        "error",
                        "基础结构",
                        "workspaceId 为空",
                        None,
                        false,
                        None,
                    );
                }
                if manifest.workspace_version > workspace::WORKSPACE_VERSION {
                    push_item(
                        &mut items,
                        "structure:unsupported-workspace-version",
                        "error",
                        "基础结构",
                        "workspaceVersion 高于当前应用支持版本",
                        Some(format!(
                            "{} > {}",
                            manifest.workspace_version,
                            workspace::WORKSPACE_VERSION
                        )),
                        false,
                        None,
                    );
                }
            }
            Err(err) => push_item(
                &mut items,
                "structure:manifest-invalid-json",
                "error",
                "基础结构",
                "工作区 manifest 无法解析",
                Some(err.to_string()),
                false,
                None,
            ),
        }
    }

    let foreign_keys: i64 = conn
        .query_row("PRAGMA foreign_keys;", [], |row| row.get(0))
        .unwrap_or(0);
    if foreign_keys != 1 {
        push_item(
            &mut items,
            "database:foreign-keys-disabled",
            "error",
            "数据库",
            "SQLite foreign_keys 未启用",
            None,
            false,
            None,
        );
    }
    let integrity: String = conn
        .query_row("PRAGMA integrity_check;", [], |row| row.get(0))
        .unwrap_or_else(|_| "failed".to_string());
    if !integrity.eq_ignore_ascii_case("ok") {
        push_item(
            &mut items,
            "database:integrity-check-failed",
            "error",
            "数据库",
            "SQLite integrity_check 未通过",
            Some(integrity),
            false,
            None,
        );
    }
    let schema_version: Option<String> = conn
        .query_row(
            "SELECT value FROM app_settings WHERE key='schema_version'",
            [],
            |row| row.get(0),
        )
        .optional()
        .unwrap_or(None);
    if schema_version
        .as_deref()
        .and_then(|v| v.parse::<i32>().ok())
        .unwrap_or(0)
        < 5
    {
        push_item(
            &mut items,
            "database:schema-version-old",
            "error",
            "数据库",
            "schema_version 低于第四阶段要求",
            schema_version,
            false,
            None,
        );
    }

    for table in [
        "projects",
        "project_lifecycle_states",
        "project_cashflow_states",
        "project_template_states",
        "project_template_assets",
        "benefit_schemes",
        "benefit_snapshots",
        "project_files",
        "project_directories",
        "project_roots",
        "project_settings",
    ] {
        if !table_exists(conn, table) {
            push_item(
                &mut items,
                format!("database:missing-table:{}", table),
                "error",
                "数据库",
                format!("缺少必要表 {}", table),
                None,
                false,
                None,
            );
        }
    }
    for index in [
        "idx_project_lifecycle_project_id",
        "idx_project_cashflow_project_id",
        "idx_project_template_states_project_id",
        "idx_project_template_assets_project_template",
    ] {
        if !index_exists(conn, index) {
            push_item(
                &mut items,
                format!("database:missing-index:{}", index),
                "warning",
                "数据库",
                format!("缺少建议索引 {}", index),
                None,
                false,
                None,
            );
        }
    }

    let path_candidates = collect_internal_absolute_path_candidates(conn, &root)?;
    if !path_candidates.is_empty() {
        push_item(
            &mut items,
            "paths:convert-internal-absolute",
            "warning",
            "路径",
            format!(
                "发现 {} 个 Workspace 内部绝对路径，可转为相对路径",
                path_candidates.len()
            ),
            Some(serde_json::to_string(&path_candidates).unwrap_or_default()),
            true,
            Some("转换内部绝对路径为相对路径"),
        );
    }

    for external in collect_external_paths(app, conn, &root)? {
        let module_id = external.path_type.strip_prefix("module_path:");
        let item_id = if let Some(module_id) = module_id {
            format!("paths:external-module-path:{}", module_id)
        } else {
            format!("paths:external:{}", items.len())
        };
        let repairable = module_id.is_some();
        push_item(
            &mut items,
            item_id,
            "warning",
            "外部路径",
            "以下文件或文件夹位于工作区外部，复制或导出 Workspace 时不会自动迁移",
            Some(serde_json::to_string(&external).unwrap_or_default()),
            repairable,
            if repairable {
                Some("重置模块目录到当前 Workspace 内")
            } else {
                None
            },
        );
    }

    let mut project_ids = HashSet::new();
    if table_exists(conn, "projects") {
        let mut stmt = conn
            .prepare("SELECT id, name, folder_name, relative_path, folder_path FROM projects")
            .map_err(|e| e.to_string())?;
        let rows = stmt
            .query_map([], |row| {
                Ok((
                    row.get::<_, String>(0)?,
                    row.get::<_, String>(1)?,
                    row.get::<_, Option<String>>(2)?,
                    row.get::<_, Option<String>>(3)?,
                    row.get::<_, Option<String>>(4)?,
                ))
            })
            .map_err(|e| e.to_string())?;
        for row in rows {
            let (project_id, name, folder_name, relative_path, folder_path) =
                row.map_err(|e| e.to_string())?;
            project_ids.insert(project_id.clone());
            let rel = relative_path.or(folder_path).unwrap_or_default();
            let dir_name = folder_name
                .or_else(|| {
                    Path::new(&rel)
                        .file_name()
                        .map(|n| n.to_string_lossy().to_string())
                })
                .unwrap_or_default();
            if workspace::is_reserved_workspace_entry_name(&dir_name) {
                push_item(
                    &mut items,
                    format!("project:reserved-name:{}", project_id),
                    "error",
                    "项目结构",
                    format!("项目 {} 的目录名与工作区系统保留名冲突", name),
                    Some(dir_name),
                    false,
                    None,
                );
            }
            if rel.trim().is_empty() {
                push_item(
                    &mut items,
                    format!("project:missing-relative-path:{}", project_id),
                    "error",
                    "项目结构",
                    format!("项目 {} 缺少 relative_path", name),
                    None,
                    false,
                    None,
                );
                continue;
            }
            let project_dir = workspace::resolve_workspace_path(&root, &rel);
            if !project_dir.exists() || !project_dir.is_dir() {
                push_item(
                    &mut items,
                    format!("project:missing-directory:{}", project_id),
                    "error",
                    "项目结构",
                    format!("数据库中存在项目 {}，但目录缺失", name),
                    Some(project_dir.to_string_lossy().to_string()),
                    false,
                    None,
                );
                continue;
            }
            let project_json_path = project_dir.join("project.json");
            if !project_json_path.exists() {
                push_item(
                    &mut items,
                    format!("project:missing-project-json:{}", project_id),
                    "warning",
                    "项目结构",
                    format!("项目 {} 缺少 project.json", name),
                    Some(project_json_path.to_string_lossy().to_string()),
                    true,
                    Some("重新生成 project.json"),
                );
            }
            for sub in ["assets", "documents", "analyses"] {
                let sub_path = project_dir.join(sub);
                if !sub_path.exists() || !sub_path.is_dir() {
                    push_item(
                        &mut items,
                        format!("project:missing-subdir:{}:{}", project_id, sub),
                        "warning",
                        "项目结构",
                        format!("项目 {} 缺少 {} 目录", name, sub),
                        Some(sub_path.to_string_lossy().to_string()),
                        true,
                        Some("创建缺失项目子目录"),
                    );
                }
            }
        }
    }

    if root.exists() {
        for entry in fs::read_dir(&root).map_err(|e| e.to_string())? {
            let entry = match entry {
                Ok(v) => v,
                Err(_) => continue,
            };
            let path = entry.path();
            if !path.is_dir() {
                continue;
            }
            let name = path
                .file_name()
                .map(|n| n.to_string_lossy().to_string())
                .unwrap_or_default();
            if name.starts_with('.') || workspace::is_reserved_workspace_entry_name(&name) {
                continue;
            }
            let project_json = path.join("project.json");
            if !project_json.exists() {
                continue;
            }
            if let Ok(raw) = fs::read_to_string(&project_json) {
                if let Ok(value) = serde_json::from_str::<serde_json::Value>(&raw) {
                    let project_id = value
                        .get("projectId")
                        .and_then(|v| v.as_str())
                        .unwrap_or_default()
                        .to_string();
                    if !project_id.is_empty() && !project_ids.contains(&project_id) {
                        push_item(
                            &mut items,
                            format!("project:unregistered-project-json:{}", project_id),
                            "warning",
                            "项目结构",
                            "workspace 根层发现 project.json，但数据库中没有对应项目记录",
                            Some(project_json.to_string_lossy().to_string()),
                            false,
                            Some("提示导入项目"),
                        );
                    }
                }
            }
        }
    }

    for (table, json_cols) in [
        (
            "project_lifecycle_states",
            vec![
                "profile_json",
                "parameters_json",
                "background_json",
                "input_payload_json",
            ],
        ),
        (
            "project_cashflow_states",
            vec![
                "payment_model_json",
                "yearly_cashflow_json",
                "sector_cashflow_json",
                "assumptions_json",
                "metrics_json",
            ],
        ),
        (
            "project_template_states",
            vec![
                "filled_data_json",
                "field_mapping_json",
                "output_config_json",
            ],
        ),
    ] {
        if !table_exists(conn, table) {
            continue;
        }
        let orphan_count: i64 = conn
            .query_row(
                &format!("SELECT COUNT(*) FROM {} s LEFT JOIN projects p ON p.id=s.project_id WHERE p.id IS NULL", table),
                [],
                |row| row.get(0),
            )
            .unwrap_or(0);
        if orphan_count > 0 {
            push_item(
                &mut items,
                format!("state:orphan-project:{}", table),
                "error",
                "第三阶段状态表",
                format!("{} 存在 {} 条孤儿 project_id 记录", table, orphan_count),
                None,
                false,
                None,
            );
        }
        let select_cols = format!("id, {}", json_cols.join(", "));
        let sql = format!("SELECT {} FROM {}", select_cols, table);
        let mut stmt = conn.prepare(&sql).map_err(|e| e.to_string())?;
        let rows = stmt
            .query_map([], |row| {
                let id: String = row.get(0)?;
                let mut vals = Vec::new();
                for idx in 0..json_cols.len() {
                    vals.push(row.get::<_, String>(idx + 1)?);
                }
                Ok((id, vals))
            })
            .map_err(|e| e.to_string())?;
        for row in rows {
            let (id, vals) = row.map_err(|e| e.to_string())?;
            for (idx, raw) in vals.into_iter().enumerate() {
                if serde_json::from_str::<serde_json::Value>(&raw).is_err() {
                    push_item(
                        &mut items,
                        format!("state:invalid-json:{}:{}:{}", table, id, json_cols[idx]),
                        "error",
                        "第三阶段状态表",
                        format!("{}.{} JSON 无法解析", table, json_cols[idx]),
                        Some(id.clone()),
                        false,
                        None,
                    );
                }
            }
        }
    }

    if table_exists(conn, "benefit_schemes") {
        let count: i64 = conn
            .query_row("SELECT COUNT(*) FROM benefit_schemes s LEFT JOIN projects p ON p.id=s.project_id WHERE p.id IS NULL", [], |row| row.get(0))
            .unwrap_or(0);
        if count > 0 {
            push_item(
                &mut items,
                "state:orphan-benefit-schemes",
                "error",
                "第三阶段状态表",
                format!("benefit_schemes 存在 {} 条孤儿 project_id", count),
                None,
                false,
                None,
            );
        }
    }
    if table_exists(conn, "benefit_snapshots") {
        let count: i64 = conn
            .query_row(
                "SELECT COUNT(*) FROM benefit_snapshots s LEFT JOIN projects p ON p.id=s.project_id LEFT JOIN benefit_schemes b ON b.id=s.scheme_id WHERE p.id IS NULL OR b.id IS NULL",
                [],
                |row| row.get(0),
            )
            .unwrap_or(0);
        if count > 0 {
            push_item(
                &mut items,
                "state:orphan-benefit-snapshots",
                "error",
                "第三阶段状态表",
                format!(
                    "benefit_snapshots 存在 {} 条孤儿 scheme_id/project_id",
                    count
                ),
                None,
                false,
                None,
            );
        }
    }

    if table_exists(conn, "project_template_assets") {
        let mut stmt = conn
            .prepare("SELECT id, project_id, template_name, relative_path, absolute_path_snapshot FROM project_template_assets WHERE deleted_at IS NULL")
            .map_err(|e| e.to_string())?;
        let rows = stmt
            .query_map([], |row| {
                Ok((
                    row.get::<_, String>(0)?,
                    row.get::<_, String>(1)?,
                    row.get::<_, String>(2)?,
                    row.get::<_, String>(3)?,
                    row.get::<_, String>(4)?,
                ))
            })
            .map_err(|e| e.to_string())?;
        for row in rows {
            let (id, project_id, template_name, relative_path, snapshot) =
                row.map_err(|e| e.to_string())?;
            if is_absolute_path_str(&relative_path)
                && !path_inside_workspace_str(&root, &relative_path)
            {
                push_item(
                    &mut items,
                    format!("template:external-asset:{}", id),
                    "warning",
                    "模板资源",
                    "模板资产 relative_path 指向工作区外部",
                    Some(relative_path.clone()),
                    false,
                    None,
                );
            }
            if !resolved_path_exists(&root, &relative_path) {
                push_item(
                    &mut items,
                    format!("template:missing-asset:{}", id),
                    "error",
                    "模板资源",
                    format!("模板 {} 的资产文件缺失", template_name),
                    Some(format!(
                        "projectId={}, relativePath={}, snapshot={}",
                        project_id, relative_path, snapshot
                    )),
                    false,
                    None,
                );
            }
        }
    }

    if table_exists(conn, "project_settings") {
        let mut stmt = conn
            .prepare("SELECT project_id, key, value FROM project_settings")
            .map_err(|e| e.to_string())?;
        let rows = stmt
            .query_map([], |row| {
                Ok((
                    row.get::<_, String>(0)?,
                    row.get::<_, String>(1)?,
                    row.get::<_, String>(2)?,
                ))
            })
            .map_err(|e| e.to_string())?;
        for row in rows {
            let (project_id, key, value) = row.map_err(|e| e.to_string())?;
            if value.contains("data:image/") {
                push_item(
                    &mut items,
                    format!("template:legacy-base64:{}:{}", project_id, key),
                    "warning",
                    "模板资源",
                    "发现旧版 base64 图片资源，建议通过模板保存流程迁移为文件资产",
                    Some(format!("projectId={}, key={}", project_id, key)),
                    false,
                    None,
                );
            }
        }
    }

    let status = if items.iter().any(|i| i.severity == "error") {
        "error"
    } else if items.iter().any(|i| i.severity == "warning") {
        "warning"
    } else {
        "normal"
    };
    Ok(HealthCheckResult {
        status: status.to_string(),
        items,
    })
}

#[tauri::command]
pub async fn run_workspace_health_check(
    app: AppHandle,
    runtime: State<'_, Arc<WorkspaceRuntime>>,
) -> Result<HealthCheckResult, String> {
    let ws = runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    run_workspace_health_check_internal(&app, &ws, &conn)
}

fn regenerate_project_json(
    conn: &Connection,
    workspace_root: &Path,
    project_id: &str,
) -> Result<(), String> {
    let row: Option<(
        String,
        String,
        Option<String>,
        Option<String>,
        Option<String>,
        Option<String>,
        String,
        String,
    )> = conn
        .query_row(
            "SELECT id, name, folder_name, relative_path, linked_folder_relative_path, folder_path, created_at, updated_at FROM projects WHERE id=?1",
            [project_id],
            |row| {
                Ok((
                    row.get::<_, String>(0)?,
                    row.get::<_, String>(1)?,
                    row.get::<_, Option<String>>(2)?,
                    row.get::<_, Option<String>>(3)?,
                    row.get::<_, Option<String>>(4)?,
                    row.get::<_, Option<String>>(5)?,
                    row.get::<_, String>(6)?,
                    row.get::<_, String>(7)?,
                ))
            },
        )
        .optional()
        .map_err(|e| e.to_string())?;
    let (
        id,
        name,
        folder_name,
        relative_path,
        linked_relative_path,
        folder_path,
        created_at,
        updated_at,
    ) = row.ok_or_else(|| "项目不存在".to_string())?;
    let path_value = [
        relative_path,
        linked_relative_path,
        folder_path,
        folder_name,
    ]
    .into_iter()
    .flatten()
    .find(|value| !value.trim().is_empty())
    .ok_or_else(|| "项目缺少目录路径".to_string())?;
    let project_dir = workspace::resolve_workspace_path(workspace_root, &path_value);
    let rel = if project_dir.is_absolute()
        && workspace::is_inside_workspace(workspace_root, &project_dir)
    {
        workspace::to_relative_workspace_path(workspace_root, &project_dir)
    } else {
        path_value.clone()
    };
    fs::create_dir_all(&project_dir).map_err(|e| format!("创建项目目录失败: {}", e))?;
    let payload = json!({
        "projectId": id,
        "name": name,
        "folderName": project_dir.file_name().map(|n| n.to_string_lossy().to_string()).unwrap_or_default(),
        "relativePath": rel,
        "createdAt": created_at,
        "updatedAt": updated_at,
        "source": "healthRepair"
    });
    fs::write(
        project_dir.join("project.json"),
        serde_json::to_string_pretty(&payload).map_err(|e| e.to_string())?,
    )
    .map_err(|e| format!("写入 project.json 失败: {}", e))
}

fn reset_module_path_to_workspace(
    app: &AppHandle,
    config_state: &State<'_, Mutex<AppConfig>>,
    workspace_root: &Path,
    module_id: &str,
) -> Result<(), String> {
    let module_id = module_id.trim();
    if module_id.is_empty() {
        return Err("模块 ID 不能为空".to_string());
    }

    let safe_module_name = workspace::sanitize_folder_name(module_id);
    let module_root = workspace_root
        .join(PROJECT_ASSETS_DIR)
        .join("modules")
        .join(safe_module_name);
    fs::create_dir_all(module_root.join("templates"))
        .map_err(|e| format!("创建模块模板目录失败: {}", e))?;
    fs::create_dir_all(module_root.join("output"))
        .map_err(|e| format!("创建模块输出目录失败: {}", e))?;

    workspace::mark_path_hidden_if_supported(&workspace_root.join(PROJECT_ASSETS_DIR));

    let module_root_str = module_root.to_string_lossy().to_string();
    let manager = ConfigManager::new(app);
    let mut config = config_state.lock().map_err(|e| e.to_string())?;
    config
        .module_paths
        .insert(module_id.to_string(), module_root_str);
    manager.save(&config)
}

#[tauri::command]
pub async fn repair_workspace_issues(
    app: AppHandle,
    runtime: State<'_, Arc<WorkspaceRuntime>>,
    config_state: State<'_, Mutex<AppConfig>>,
    issue_ids: Vec<String>,
) -> Result<RepairWorkspaceResult, String> {
    if issue_ids.is_empty() {
        let ws = runtime.require_workspace()?;
        let db = runtime.require_db()?;
        let conn = db.lock().map_err(|e| e.to_string())?;
        return Ok(RepairWorkspaceResult {
            repaired: 0,
            backup_path: None,
            health: run_workspace_health_check_internal(&app, &ws, &conn)?,
        });
    }
    let ws = runtime.require_workspace()?;
    let root = PathBuf::from(&ws.workspace_root);
    let db = runtime.require_db()?;
    let mut conn = db.lock().map_err(|e| e.to_string())?;
    let backup_path = create_backup_with_name(
        &root,
        &conn,
        &format!("lamber-before-repair-{}.sqlite", now_stamp()),
    )?;
    let mut repaired = 0;

    for issue_id in issue_ids {
        if let Some(name) = issue_id.strip_prefix("structure:missing-dir:") {
            if matches!(name, BACKUPS_DIR | EXPORTS_DIR | PROJECT_ASSETS_DIR) {
                fs::create_dir_all(root.join(name)).map_err(|e| format!("创建目录失败: {}", e))?;
                workspace::mark_path_hidden_if_supported(&root.join(name));
                repaired += 1;
            }
        } else if let Some(rest) = issue_id.strip_prefix("project:missing-subdir:") {
            let parts: Vec<&str> = rest.split(':').collect();
            if parts.len() == 2 && matches!(parts[1], "assets" | "documents" | "analyses") {
                let rel: Option<String> = conn
                    .query_row(
                        "SELECT COALESCE(relative_path, folder_path) FROM projects WHERE id=?1",
                        [parts[0]],
                        |row| row.get(0),
                    )
                    .optional()
                    .map_err(|e| e.to_string())?
                    .flatten();
                if let Some(rel) = rel {
                    fs::create_dir_all(
                        workspace::resolve_workspace_path(&root, &rel).join(parts[1]),
                    )
                    .map_err(|e| format!("创建项目子目录失败: {}", e))?;
                    repaired += 1;
                }
            }
        } else if let Some(project_id) = issue_id.strip_prefix("project:missing-project-json:") {
            regenerate_project_json(&conn, &root, project_id)?;
            repaired += 1;
        } else if let Some(module_id) = issue_id.strip_prefix("paths:external-module-path:") {
            reset_module_path_to_workspace(&app, &config_state, &root, module_id)?;
            repaired += 1;
        } else if issue_id == "paths:convert-internal-absolute" {
            let candidates = collect_internal_absolute_path_candidates(&conn, &root)?;
            repaired += apply_path_candidates(&mut conn, &candidates)?;
        }
    }

    let health = run_workspace_health_check_internal(&app, &ws, &conn)?;
    Ok(RepairWorkspaceResult {
        repaired,
        backup_path: Some(backup_path.to_string_lossy().to_string()),
        health,
    })
}

#[tauri::command]
pub async fn repair_workspace_issue(
    app: AppHandle,
    runtime: State<'_, Arc<WorkspaceRuntime>>,
    config_state: State<'_, Mutex<AppConfig>>,
    issue_id: String,
) -> Result<RepairWorkspaceResult, String> {
    repair_workspace_issues(app, runtime, config_state, vec![issue_id]).await
}

fn add_file_to_zip<W: Write + std::io::Seek>(
    zip: &mut ZipWriter<W>,
    source: &Path,
    zip_name: &str,
) -> Result<(), String> {
    let options = SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated)
        .unix_permissions(0o644);
    zip.start_file(zip_name, options)
        .map_err(|e| e.to_string())?;
    let mut file =
        File::open(source).map_err(|e| format!("读取文件失败 {}: {}", source.display(), e))?;
    std::io::copy(&mut file, zip).map_err(|e| format!("写入 zip 失败: {}", e))?;
    Ok(())
}

fn add_dir_to_zip<W: Write + std::io::Seek>(
    zip: &mut ZipWriter<W>,
    source_dir: &Path,
    base_dir: &Path,
    skip_dirs: &HashSet<String>,
) -> Result<(), String> {
    let mut entries: Vec<PathBuf> = fs::read_dir(source_dir)
        .map_err(|e| format!("读取目录失败 {}: {}", source_dir.display(), e))?
        .filter_map(|entry| entry.ok().map(|e| e.path()))
        .collect();
    entries.sort();
    for path in entries {
        let name = path
            .file_name()
            .map(|n| n.to_string_lossy().to_string())
            .unwrap_or_default();
        if path.is_dir() && skip_dirs.contains(&name) {
            continue;
        }
        let rel = path.strip_prefix(base_dir).map_err(|e| e.to_string())?;
        let zip_name = to_zip_name(rel);
        if path.is_dir() {
            let options = SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated)
                .unix_permissions(0o755);
            zip.add_directory(format!("{}/", zip_name.trim_end_matches('/')), options)
                .map_err(|e| e.to_string())?;
            add_dir_to_zip(zip, &path, base_dir, skip_dirs)?;
        } else {
            add_file_to_zip(zip, &path, &zip_name)?;
        }
    }
    Ok(())
}

fn unique_default_export_path(root: &Path, workspace_name: &str) -> PathBuf {
    let export_dir = root.join(EXPORTS_DIR);
    let _ = fs::create_dir_all(&export_dir);
    workspace::mark_path_hidden_if_supported(&export_dir);
    let base = format!(
        "LamberWorkspace-{}-{}.lamber.zip",
        sanitize_file_part(workspace_name),
        today_stamp()
    );
    let mut candidate = export_dir.join(&base);
    if !candidate.exists() {
        return candidate;
    }
    for idx in 1..1000 {
        let name = format!(
            "LamberWorkspace-{}-{}-{}.lamber.zip",
            sanitize_file_part(workspace_name),
            today_stamp(),
            idx
        );
        candidate = export_dir.join(name);
        if !candidate.exists() {
            return candidate;
        }
    }
    export_dir.join(format!(
        "LamberWorkspace-{}-{}.lamber.zip",
        sanitize_file_part(workspace_name),
        now_stamp()
    ))
}

#[tauri::command]
pub async fn export_workspace(
    app: AppHandle,
    runtime: State<'_, Arc<WorkspaceRuntime>>,
    target_path: Option<String>,
    options: WorkspaceExportOptions,
) -> Result<ExportWorkspaceResult, String> {
    let ws = runtime.require_workspace()?;
    let root = PathBuf::from(&ws.workspace_root);
    let include_backups = options.include_backups.unwrap_or(false);
    let include_exports = options.include_exports.unwrap_or(false);
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;

    let health = run_workspace_health_check_internal(&app, &ws, &conn)?;
    let blocking_errors: Vec<_> = health
        .items
        .iter()
        .filter(|item| item.severity == "error")
        .cloned()
        .collect();
    if !blocking_errors.is_empty() && !options.allow_warnings.unwrap_or(false) {
        return Err(format!(
            "工作区健康检查存在严重问题，已停止导出: {}",
            blocking_errors[0].message
        ));
    }
    let warnings: Vec<HealthCheckItem> = health
        .items
        .iter()
        .filter(|item| item.severity == "warning")
        .cloned()
        .collect();
    let external_paths = collect_external_paths(&app, &conn, &root)?;

    let db_backup = create_backup_with_name(
        &root,
        &conn,
        &format!("lamber-before-export-{}.sqlite", now_stamp()),
    )?;
    let archive_path = target_path
        .map(PathBuf::from)
        .unwrap_or_else(|| unique_default_export_path(&root, &ws.workspace_name));
    if archive_path.exists() {
        return Err(format!("目标导出文件已存在: {}", archive_path.display()));
    }
    if let Some(parent) = archive_path.parent() {
        fs::create_dir_all(parent).map_err(|e| format!("创建导出目录失败: {}", e))?;
    }
    let tmp_path = archive_path.with_extension("lamber.zip.tmp");

    let result = (|| -> Result<ExportWorkspaceResult, String> {
        let file = File::create(&tmp_path).map_err(|e| format!("创建临时导出文件失败: {}", e))?;
        let mut zip = ZipWriter::new(file);
        let mut skip_dirs = HashSet::new();
        if !include_backups {
            skip_dirs.insert(BACKUPS_DIR.to_string());
        }
        if !include_exports {
            skip_dirs.insert(EXPORTS_DIR.to_string());
        }

        let manifest = ExportManifest {
            exported_at: Utc::now().to_rfc3339(),
            app_version: env!("CARGO_PKG_VERSION").to_string(),
            workspace_id: ws.workspace_id.clone(),
            workspace_name: ws.workspace_name.clone(),
            workspace_version: ws.manifest.workspace_version,
            include_backups,
            include_exports,
            external_path_count: external_paths.len(),
            warnings: warnings.iter().map(|item| item.message.clone()).collect(),
        };

        for entry in fs::read_dir(&root).map_err(|e| format!("读取工作区目录失败: {}", e))?
        {
            let path = entry.map_err(|e| e.to_string())?.path();
            let name = path
                .file_name()
                .map(|n| n.to_string_lossy().to_string())
                .unwrap_or_default();
            if skip_dirs.contains(&name) || name == workspace::DATABASE_FILE {
                continue;
            }
            if path == tmp_path || path == archive_path {
                continue;
            }
            if path.is_dir() {
                let options = SimpleFileOptions::default()
                    .compression_method(zip::CompressionMethod::Deflated)
                    .unix_permissions(0o755);
                zip.add_directory(format!("{}/", name), options)
                    .map_err(|e| e.to_string())?;
                add_dir_to_zip(&mut zip, &path, &root, &skip_dirs)?;
            } else {
                let rel = path.strip_prefix(&root).map_err(|e| e.to_string())?;
                add_file_to_zip(&mut zip, &path, &to_zip_name(rel))?;
            }
        }

        add_file_to_zip(&mut zip, &db_backup, workspace::DATABASE_FILE)?;
        let options = SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated)
            .unix_permissions(0o644);
        zip.start_file("export-manifest.json", options)
            .map_err(|e| e.to_string())?;
        zip.write_all(
            serde_json::to_string_pretty(&manifest)
                .map_err(|e| e.to_string())?
                .as_bytes(),
        )
        .map_err(|e| e.to_string())?;
        zip.finish()
            .map_err(|e| format!("完成 zip 写入失败: {}", e))?;
        fs::rename(&tmp_path, &archive_path).map_err(|e| format!("移动导出文件失败: {}", e))?;

        Ok(ExportWorkspaceResult {
            archive_path: archive_path.to_string_lossy().to_string(),
            database_backup_path: db_backup.to_string_lossy().to_string(),
            manifest,
            warnings,
        })
    })();

    if result.is_err() {
        let _ = fs::remove_file(&tmp_path);
    }
    result
}

fn safe_zip_entry_name(name: &str) -> Result<PathBuf, String> {
    let path = Path::new(name);
    if path.is_absolute() {
        return Err(format!("压缩包包含绝对路径: {}", name));
    }
    for component in path.components() {
        match component {
            Component::Normal(_) => {}
            Component::CurDir => {}
            _ => return Err(format!("压缩包包含不安全路径: {}", name)),
        }
    }
    Ok(path.to_path_buf())
}

fn validate_archive_internal(zip_path: &Path) -> Result<ArchiveValidationResult, String> {
    let file = File::open(zip_path).map_err(|e| format!("打开压缩包失败: {}", e))?;
    let mut archive = ZipArchive::new(file).map_err(|e| format!("读取压缩包失败: {}", e))?;
    let mut names = Vec::new();
    let mut errors = Vec::new();
    for idx in 0..archive.len() {
        let file = archive.by_index(idx).map_err(|e| e.to_string())?;
        let name = file.name().replace('\\', "/");
        if let Err(err) = safe_zip_entry_name(&name) {
            errors.push(err);
        }
        names.push(name);
    }

    let has_at = |prefix: &str, target: &str| {
        let full = if prefix.is_empty() {
            target.to_string()
        } else {
            format!("{}/{}", prefix.trim_end_matches('/'), target)
        };
        names.iter().any(|name| name == &full)
    };

    let mut root_prefix = String::new();
    if !has_at("", workspace::MANIFEST_FILE) || !has_at("", workspace::DATABASE_FILE) {
        let mut top_levels = HashSet::new();
        for name in &names {
            if let Some(first) = name.split('/').next() {
                if !first.is_empty() {
                    top_levels.insert(first.to_string());
                }
            }
        }
        if top_levels.len() == 1 {
            root_prefix = top_levels.into_iter().next().unwrap_or_default();
        }
    }

    if !has_at(&root_prefix, workspace::MANIFEST_FILE) {
        errors.push("压缩包缺少 .lamber.workspace.json".to_string());
    }
    if !has_at(&root_prefix, workspace::DATABASE_FILE) {
        errors.push("压缩包缺少 .lamber.sqlite".to_string());
    }

    let mut workspace_id = None;
    let mut workspace_name = None;
    let mut workspace_version = None;
    if errors.is_empty() {
        let manifest_name = if root_prefix.is_empty() {
            workspace::MANIFEST_FILE.to_string()
        } else {
            format!("{}/{}", root_prefix, workspace::MANIFEST_FILE)
        };
        let mut manifest_file = archive.by_name(&manifest_name).map_err(|e| e.to_string())?;
        let mut raw = String::new();
        manifest_file
            .read_to_string(&mut raw)
            .map_err(|e| e.to_string())?;
        let manifest: WorkspaceManifest =
            serde_json::from_str(&raw).map_err(|e| format!("manifest 解析失败: {}", e))?;
        if manifest.workspace_version > workspace::WORKSPACE_VERSION {
            errors.push(format!(
                "workspaceVersion {} 高于当前支持版本 {}",
                manifest.workspace_version,
                workspace::WORKSPACE_VERSION
            ));
        }
        workspace_id = Some(manifest.workspace_id);
        workspace_name = Some(manifest.name);
        workspace_version = Some(manifest.workspace_version);
    }

    Ok(ArchiveValidationResult {
        valid: errors.is_empty(),
        root_prefix: if root_prefix.is_empty() {
            None
        } else {
            Some(root_prefix)
        },
        workspace_id,
        workspace_name,
        workspace_version,
        errors,
        warnings: Vec::new(),
    })
}

#[tauri::command]
pub async fn validate_workspace_archive(
    zip_path: String,
) -> Result<ArchiveValidationResult, String> {
    validate_archive_internal(Path::new(&zip_path))
}

fn is_dir_empty(path: &Path) -> Result<bool, String> {
    if !path.exists() {
        return Ok(true);
    }
    Ok(fs::read_dir(path)
        .map_err(|e| e.to_string())?
        .next()
        .is_none())
}

fn resolve_import_destination(
    target_dir: &Path,
    validation: &ArchiveValidationResult,
    options: &ImportWorkspaceOptions,
) -> PathBuf {
    let name = if let Some(name) = options
        .destination_name
        .as_ref()
        .filter(|s| !s.trim().is_empty())
    {
        name.as_str()
    } else {
        validation
            .workspace_name
            .as_deref()
            .filter(|name| !name.trim().is_empty())
            .unwrap_or("LamberWorkspace")
    };
    target_dir.join(sanitize_file_part(name))
}

fn extract_archive_to(
    zip_path: &Path,
    temp_dir: &Path,
    root_prefix: Option<&str>,
) -> Result<(), String> {
    let file = File::open(zip_path).map_err(|e| format!("打开压缩包失败: {}", e))?;
    let mut archive = ZipArchive::new(file).map_err(|e| format!("读取压缩包失败: {}", e))?;
    fs::create_dir_all(temp_dir).map_err(|e| format!("创建导入临时目录失败: {}", e))?;
    let temp_canon =
        fs::canonicalize(temp_dir).map_err(|e| format!("解析导入临时目录失败: {}", e))?;
    for idx in 0..archive.len() {
        let mut file = archive.by_index(idx).map_err(|e| e.to_string())?;
        let raw_name = file.name().replace('\\', "/");
        let safe_name = safe_zip_entry_name(&raw_name)?;
        let stripped = if let Some(prefix) = root_prefix {
            let prefix_path = Path::new(prefix);
            match safe_name.strip_prefix(prefix_path) {
                Ok(v) => v.to_path_buf(),
                Err(_) => continue,
            }
        } else {
            safe_name
        };
        if stripped.as_os_str().is_empty() {
            continue;
        }
        let out_path = temp_dir.join(&stripped);
        let parent = out_path
            .parent()
            .ok_or_else(|| "非法压缩包路径".to_string())?;
        fs::create_dir_all(parent).map_err(|e| format!("创建解压目录失败: {}", e))?;
        let parent_canon =
            fs::canonicalize(parent).map_err(|e| format!("解析解压目录失败: {}", e))?;
        if !parent_canon.starts_with(&temp_canon) {
            return Err(format!("压缩包路径穿越风险: {}", raw_name));
        }
        if file.is_dir() || raw_name.ends_with('/') {
            fs::create_dir_all(&out_path).map_err(|e| format!("创建目录失败: {}", e))?;
        } else {
            let mut out =
                File::create(&out_path).map_err(|e| format!("创建解压文件失败: {}", e))?;
            std::io::copy(&mut file, &mut out).map_err(|e| format!("写入解压文件失败: {}", e))?;
        }
    }
    Ok(())
}

#[tauri::command]
pub async fn import_workspace(
    app: AppHandle,
    runtime: State<'_, Arc<WorkspaceRuntime>>,
    zip_path: String,
    target_dir: String,
    open_after_import: Option<Value>,
    conflict_strategy: Option<String>,
    destination_name: Option<String>,
) -> Result<ImportWorkspaceResult, String> {
    let open_after_import_value = open_after_import.as_ref();
    let options = ImportWorkspaceOptions {
        open_after_import: value_bool_arg(open_after_import_value, "openAfterImport"),
        conflict_strategy: conflict_strategy
            .or_else(|| {
                value_string_field(
                    open_after_import_value,
                    "conflictStrategy",
                    "conflict_strategy",
                )
            })
            .or_else(|| Some("rename".to_string())),
        destination_name: destination_name.or_else(|| {
            value_string_field(
                open_after_import_value,
                "destinationName",
                "destination_name",
            )
        }),
    };
    let zip_path = PathBuf::from(zip_path);
    let target_dir = PathBuf::from(target_dir);
    let validation = validate_archive_internal(&zip_path)?;
    if !validation.valid {
        return Err(format!(
            "工作区压缩包无效: {}",
            validation.errors.join("; ")
        ));
    }

    let mut final_root = resolve_import_destination(&target_dir, &validation, &options);
    let mut moved_existing: Option<(PathBuf, PathBuf)> = None;
    if final_root.exists() && !is_dir_empty(&final_root)? {
        match options.conflict_strategy.as_deref().unwrap_or("cancel") {
            "rename" => {
                let base = final_root
                    .file_name()
                    .map(|n| n.to_string_lossy().to_string())
                    .unwrap_or_else(|| "LamberWorkspace".to_string());
                let parent = final_root
                    .parent()
                    .map(Path::to_path_buf)
                    .unwrap_or_else(|| target_dir.clone());
                for idx in 1..1000 {
                    let candidate = parent.join(format!("{}-{}", base, idx));
                    if !candidate.exists() {
                        final_root = candidate;
                        break;
                    }
                }
            }
            "overwrite" => {
                let backup_existing = final_root.with_file_name(format!(
                    "{}.before-import-{}",
                    final_root
                        .file_name()
                        .map(|n| n.to_string_lossy().to_string())
                        .unwrap_or_else(|| "workspace".to_string()),
                    now_stamp()
                ));
                let backup_existing_for_rollback = backup_existing.clone();
                fs::rename(&final_root, backup_existing)
                    .map_err(|e| format!("覆盖导入前移动原目录失败: {}", e))?;
                moved_existing = Some((backup_existing_for_rollback, final_root.clone()));
            }
            _ => return Err("目标目录已存在且非空，已取消导入".to_string()),
        }
    }

    let parent = final_root
        .parent()
        .map(Path::to_path_buf)
        .unwrap_or_else(|| target_dir.clone());
    fs::create_dir_all(&parent).map_err(|e| format!("创建导入目标父目录失败: {}", e))?;
    let temp_dir = parent.join(format!(".lamber-import-tmp-{}", now_stamp()));
    let result = (|| -> Result<ImportWorkspaceResult, String> {
        extract_archive_to(&zip_path, &temp_dir, validation.root_prefix.as_deref())?;
        if !workspace::manifest_path(&temp_dir).exists() || !workspace::db_path(&temp_dir).exists()
        {
            return Err("解压后的工作区缺少 manifest 或数据库".to_string());
        }
        validate_sqlite_file(&workspace::db_path(&temp_dir))?;
        fs::rename(&temp_dir, &final_root).map_err(|e| format!("移动导入目录失败: {}", e))?;
        workspace::ensure_workspace_system_entries_hidden(&final_root);

        let manifest_raw =
            fs::read_to_string(workspace::manifest_path(&final_root)).map_err(|e| e.to_string())?;
        let manifest: WorkspaceManifest =
            serde_json::from_str(&manifest_raw).map_err(|e| e.to_string())?;
        let imported_ws = CurrentWorkspace {
            workspace_root: final_root.to_string_lossy().to_string(),
            workspace_name: manifest.name.clone(),
            workspace_id: manifest.workspace_id.clone(),
            manifest,
        };

        let opened = options.open_after_import.unwrap_or(false);
        let workspace = if opened {
            Some(workspace::open_workspace_internal(
                &app,
                &runtime,
                &final_root,
            )?)
        } else {
            workspace::update_recent(&app, &imported_ws, false)?;
            None
        };

        let warnings = if opened {
            let current = runtime.require_workspace()?;
            let db = runtime.require_db()?;
            let conn = db.lock().map_err(|e| e.to_string())?;
            run_workspace_health_check_internal(&app, &current, &conn)?
                .items
                .into_iter()
                .filter(|item| item.severity == "warning")
                .collect()
        } else {
            Vec::new()
        };

        Ok(ImportWorkspaceResult {
            workspace_root: final_root.to_string_lossy().to_string(),
            opened,
            workspace,
            warnings,
        })
    })();
    if result.is_err() {
        let _ = fs::remove_dir_all(&temp_dir);
        if let Some((backup_existing, original)) = moved_existing {
            if backup_existing.exists() && !original.exists() {
                let _ = fs::rename(backup_existing, original);
            }
        }
    }
    result
}

#[tauri::command]
pub async fn reveal_in_file_manager(path: String) -> Result<(), String> {
    let path = PathBuf::from(path);
    if !path.exists() {
        return Err("路径不存在".to_string());
    }
    #[cfg(target_os = "windows")]
    {
        if path.is_file() {
            std::process::Command::new("explorer.exe")
                .arg(format!("/select,{}", path.to_string_lossy()))
                .spawn()
                .map_err(|e| e.to_string())?;
        } else {
            std::process::Command::new("explorer.exe")
                .arg(&path)
                .spawn()
                .map_err(|e| e.to_string())?;
        }
    }
    #[cfg(any(target_os = "macos", target_os = "ios"))]
    {
        let mut cmd = std::process::Command::new("open");
        if path.is_file() {
            cmd.arg("-R");
        }
        cmd.arg(&path).spawn().map_err(|e| e.to_string())?;
    }
    #[cfg(not(any(target_os = "windows", target_os = "macos", target_os = "ios")))]
    {
        let open_path = if path.is_file() {
            path.parent().unwrap_or(&path).to_path_buf()
        } else {
            path
        };
        std::process::Command::new("xdg-open")
            .arg(open_path)
            .spawn()
            .map_err(|e| e.to_string())?;
    }
    Ok(())
}
