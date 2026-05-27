use crate::config_manager::{AppConfig, ConfigManager};
use crate::db;
use chrono::Utc;
use serde::{Deserialize, Serialize};
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::{Arc, Mutex, RwLock};
use tauri::{AppHandle, State};
use tauri_plugin_dialog::DialogExt;

const MANIFEST_FILE: &str = "lamber.workspace.json";
const DATABASE_FILE: &str = "lamber.sqlite";
const WORKSPACE_VERSION: i32 = 1;

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WorkspaceManifest {
    pub app: String,
    pub workspace_version: i32,
    pub workspace_id: String,
    pub name: String,
    pub created_at: String,
    pub last_opened_at: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RecentWorkspace {
    pub path: String,
    pub name: String,
    pub workspace_id: String,
    pub last_opened_at: String,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CurrentWorkspace {
    pub workspace_root: String,
    pub workspace_name: String,
    pub workspace_id: String,
    pub manifest: WorkspaceManifest,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WorkspaceState {
    pub current_workspace: Option<CurrentWorkspace>,
    pub recent_workspaces: Vec<RecentWorkspace>,
    pub is_workspace_ready: bool,
    pub startup_error: Option<WorkspaceErrorPayload>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WorkspacePathStatus {
    pub status: String,
    pub message: Option<String>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WorkspaceErrorPayload {
    pub code: String,
    pub message: String,
}

impl WorkspaceErrorPayload {
    fn new(code: &str, message: impl Into<String>) -> Self {
        Self {
            code: code.to_string(),
            message: message.into(),
        }
    }

    fn as_string(&self) -> String {
        serde_json::to_string(self).unwrap_or_else(|_| format!("{}::{}", self.code, self.message))
    }
}

pub struct WorkspaceRuntime {
    current: RwLock<Option<CurrentWorkspace>>,
    db: RwLock<Option<Arc<Mutex<rusqlite::Connection>>>>,
    startup_error: RwLock<Option<WorkspaceErrorPayload>>,
}

impl WorkspaceRuntime {
    pub fn new() -> Self {
        Self {
            current: RwLock::new(None),
            db: RwLock::new(None),
            startup_error: RwLock::new(None),
        }
    }

    pub fn get_current_workspace(&self) -> Option<CurrentWorkspace> {
        self.current.read().ok().and_then(|w| w.clone())
    }

    pub fn require_workspace(&self) -> Result<CurrentWorkspace, String> {
        self.get_current_workspace().ok_or_else(|| {
            WorkspaceErrorPayload::new("NotReady", "请先新建或打开 Lamber 工作区").as_string()
        })
    }

    pub fn require_db(&self) -> Result<Arc<Mutex<rusqlite::Connection>>, String> {
        self.db
            .read()
            .map_err(|e| e.to_string())?
            .clone()
            .ok_or_else(|| {
                WorkspaceErrorPayload::new("NotReady", "当前未打开工作区，无法执行数据库操作").as_string()
            })
    }

    pub fn switch_workspace(&self, workspace: CurrentWorkspace, conn: rusqlite::Connection) -> Result<(), String> {
        *self.current.write().map_err(|e| e.to_string())? = Some(workspace);
        *self.db.write().map_err(|e| e.to_string())? = Some(Arc::new(Mutex::new(conn)));
        *self.startup_error.write().map_err(|e| e.to_string())? = None;
        Ok(())
    }

    pub fn clear_workspace(&self) {
        if let Ok(mut current) = self.current.write() {
            *current = None;
        }
        if let Ok(mut db) = self.db.write() {
            *db = None;
        }
    }

    fn set_startup_error(&self, err: WorkspaceErrorPayload) {
        self.clear_workspace();
        if let Ok(mut startup_error) = self.startup_error.write() {
            *startup_error = Some(err);
        }
    }

    fn startup_error(&self) -> Option<WorkspaceErrorPayload> {
        self.startup_error.read().ok().and_then(|e| e.clone())
    }
}

fn manifest_path(root: &Path) -> PathBuf {
    root.join(MANIFEST_FILE)
}

fn db_path(root: &Path) -> PathBuf {
    root.join(DATABASE_FILE)
}

fn workspace_error(code: &str, message: impl Into<String>) -> String {
    WorkspaceErrorPayload::new(code, message).as_string()
}

fn read_manifest(root: &Path) -> Result<WorkspaceManifest, String> {
    let content = fs::read_to_string(manifest_path(root))
        .map_err(|e| workspace_error("InvalidManifest", format!("读取工作区标识失败: {}", e)))?;
    let manifest: WorkspaceManifest = serde_json::from_str(&content)
        .map_err(|e| workspace_error("InvalidManifest", format!("工作区标识文件格式无效: {}", e)))?;
    if manifest.app != "Lamber" {
        return Err(workspace_error("InvalidManifest", "该目录不是 Lamber 工作区"));
    }
    if manifest.workspace_version > WORKSPACE_VERSION {
        return Err(workspace_error(
            "UnsupportedVersion",
            format!("工作区版本 {} 高于当前应用支持版本 {}", manifest.workspace_version, WORKSPACE_VERSION),
        ));
    }
    Ok(manifest)
}

fn write_manifest(root: &Path, manifest: &WorkspaceManifest) -> Result<(), String> {
    let content = serde_json::to_string_pretty(manifest)
        .map_err(|e| workspace_error("InvalidManifest", format!("序列化工作区标识失败: {}", e)))?;
    fs::write(manifest_path(root), content)
        .map_err(|e| workspace_error("PermissionDenied", format!("写入工作区标识失败: {}", e)))
}

fn ensure_writable_dir(root: &Path) -> Result<(), String> {
    if !root.exists() {
        fs::create_dir_all(root)
            .map_err(|e| workspace_error("PermissionDenied", format!("创建目录失败: {}", e)))?;
    }
    if !root.is_dir() {
        return Err(workspace_error("PermissionDenied", "选择的路径不是文件夹"));
    }
    let probe = root.join(".lamber_write_probe");
    fs::write(&probe, b"ok")
        .map_err(|e| workspace_error("PermissionDenied", format!("目录不可写: {}", e)))?;
    let _ = fs::remove_file(probe);
    Ok(())
}

fn inspect_path(root: &Path) -> Result<WorkspacePathStatus, String> {
    if !root.exists() {
        return Ok(WorkspacePathStatus {
            status: "emptyOrInitializable".to_string(),
            message: None,
        });
    }
    if !root.is_dir() {
        return Err(workspace_error("PermissionDenied", "选择的路径不是文件夹"));
    }
    if manifest_path(root).exists() {
        return Ok(WorkspacePathStatus {
            status: "workspace".to_string(),
            message: None,
        });
    }
    if db_path(root).exists() {
        return Ok(WorkspacePathStatus {
            status: "legacySuspected".to_string(),
            message: Some("检测到 lamber.sqlite 但缺少 lamber.workspace.json，疑似旧版数据目录。本阶段不会覆盖或迁移。".to_string()),
        });
    }
    let is_empty = fs::read_dir(root)
        .map_err(|e| workspace_error("PermissionDenied", format!("读取目录失败: {}", e)))?
        .next()
        .is_none();
    Ok(WorkspacePathStatus {
        status: if is_empty { "emptyOrInitializable" } else { "nonEmptyNonWorkspace" }.to_string(),
        message: None,
    })
}

fn update_recent(app: &AppHandle, workspace: &CurrentWorkspace) -> Result<(), String> {
    let manager = ConfigManager::new(app);
    let mut config = manager.load();
    let recent = RecentWorkspace {
        path: workspace.workspace_root.clone(),
        name: workspace.workspace_name.clone(),
        workspace_id: workspace.workspace_id.clone(),
        last_opened_at: workspace.manifest.last_opened_at.clone(),
    };
    config.recent_workspaces.retain(|item| item.path != recent.path);
    config.recent_workspaces.insert(0, recent);
    config.recent_workspaces.truncate(10);
    config.last_opened_workspace_path = Some(workspace.workspace_root.clone());
    manager.save(&config)
}

fn open_workspace_internal(app: &AppHandle, runtime: &WorkspaceRuntime, root: &Path) -> Result<CurrentWorkspace, String> {
    if !manifest_path(root).exists() {
        if db_path(root).exists() {
            return Err(workspace_error("InvalidManifest", "疑似旧版 Lamber 数据目录：存在 lamber.sqlite 但缺少 lamber.workspace.json"));
        }
        return Err(workspace_error("InvalidManifest", "该目录不是 Lamber 工作区"));
    }
    let mut manifest = read_manifest(root)?;
    manifest.last_opened_at = Utc::now().to_rfc3339();
    write_manifest(root, &manifest)?;

    let db_file = db_path(root);
    let conn = db::init_db(&db_file).map_err(|e| {
        let msg = e.to_string();
        if msg.to_ascii_lowercase().contains("database disk image is malformed") {
            workspace_error("DatabaseCorrupted", format!("工作区数据库损坏: {}", msg))
        } else {
            workspace_error("DatabaseOpenFailed", format!("打开工作区数据库失败: {}", msg))
        }
    })?;
    ensure_workspace_root_registered(&conn, root)?;

    let workspace = CurrentWorkspace {
        workspace_root: root.to_string_lossy().to_string(),
        workspace_name: manifest.name.clone(),
        workspace_id: manifest.workspace_id.clone(),
        manifest,
    };
    runtime.switch_workspace(workspace.clone(), conn)?;
    update_recent(app, &workspace)?;
    Ok(workspace)
}

fn ensure_workspace_root_registered(conn: &rusqlite::Connection, root: &Path) -> Result<(), String> {
    let root_path = root.to_string_lossy().to_string();
    
    // Check if there is an existing workspace root record (id starting with "workspace_root_")
    let existing_ws_root: Option<(String, String)> = match conn.query_row(
        "SELECT id, root_path FROM project_roots WHERE id LIKE 'workspace_root_%' LIMIT 1",
        [],
        |row| Ok((row.get::<_, String>(0)?, row.get::<_, String>(1)?)),
    ) {
        Ok(val) => Some(val),
        Err(rusqlite::Error::QueryReturnedNoRows) => None,
        Err(e) => return Err(workspace_error("DatabaseOpenFailed", format!("检查工作区根目录失败: {}", e))),
    };

    let now = Utc::now().to_rfc3339();
    
    if let Some((id, old_path)) = existing_ws_root {
        if old_path != root_path {
            // Path has changed (workspace was moved), update the root_path!
            conn.execute(
                "UPDATE project_roots SET root_path = ?1, updated_at = ?2 WHERE id = ?3",
                rusqlite::params![root_path, now, id],
            )
            .map_err(|e| workspace_error("DatabaseOpenFailed", format!("更新工作区根目录失败: {}", e)))?;
        }
    } else {
        // No workspace root registered yet, check if root_path exists under any ID
        let existing_count: i64 = conn
            .query_row(
                "SELECT COUNT(*) FROM project_roots WHERE root_path = ?1",
                [&root_path],
                |row| row.get(0),
            )
            .map_err(|e| workspace_error("DatabaseOpenFailed", format!("检查工作区根目录失败: {}", e)))?;

        if existing_count == 0 {
            let any_default: i64 = conn
                .query_row("SELECT COUNT(*) FROM project_roots WHERE is_default = 1", [], |row| row.get(0))
                .map_err(|e| workspace_error("DatabaseOpenFailed", format!("检查默认根目录失败: {}", e)))?;
            conn.execute(
                "INSERT INTO project_roots (id, name, root_path, root_alias, is_default, created_at, updated_at)
                 VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7)",
                rusqlite::params![
                    format!("workspace_root_{}", uuid::Uuid::new_v4()),
                    "当前工作区",
                    root_path,
                    Option::<String>::None,
                    if any_default == 0 { 1 } else { 0 },
                    now,
                    now,
                ],
            )
            .map_err(|e| workspace_error("DatabaseOpenFailed", format!("注册工作区根目录失败: {}", e)))?;
        }
    }

    Ok(())
}

pub fn try_restore_last_workspace(app: &AppHandle, runtime: &WorkspaceRuntime, config: &AppConfig) {
    if let Some(path) = &config.last_opened_workspace_path {
        let root = PathBuf::from(path);
        if let Err(err) = open_workspace_internal(app, runtime, &root) {
            runtime.set_startup_error(WorkspaceErrorPayload::new(
                "NotReady",
                format!("自动恢复上次工作区失败，请重新选择。详情: {}", err),
            ));
        }
    }
}

#[tauri::command]
pub async fn get_workspace_state(
    app: AppHandle,
    runtime: State<'_, Arc<WorkspaceRuntime>>,
) -> Result<WorkspaceState, String> {
    let config = ConfigManager::new(&app).load();
    let current = runtime.get_current_workspace();
    Ok(WorkspaceState {
        is_workspace_ready: current.is_some(),
        current_workspace: current,
        recent_workspaces: config.recent_workspaces,
        startup_error: runtime.startup_error(),
    })
}

#[tauri::command]
pub async fn inspect_workspace_path(path: String) -> Result<WorkspacePathStatus, String> {
    inspect_path(Path::new(&path))
}

#[tauri::command]
pub async fn select_workspace_folder(app: AppHandle) -> Result<Option<String>, String> {
    let folder = app.dialog().file().blocking_pick_folder();
    Ok(folder.map(|f| f.to_string()))
}

#[tauri::command]
pub async fn create_workspace(
    app: AppHandle,
    runtime: State<'_, Arc<WorkspaceRuntime>>,
    path: String,
    name: Option<String>,
    allow_non_empty: Option<bool>,
) -> Result<CurrentWorkspace, String> {
    let root = PathBuf::from(path);
    ensure_writable_dir(&root)?;
    let status = inspect_path(&root)?;
    match status.status.as_str() {
        "workspace" => return open_workspace_internal(&app, &runtime, &root),
        "legacySuspected" => {
            return Err(workspace_error(
                "InvalidManifest",
                status.message.unwrap_or_else(|| "疑似旧版数据目录，本阶段不会覆盖".to_string()),
            ));
        }
        "nonEmptyNonWorkspace" if !allow_non_empty.unwrap_or(false) => {
            return Err(workspace_error(
                "InvalidManifest",
                "该目录非空且不是 Lamber 工作区，需要用户确认后才能初始化",
            ));
        }
        _ => {}
    }

    fs::create_dir_all(root.join("projects"))
        .map_err(|e| workspace_error("PermissionDenied", format!("创建 projects 目录失败: {}", e)))?;
    fs::create_dir_all(root.join("backups"))
        .map_err(|e| workspace_error("PermissionDenied", format!("创建 backups 目录失败: {}", e)))?;
    fs::create_dir_all(root.join("exports"))
        .map_err(|e| workspace_error("PermissionDenied", format!("创建 exports 目录失败: {}", e)))?;

    let now = Utc::now().to_rfc3339();
    let workspace_name = name
        .filter(|n| !n.trim().is_empty())
        .unwrap_or_else(|| root.file_name().map(|n| n.to_string_lossy().to_string()).unwrap_or_else(|| "Lamber Workspace".to_string()));
    let manifest = WorkspaceManifest {
        app: "Lamber".to_string(),
        workspace_version: WORKSPACE_VERSION,
        workspace_id: uuid::Uuid::new_v4().to_string(),
        name: workspace_name,
        created_at: now.clone(),
        last_opened_at: now,
    };
    write_manifest(&root, &manifest)?;
    open_workspace_internal(&app, &runtime, &root)
}

#[tauri::command]
pub async fn open_workspace(
    app: AppHandle,
    runtime: State<'_, Arc<WorkspaceRuntime>>,
    path: String,
) -> Result<CurrentWorkspace, String> {
    open_workspace_internal(&app, &runtime, &PathBuf::from(path))
}

#[tauri::command]
pub async fn clear_workspace(runtime: State<'_, Arc<WorkspaceRuntime>>) -> Result<(), String> {
    runtime.clear_workspace();
    Ok(())
}

use crate::benefit::models::Project;

// Checks if `path` is inside `workspace_root`
pub fn is_inside_workspace(workspace_root: &Path, path: &Path) -> bool {
    if let (Ok(ws_canon), Ok(p_canon)) = (fs::canonicalize(workspace_root), fs::canonicalize(path)) {
        p_canon.starts_with(ws_canon)
    } else {
        let ws_str = workspace_root.to_string_lossy().to_string().replace("\\", "/");
        let p_str = path.to_string_lossy().to_string().replace("\\", "/");
        p_str.starts_with(&ws_str)
    }
}

// Converts absolute path to relative path under workspace root, or returns absolute path if outside
pub fn to_relative_workspace_path(workspace_root: &Path, absolute_path: &Path) -> String {
    let ws_abs = if workspace_root.is_absolute() {
        workspace_root.to_path_buf()
    } else {
        fs::canonicalize(workspace_root).unwrap_or_else(|_| workspace_root.to_path_buf())
    };
    
    let target_abs = if absolute_path.is_absolute() {
        absolute_path.to_path_buf()
    } else {
        fs::canonicalize(absolute_path).unwrap_or_else(|_| absolute_path.to_path_buf())
    };

    if let Ok(rel) = target_abs.strip_prefix(&ws_abs) {
        rel.to_string_lossy().to_string().replace("\\", "/")
    } else {
        let ws_str = ws_abs.to_string_lossy().to_string().replace("\\", "/");
        let target_str = target_abs.to_string_lossy().to_string().replace("\\", "/");
        if target_str.starts_with(&ws_str) {
            let mut rel = &target_str[ws_str.len()..];
            rel = rel.trim_start_matches('/');
            rel.to_string()
        } else {
            absolute_path.to_string_lossy().to_string().replace("\\", "/")
        }
    }
}

// Resolves relative path to absolute path under workspace root. If already absolute, returns it.
pub fn resolve_workspace_path(workspace_root: &Path, relative_path: &str) -> PathBuf {
    let p = Path::new(relative_path);
    if p.is_absolute() {
        p.to_path_buf()
    } else {
        workspace_root.join(relative_path)
    }
}

// Sanitize folder name by replacing invalid characters
pub fn sanitize_folder_name(name: &str) -> String {
    let trimmed = name.trim();
    if trimmed.is_empty() {
        return "unnamed_project".to_string();
    }
    
    let mut sanitized = String::new();
    for c in trimmed.chars() {
        if c.is_alphanumeric() || c == '_' || c == '-' || c == '.' || c == '(' || c == ')' || c == '[' || c == ']' {
            sanitized.push(c);
        } else {
            sanitized.push('_');
        }
    }
    
    let mut cleaned = sanitized.replace("__", "_");
    while cleaned.contains("__") {
        cleaned = cleaned.replace("__", "_");
    }
    cleaned = cleaned.trim_matches('_').to_string();
    if cleaned.is_empty() {
        cleaned = "project".to_string();
    }
    
    cleaned
}

// Ensures all required directories exist for a project: assets/, documents/, analyses/
pub fn ensure_project_dirs(workspace_root: &Path, folder_name: &str) -> Result<(), String> {
    let project_dir = workspace_root.join("projects").join(folder_name);
    fs::create_dir_all(&project_dir).map_err(|e| format!("无法创建项目目录: {}", e))?;
    fs::create_dir_all(project_dir.join("assets")).map_err(|e| format!("无法创建 assets 目录: {}", e))?;
    fs::create_dir_all(project_dir.join("documents")).map_err(|e| format!("无法创建 documents 目录: {}", e))?;
    fs::create_dir_all(project_dir.join("analyses")).map_err(|e| format!("无法创建 analyses 目录: {}", e))?;
    Ok(())
}

pub fn normalize_project_paths(workspace_root: &Path, project: &mut Project) {
    if let Some(folder_path) = &project.folder_path {
        if folder_path.trim().is_empty() {
            project.folder_path = None;
            project.linked_folder_type = Some("none".to_string());
            project.linked_folder_relative_path = None;
            project.linked_folder_external_path = None;
            project.folder_name = None;
            project.relative_path = None;
            return;
        }

        let path = Path::new(folder_path);
        if path.is_absolute() {
            if is_inside_workspace(workspace_root, path) {
                let rel = to_relative_workspace_path(workspace_root, path);
                project.folder_path = Some(rel.clone());
                project.linked_folder_type = Some("internal".to_string());
                project.linked_folder_relative_path = Some(rel.clone());
                project.linked_folder_external_path = None;
                project.relative_path = Some(rel.clone());
                if let Some(name) = Path::new(&rel).file_name() {
                    project.folder_name = Some(name.to_string_lossy().to_string());
                }
            } else {
                project.linked_folder_type = Some("external".to_string());
                project.linked_folder_relative_path = None;
                project.linked_folder_external_path = Some(folder_path.clone());
                project.relative_path = Some(folder_path.clone());
                if let Some(name) = path.file_name() {
                    project.folder_name = Some(name.to_string_lossy().to_string());
                }
            }
        } else {
            project.linked_folder_type = Some("internal".to_string());
            project.linked_folder_relative_path = Some(folder_path.clone());
            project.linked_folder_external_path = None;
            project.relative_path = Some(folder_path.clone());
            if let Some(name) = path.file_name() {
                project.folder_name = Some(name.to_string_lossy().to_string());
            }
        }
    } else {
        project.linked_folder_type = Some("none".to_string());
        project.linked_folder_relative_path = None;
        project.linked_folder_external_path = None;
        project.folder_name = None;
        project.relative_path = None;
    }
}
