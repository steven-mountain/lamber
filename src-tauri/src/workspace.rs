use crate::config_manager::{AppConfig, ConfigManager};
use crate::db;
use chrono::Utc;
use serde::{Deserialize, Serialize};
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::{Arc, Mutex, RwLock};
use tauri::{AppHandle, Emitter, State};
use tauri_plugin_dialog::DialogExt;

pub(crate) const MANIFEST_FILE: &str = ".lamber.workspace.json";
pub(crate) const DATABASE_FILE: &str = ".lamber.sqlite";
pub(crate) const WORKSPACE_VERSION: i32 = 1;

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
                WorkspaceErrorPayload::new("NotReady", "当前未打开工作区，无法执行数据库操作")
                    .as_string()
            })
    }

    pub fn switch_workspace(
        &self,
        workspace: CurrentWorkspace,
        conn: rusqlite::Connection,
    ) -> Result<(), String> {
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

    pub fn close_database(&self) -> Result<(), String> {
        *self.db.write().map_err(|e| e.to_string())? = None;
        Ok(())
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

pub(crate) fn manifest_path(root: &Path) -> PathBuf {
    root.join(MANIFEST_FILE)
}

pub(crate) fn db_path(root: &Path) -> PathBuf {
    root.join(DATABASE_FILE)
}

pub(crate) fn mark_path_hidden_if_supported(path: &Path) {
    #[cfg(windows)]
    {
        if !path.exists() {
            return;
        }
        if let Err(err) = std::process::Command::new("attrib")
            .arg("+h")
            .arg(path)
            .status()
        {
            eprintln!(
                "Failed to mark path hidden on Windows: {} ({})",
                path.display(),
                err
            );
        }
    }

    #[cfg(not(windows))]
    {
        let _ = path;
    }
}

pub(crate) fn ensure_workspace_system_entries_hidden(root: &Path) {
    for name in [
        MANIFEST_FILE,
        DATABASE_FILE,
        ".backups",
        ".exports",
        ".projects",
    ] {
        mark_path_hidden_if_supported(&root.join(name));
    }
}

fn workspace_error(code: &str, message: impl Into<String>) -> String {
    WorkspaceErrorPayload::new(code, message).as_string()
}

fn read_manifest(root: &Path) -> Result<WorkspaceManifest, String> {
    let content = fs::read_to_string(manifest_path(root))
        .map_err(|e| workspace_error("InvalidManifest", format!("读取工作区标识失败: {}", e)))?;
    let manifest: WorkspaceManifest = serde_json::from_str(&content).map_err(|e| {
        workspace_error("InvalidManifest", format!("工作区标识文件格式无效: {}", e))
    })?;
    if manifest.app != "Lamber" {
        return Err(workspace_error(
            "InvalidManifest",
            "该目录不是 Lamber 工作区",
        ));
    }
    if manifest.workspace_version > WORKSPACE_VERSION {
        return Err(workspace_error(
            "UnsupportedVersion",
            format!(
                "工作区版本 {} 高于当前应用支持版本 {}",
                manifest.workspace_version, WORKSPACE_VERSION
            ),
        ));
    }
    Ok(manifest)
}

fn write_manifest(root: &Path, manifest: &WorkspaceManifest) -> Result<(), String> {
    let content = serde_json::to_string_pretty(manifest)
        .map_err(|e| workspace_error("InvalidManifest", format!("序列化工作区标识失败: {}", e)))?;
    let path = manifest_path(root);
    fs::write(&path, content)
        .map_err(|e| workspace_error("PermissionDenied", format!("写入工作区标识失败: {}", e)))?;
    mark_path_hidden_if_supported(&path);
    Ok(())
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

fn get_import_candidates(root: &Path) -> Result<Vec<String>, String> {
    if !root.exists() || !root.is_dir() {
        return Ok(Vec::new());
    }
    let mut candidates = Vec::new();
    let entries = fs::read_dir(root)
        .map_err(|e| workspace_error("PermissionDenied", format!("读取目录失败: {}", e)))?;

    for entry in entries {
        let entry = match entry {
            Ok(e) => e,
            Err(_) => continue,
        };
        let path = entry.path();
        if path.is_dir() {
            if let Some(name) = path.file_name() {
                let name_str = name.to_string_lossy().to_string();
                if name_str.starts_with('.') {
                    continue;
                }
                if is_reserved_workspace_entry_name(&name_str)
                    || matches!(
                        name_str.as_str(),
                        "node_modules"
                            | "target"
                            | "dist"
                            | "build"
                            | ".vscode"
                            | ".idea"
                            | "__pycache__"
                    )
                {
                    continue;
                }
                candidates.push(name_str);
            }
        }
    }
    candidates.sort();
    Ok(candidates)
}

fn migrate_legacy_workspace_files(root: &Path) {
    // 1. Rename lamber.workspace.json -> .lamber.workspace.json
    let old_manifest = root.join("lamber.workspace.json");
    let new_manifest = root.join(".lamber.workspace.json");
    if old_manifest.exists() && !new_manifest.exists() {
        let _ = fs::rename(&old_manifest, &new_manifest);
    }

    // 2. Rename lamber.sqlite -> .lamber.sqlite
    let old_db = root.join("lamber.sqlite");
    let new_db = root.join(".lamber.sqlite");
    if old_db.exists() && !new_db.exists() {
        let _ = fs::rename(&old_db, &new_db);
    }

    // 3. Rename backups -> .backups
    let old_backups = root.join("backups");
    let new_backups = root.join(".backups");
    if old_backups.exists() && old_backups.is_dir() && !new_backups.exists() {
        let _ = fs::rename(&old_backups, &new_backups);
    }

    // 4. Rename exports -> .exports
    let old_exports = root.join("exports");
    let new_exports = root.join(".exports");
    if old_exports.exists() && old_exports.is_dir() && !new_exports.exists() {
        let _ = fs::rename(&old_exports, &new_exports);
    }
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

    // 自动迁移可见的遗留系统文件/目录为隐藏版
    migrate_legacy_workspace_files(root);
    ensure_workspace_system_entries_hidden(root);

    if manifest_path(root).exists() {
        return Ok(WorkspacePathStatus {
            status: "workspace".to_string(),
            message: None,
        });
    }
    if db_path(root).exists() {
        return Ok(WorkspacePathStatus {
            status: "legacySuspected".to_string(),
            message: Some("检测到 .lamber.sqlite 但缺少 .lamber.workspace.json，疑似旧版数据目录。本阶段不会覆盖或迁移。".to_string()),
        });
    }

    let candidates = get_import_candidates(root)?;
    if !candidates.is_empty() {
        return Ok(WorkspacePathStatus {
            status: "importablePlainDirectory".to_string(),
            message: Some(serde_json::to_string(&candidates).unwrap_or_default()),
        });
    }

    let is_empty = fs::read_dir(root)
        .map_err(|e| workspace_error("PermissionDenied", format!("读取目录失败: {}", e)))?
        .next()
        .is_none();
    Ok(WorkspacePathStatus {
        status: if is_empty {
            "emptyOrInitializable"
        } else {
            "nonEmptyNonWorkspace"
        }
        .to_string(),
        message: None,
    })
}

pub(crate) fn update_recent(
    app: &AppHandle,
    workspace: &CurrentWorkspace,
    set_last_opened: bool,
) -> Result<(), String> {
    let manager = ConfigManager::new(app);
    let mut config = manager.load();
    let recent = RecentWorkspace {
        path: workspace.workspace_root.clone(),
        name: workspace.workspace_name.clone(),
        workspace_id: workspace.workspace_id.clone(),
        last_opened_at: workspace.manifest.last_opened_at.clone(),
    };
    config
        .recent_workspaces
        .retain(|item| item.path != recent.path);
    config.recent_workspaces.insert(0, recent);
    config.recent_workspaces.truncate(10);
    if set_last_opened {
        config.last_opened_workspace_path = Some(workspace.workspace_root.clone());
    }
    manager.save(&config)
}

fn workspace_path_key(path: &str) -> String {
    let normalized = path
        .trim()
        .trim_end_matches(|c| c == '\\' || c == '/')
        .replace('\\', "/");
    #[cfg(windows)]
    {
        normalized.to_ascii_lowercase()
    }
    #[cfg(not(windows))]
    {
        normalized
    }
}

fn workspace_paths_match(a: &str, b: &str) -> bool {
    if workspace_path_key(a) == workspace_path_key(b) {
        return true;
    }
    match (fs::canonicalize(a), fs::canonicalize(b)) {
        (Ok(a_path), Ok(b_path)) => a_path == b_path,
        _ => false,
    }
}

pub(crate) fn open_workspace_internal(
    app: &AppHandle,
    runtime: &WorkspaceRuntime,
    root: &Path,
) -> Result<CurrentWorkspace, String> {
    // 自动迁移可见的遗留系统文件/目录为隐藏版
    migrate_legacy_workspace_files(root);
    ensure_workspace_system_entries_hidden(root);

    if !manifest_path(root).exists() {
        if db_path(root).exists() {
            return Err(workspace_error(
                "InvalidManifest",
                "疑似旧版 Lamber 数据目录：存在 .lamber.sqlite 但缺少 .lamber.workspace.json",
            ));
        }
        return Err(workspace_error(
            "InvalidManifest",
            "该目录不是 Lamber 工作区",
        ));
    }
    let mut manifest = read_manifest(root)?;
    manifest.last_opened_at = Utc::now().to_rfc3339();
    write_manifest(root, &manifest)?;

    let db_file = db_path(root);
    let conn = db::init_db(&db_file).map_err(|e| {
        let msg = e.to_string();
        if msg
            .to_ascii_lowercase()
            .contains("database disk image is malformed")
        {
            workspace_error("DatabaseCorrupted", format!("工作区数据库损坏: {}", msg))
        } else {
            workspace_error(
                "DatabaseOpenFailed",
                format!("打开工作区数据库失败: {}", msg),
            )
        }
    })?;
    ensure_workspace_root_registered(&conn, root)?;
    if let Err(err) = crate::workspace_maintenance::ensure_daily_workspace_backup(root, &conn) {
        eprintln!("Workspace auto backup failed: {}", err);
    }
    ensure_workspace_system_entries_hidden(root);

    let workspace = CurrentWorkspace {
        workspace_root: root.to_string_lossy().to_string(),
        workspace_name: manifest.name.clone(),
        workspace_id: manifest.workspace_id.clone(),
        manifest,
    };
    runtime.switch_workspace(workspace.clone(), conn)?;
    update_recent(app, &workspace, true)?;
    Ok(workspace)
}

fn ensure_workspace_root_registered(
    conn: &rusqlite::Connection,
    root: &Path,
) -> Result<(), String> {
    let root_path = root.to_string_lossy().to_string();

    // Check if there is an existing workspace root record (id starting with "workspace_root_")
    let existing_ws_root: Option<(String, String)> = match conn.query_row(
        "SELECT id, root_path FROM project_roots WHERE id LIKE 'workspace_root_%' LIMIT 1",
        [],
        |row| Ok((row.get::<_, String>(0)?, row.get::<_, String>(1)?)),
    ) {
        Ok(val) => Some(val),
        Err(rusqlite::Error::QueryReturnedNoRows) => None,
        Err(e) => {
            return Err(workspace_error(
                "DatabaseOpenFailed",
                format!("检查工作区根目录失败: {}", e),
            ))
        }
    };

    let now = Utc::now().to_rfc3339();

    if let Some((id, old_path)) = existing_ws_root {
        if old_path != root_path {
            // Path has changed (workspace was moved), update the root_path!
            conn.execute(
                "UPDATE project_roots SET root_path = ?1, updated_at = ?2 WHERE id = ?3",
                rusqlite::params![root_path, now, id],
            )
            .map_err(|e| {
                workspace_error("DatabaseOpenFailed", format!("更新工作区根目录失败: {}", e))
            })?;
        }
    } else {
        // No workspace root registered yet, check if root_path exists under any ID
        let existing_count: i64 = conn
            .query_row(
                "SELECT COUNT(*) FROM project_roots WHERE root_path = ?1",
                [&root_path],
                |row| row.get(0),
            )
            .map_err(|e| {
                workspace_error("DatabaseOpenFailed", format!("检查工作区根目录失败: {}", e))
            })?;

        if existing_count == 0 {
            let any_default: i64 = conn
                .query_row(
                    "SELECT COUNT(*) FROM project_roots WHERE is_default = 1",
                    [],
                    |row| row.get(0),
                )
                .map_err(|e| {
                    workspace_error("DatabaseOpenFailed", format!("检查默认根目录失败: {}", e))
                })?;
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

pub const WORKSPACE_STATE_CHANGED_EVENT: &str = "lamber-workspace-state-changed";

fn restore_last_workspace_blocking(
    app: &AppHandle,
    runtime: &WorkspaceRuntime,
    config: &AppConfig,
) {
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

pub fn spawn_restore_last_workspace(
    app: AppHandle,
    runtime: Arc<WorkspaceRuntime>,
    config: AppConfig,
) {
    if config.last_opened_workspace_path.is_none() {
        return;
    }

    tauri::async_runtime::spawn_blocking(move || {
        restore_last_workspace_blocking(&app, &runtime, &config);
        if let Err(err) = app.emit(WORKSPACE_STATE_CHANGED_EVENT, ()) {
            eprintln!(
                "Failed to emit workspace state change after startup restore: {}",
                err
            );
        }
    });
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
pub async fn forget_workspace(
    app: AppHandle,
    runtime: State<'_, Arc<WorkspaceRuntime>>,
    path: String,
) -> Result<WorkspaceState, String> {
    let manager = ConfigManager::new(&app);
    let mut config = manager.load();
    let current = runtime.get_current_workspace();
    let is_current = current
        .as_ref()
        .map(|workspace| workspace_paths_match(&workspace.workspace_root, &path))
        .unwrap_or(false);

    config
        .recent_workspaces
        .retain(|item| !workspace_paths_match(&item.path, &path));

    if is_current {
        runtime.clear_workspace();
        config.last_opened_workspace_path = None;
    } else if config
        .last_opened_workspace_path
        .as_ref()
        .map(|last_path| workspace_paths_match(last_path, &path))
        .unwrap_or(false)
    {
        config.last_opened_workspace_path = None;
    }

    manager.save(&config)?;

    let current = runtime.get_current_workspace();
    Ok(WorkspaceState {
        is_workspace_ready: current.is_some(),
        current_workspace: current,
        recent_workspaces: config.recent_workspaces,
        startup_error: runtime.startup_error(),
    })
}

#[tauri::command]
pub async fn close_current_workspace(
    app: AppHandle,
    runtime: State<'_, Arc<WorkspaceRuntime>>,
) -> Result<WorkspaceState, String> {
    runtime.clear_workspace();
    let manager = ConfigManager::new(&app);
    let mut config = manager.load();
    config.last_opened_workspace_path = None;
    manager.save(&config)?;

    Ok(WorkspaceState {
        is_workspace_ready: false,
        current_workspace: None,
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
                status
                    .message
                    .unwrap_or_else(|| "疑似旧版数据目录，本阶段不会覆盖".to_string()),
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

    fs::create_dir_all(root.join(".backups")).map_err(|e| {
        workspace_error("PermissionDenied", format!("创建 backups 目录失败: {}", e))
    })?;
    fs::create_dir_all(root.join(".exports")).map_err(|e| {
        workspace_error("PermissionDenied", format!("创建 exports 目录失败: {}", e))
    })?;

    let now = Utc::now().to_rfc3339();
    let workspace_name = name.filter(|n| !n.trim().is_empty()).unwrap_or_else(|| {
        root.file_name()
            .map(|n| n.to_string_lossy().to_string())
            .unwrap_or_else(|| "Lamber Workspace".to_string())
    });
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

#[derive(Debug, serde::Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct InitializeWorkspaceOptions {
    pub workspace_name: Option<String>,
    pub selected_directories: Vec<String>,
    pub create_project_json: Option<bool>,
    pub create_sub_dirs: Option<bool>,
}

#[tauri::command]
pub async fn initialize_workspace_from_existing_directory(
    app: AppHandle,
    runtime: State<'_, Arc<WorkspaceRuntime>>,
    path: String,
    options: InitializeWorkspaceOptions,
) -> Result<CurrentWorkspace, String> {
    let root = PathBuf::from(&path);
    if !root.exists() {
        return Err(workspace_error("NotFound", "指定的目录不存在"));
    }
    if !root.is_dir() {
        return Err(workspace_error("PermissionDenied", "指定的路径不是文件夹"));
    }

    if manifest_path(&root).exists() && db_path(&root).exists() {
        return Err(workspace_error(
            "AlreadyWorkspace",
            "该目录下已存在 .lamber.workspace.json 且包含数据库",
        ));
    }

    // Clean up broken manifest if database is missing
    if manifest_path(&root).exists() && !db_path(&root).exists() {
        let _ = fs::remove_file(manifest_path(&root));
    }

    if db_path(&root).exists() && !manifest_path(&root).exists() {
        return Err(workspace_error(
            "LegacySuspected",
            "该目录下已存在 .lamber.sqlite，为防覆盖已中止",
        ));
    }

    ensure_writable_dir(&root)?;

    let created_manifest = !manifest_path(&root).exists();
    let created_db = !db_path(&root).exists();

    // Helper block to execute the steps and cleanup on error
    let result = (|| -> Result<CurrentWorkspace, String> {
        // 1. Create standard workspace directories
        fs::create_dir_all(root.join(".backups")).map_err(|e| {
            workspace_error("PermissionDenied", format!("创建 backups 目录失败: {}", e))
        })?;
        fs::create_dir_all(root.join(".exports")).map_err(|e| {
            workspace_error("PermissionDenied", format!("创建 exports 目录失败: {}", e))
        })?;

        // 2. Create and write manifest
        let now = Utc::now().to_rfc3339();
        let workspace_name = options
            .workspace_name
            .as_ref()
            .filter(|n| !n.trim().is_empty())
            .cloned()
            .unwrap_or_else(|| {
                root.file_name()
                    .map(|n| n.to_string_lossy().to_string())
                    .unwrap_or_else(|| "Lamber Workspace".to_string())
            });

        let manifest = WorkspaceManifest {
            app: "Lamber".to_string(),
            workspace_version: WORKSPACE_VERSION,
            workspace_id: uuid::Uuid::new_v4().to_string(),
            name: workspace_name,
            created_at: now.clone(),
            last_opened_at: now,
        };
        write_manifest(&root, &manifest)?;

        // 3. Open workspace to bind connection
        let workspace = open_workspace_internal(&app, &runtime, &root)?;
        let db_conn = runtime.require_db()?;
        let mut conn_guard = db_conn.lock().map_err(|e| e.to_string())?;

        // 4. Wrap database project inserts inside a transaction
        let tx = conn_guard.transaction().map_err(|e| e.to_string())?;

        let candidates = get_import_candidates(&root)?;
        let create_project_json = options.create_project_json.unwrap_or(true);
        let create_sub_dirs = options.create_sub_dirs.unwrap_or(true);

        let mut imported_projects = Vec::new();

        for subdir in candidates {
            if !options.selected_directories.contains(&subdir) {
                continue;
            }

            let project_dir = root.join(&subdir);
            let project_json_path = project_dir.join("project.json");
            let mut project_id = format!("id_{}", uuid::Uuid::new_v4().simple());
            let mut name = subdir.clone();
            let mut created_at = Utc::now().to_rfc3339();
            let mut updated_at = Utc::now().to_rfc3339();
            let mut project_type = "ict".to_string();
            let mut existing_json: Option<serde_json::Value> = None;

            if project_json_path.exists() {
                if let Ok(content) = fs::read_to_string(&project_json_path) {
                    if let Ok(json_val) = serde_json::from_str::<serde_json::Value>(&content) {
                        if let Some(id_str) = json_val.get("projectId").and_then(|v| v.as_str()) {
                            if !id_str.trim().is_empty() {
                                project_id = id_str.to_string();
                            }
                        }
                        if let Some(n_str) = json_val.get("name").and_then(|v| v.as_str()) {
                            if !n_str.trim().is_empty() {
                                name = n_str.to_string();
                            }
                        }
                        if let Some(c_str) = json_val.get("createdAt").and_then(|v| v.as_str()) {
                            if !c_str.trim().is_empty() {
                                created_at = c_str.to_string();
                            }
                        }
                        if let Some(u_str) = json_val.get("updatedAt").and_then(|v| v.as_str()) {
                            if !u_str.trim().is_empty() {
                                updated_at = u_str.to_string();
                            }
                        }
                        if matches!(
                            json_val.get("projectType").and_then(|v| v.as_str()),
                            Some("intelligent_compute")
                        ) {
                            project_type = "intelligent_compute".to_string();
                        }
                        existing_json = Some(json_val);
                    }
                }
            }

            // Check duplicates
            let exists_by_id: bool = tx
                .query_row(
                    "SELECT EXISTS(SELECT 1 FROM projects WHERE id = ?1)",
                    [&project_id],
                    |row| row.get(0),
                )
                .map_err(|e| e.to_string())?;

            let exists_by_rel_path: bool = tx.query_row(
                "SELECT EXISTS(SELECT 1 FROM projects WHERE relative_path = ?1 OR folder_path = ?1)",
                [&subdir],
                |row| row.get(0)
            ).map_err(|e| e.to_string())?;

            let exists_by_name: bool = tx
                .query_row(
                    "SELECT EXISTS(SELECT 1 FROM projects WHERE name = ?1)",
                    [&name],
                    |row| row.get(0),
                )
                .map_err(|e| e.to_string())?;

            if exists_by_id || exists_by_rel_path || exists_by_name {
                continue;
            }

            // project.json
            if create_project_json {
                let mut json_val = existing_json.unwrap_or_else(|| serde_json::json!({}));
                if json_val.get("projectId").is_none() {
                    json_val["projectId"] = serde_json::Value::String(project_id.clone());
                }
                if json_val.get("name").is_none() {
                    json_val["name"] = serde_json::Value::String(name.clone());
                }
                if json_val.get("folderName").is_none() {
                    json_val["folderName"] = serde_json::Value::String(subdir.clone());
                }
                if json_val.get("relativePath").is_none() {
                    json_val["relativePath"] = serde_json::Value::String(subdir.clone());
                }
                if json_val.get("createdAt").is_none() {
                    json_val["createdAt"] = serde_json::Value::String(created_at.clone());
                }
                if json_val.get("updatedAt").is_none() {
                    json_val["updatedAt"] = serde_json::Value::String(updated_at.clone());
                }
                if json_val.get("projectType").is_none() {
                    json_val["projectType"] = serde_json::Value::String(project_type.clone());
                }
                if json_val.get("source").is_none() {
                    json_val["source"] =
                        serde_json::Value::String("importedPlainDirectory".to_string());
                }
                let updated_content = serde_json::to_string_pretty(&json_val)
                    .map_err(|e| format!("序列化 project.json 失败: {}", e))?;
                fs::write(&project_json_path, updated_content)
                    .map_err(|e| format!("写入 project.json 失败: {}", e))?;
            }

            // Subdirectories
            if create_sub_dirs {
                fs::create_dir_all(project_dir.join("assets"))
                    .map_err(|e| format!("无法创建 assets 目录: {}", e))?;
                fs::create_dir_all(project_dir.join("documents"))
                    .map_err(|e| format!("无法创建 documents 目录: {}", e))?;
                fs::create_dir_all(project_dir.join("analyses"))
                    .map_err(|e| format!("无法创建 analyses 目录: {}", e))?;
            }

            // Insert
            tx.execute(
                "INSERT INTO projects (
                    id, name, customer_name, project_type, status, benefit_status, total_revenue_incl, total_cost_incl, project_years, discount_rate, cashflow_model, created_at, updated_at, folder_path, logs, folder_name, relative_path, progress, linked_folder_type, linked_folder_relative_path
                ) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10, ?11, ?12, ?13, ?14, ?15, ?16, ?17, ?18, ?19, ?20)",
                rusqlite::params![
                    project_id,
                    name,
                    "CMCC",
                    project_type,
                    "需求导入",
                    "not_started",
                    0.0,
                    0.0,
                    1,
                    0.055,
                    "model_a",
                    created_at,
                    updated_at,
                    subdir, // folder_path
                    "[]",   // logs
                    subdir, // folder_name
                    subdir, // relative_path
                    0.0,    // progress
                    "internal", // linked_folder_type
                    subdir, // linked_folder_relative_path
                ]
            ).map_err(|e| format!("写入项目数据库失败: {}", e))?;

            imported_projects.push((project_id, subdir));
        }

        tx.commit().map_err(|e| format!("提交事务失败: {}", e))?;
        // Project folder scanning uses repositories backed by the same SQLite mutex.
        // Release the initialization transaction lock before starting scans.
        drop(conn_guard);

        // 5. Scan folders and auto-import calculation in the background.
        // Initialization must return quickly; scanning/parsing can touch large user files.
        if !imported_projects.is_empty() {
            let scan_db = db_conn.clone();
            let scan_root = root.clone();
            let scan_projects = imported_projects;
            tauri::async_runtime::spawn_blocking(move || {
                let file_repo = std::sync::Arc::new(
                    crate::project_files::repository::SqliteProjectFileRepository::new(
                        scan_db.clone(),
                    ),
                );
                let file_service = crate::project_files::service::ProjectFileService::new(
                    file_repo,
                    scan_root.clone(),
                );

                for (p_id, _subdir) in scan_projects {
                    match file_service.scan_project_folder(&p_id, false) {
                        Ok(files) => {
                            let _ = crate::project_files::commands::auto_import_excel_if_needed_with_context(
                                scan_db.clone(),
                                &scan_root,
                                &p_id,
                                &files,
                            );
                        }
                        Err(err) => {
                            eprintln!(
                                "Workspace initialization background scan failed for {}: {}",
                                p_id, err
                            );
                        }
                    }
                }
            });
        }

        Ok(workspace)
    })();

    if let Err(_err) = &result {
        if created_manifest {
            let _ = fs::remove_file(manifest_path(&root));
        }
        if created_db {
            let _ = fs::remove_file(db_path(&root));
        }
        runtime.clear_workspace();
    }

    result
}

#[tauri::command]
pub async fn clear_workspace(runtime: State<'_, Arc<WorkspaceRuntime>>) -> Result<(), String> {
    runtime.clear_workspace();
    Ok(())
}

#[tauri::command]
pub async fn scan_and_import_all_workspace_calculations(
    runtime: State<'_, Arc<WorkspaceRuntime>>,
) -> Result<usize, String> {
    let ws = runtime.require_workspace()?;
    let db_conn = runtime.require_db()?;

    let projects: Vec<String> = {
        let conn_guard = db_conn.lock().map_err(|e| e.to_string())?;
        let mut stmt = conn_guard
            .prepare("SELECT id FROM projects")
            .map_err(|e| e.to_string())?;
        let rows = stmt
            .query_map([], |row| row.get::<_, String>(0))
            .map_err(|e| e.to_string())?;

        let mut list = Vec::new();
        for r in rows {
            if let Ok(id) = r {
                list.push(id);
            }
        }
        list
    };

    let root_path = std::path::PathBuf::from(&ws.workspace_root);
    let file_repo = std::sync::Arc::new(
        crate::project_files::repository::SqliteProjectFileRepository::new(db_conn.clone()),
    );
    let file_service = crate::project_files::service::ProjectFileService::new(file_repo, root_path);

    let mut import_count = 0;
    for p_id in projects {
        let project_repo =
            crate::benefit::repository::SqliteProjectRepository::new(db_conn.clone());
        let project_service = crate::benefit::service::ProjectService::new(Box::new(project_repo));

        let schemes_before = project_service
            .get_schemes(&p_id)
            .map(|s| s.len())
            .unwrap_or(0);
        if schemes_before == 0 {
            if let Ok(files) = file_service.scan_project_folder(&p_id, false) {
                if let Ok(_) = crate::project_files::commands::auto_import_excel_if_needed(
                    &runtime, &p_id, &files,
                ) {
                    let schemes_after = project_service
                        .get_schemes(&p_id)
                        .map(|s| s.len())
                        .unwrap_or(0);
                    if schemes_after > 0 {
                        import_count += 1;
                    }
                }
            }
        }
    }

    Ok(import_count)
}

use crate::benefit::models::Project;

// Checks if `path` is inside `workspace_root`
pub fn is_inside_workspace(workspace_root: &Path, path: &Path) -> bool {
    if let (Ok(ws_canon), Ok(p_canon)) = (fs::canonicalize(workspace_root), fs::canonicalize(path))
    {
        p_canon.starts_with(ws_canon)
    } else {
        let ws_str = workspace_root
            .to_string_lossy()
            .to_string()
            .replace("\\", "/");
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
            absolute_path
                .to_string_lossy()
                .to_string()
                .replace("\\", "/")
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
        if c.is_alphanumeric()
            || c == '_'
            || c == '-'
            || c == '.'
            || c == '('
            || c == ')'
            || c == '['
            || c == ']'
        {
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

pub fn is_reserved_workspace_entry_name(name: &str) -> bool {
    let normalized = name.trim().replace("\\", "/").to_ascii_lowercase();
    if normalized.is_empty() || normalized.contains('/') {
        return true;
    }
    matches!(
        normalized.as_str(),
        ".lamber.workspace.json"
            | ".lamber.sqlite"
            | ".backups"
            | ".exports"
            | ".projects"
            | "backups"
            | "exports"
            | "projects"
    )
}

// Ensures all required directories exist for a project: assets/, documents/, analyses/
pub fn ensure_project_dirs(workspace_root: &Path, folder_name: &str) -> Result<(), String> {
    let project_dir = workspace_root.join(folder_name);
    fs::create_dir_all(&project_dir).map_err(|e| format!("无法创建项目目录: {}", e))?;
    fs::create_dir_all(project_dir.join("assets"))
        .map_err(|e| format!("无法创建 assets 目录: {}", e))?;
    fs::create_dir_all(project_dir.join("documents"))
        .map_err(|e| format!("无法创建 documents 目录: {}", e))?;
    fs::create_dir_all(project_dir.join("analyses"))
        .map_err(|e| format!("无法创建 analyses 目录: {}", e))?;
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
