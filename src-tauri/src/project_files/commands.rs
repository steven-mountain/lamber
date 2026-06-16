use super::models::ProjectFile;
use super::repository::SqliteProjectFileRepository;
use super::service::ProjectFileService;
use std::path::Path;
use std::sync::{Arc, Mutex};
use tauri::{AppHandle, State};

fn file_service_from_workspace(
    runtime: &crate::workspace::WorkspaceRuntime,
) -> Result<ProjectFileService, String> {
    let ws = runtime.require_workspace()?;
    let repo = Arc::new(SqliteProjectFileRepository::new(runtime.require_db()?));
    Ok(ProjectFileService::new(
        repo,
        std::path::PathBuf::from(ws.workspace_root),
    ))
}

#[tauri::command]
pub async fn get_project_files(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
) -> Result<Vec<ProjectFile>, String> {
    file_service_from_workspace(&runtime)?.get_project_files(&project_id)
}

#[tauri::command]
pub async fn bind_project_folder(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
    folder_path: String,
    force_mode: Option<String>,
) -> Result<(), String> {
    let file_service = file_service_from_workspace(&runtime)?;
    file_service.bind_project_folder(&project_id, &folder_path, force_mode)?;
    if let Ok(files) = file_service.get_project_files(&project_id) {
        let _ = auto_import_excel_if_needed(&runtime, &project_id, &files);
    }
    Ok(())
}

#[tauri::command]
pub async fn create_project_folder(
    parent_path: String,
    folder_name: String,
) -> Result<String, String> {
    ProjectFileService::create_project_folder(&parent_path, &folder_name)
}

#[tauri::command]
pub async fn unbind_project_folder(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
) -> Result<(), String> {
    file_service_from_workspace(&runtime)?.unbind_project_folder(&project_id)
}

#[tauri::command]
pub async fn scan_project_folder(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
    recursive: Option<bool>,
) -> Result<Vec<ProjectFile>, String> {
    let file_service = file_service_from_workspace(&runtime)?;
    let files = file_service.scan_project_folder(&project_id, recursive.unwrap_or(false))?;
    let _ = auto_import_excel_if_needed(&runtime, &project_id, &files);
    Ok(files)
}

#[tauri::command]
pub async fn add_project_file(
    app: AppHandle,
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
    src_path: String,
    storage_mode: String,
) -> Result<ProjectFile, String> {
    let workspace = runtime.require_workspace()?;
    file_service_from_workspace(&runtime)?.add_project_file(
        &app,
        &workspace.workspace_root,
        &project_id,
        &src_path,
        &storage_mode,
    )
}

#[tauri::command]
pub async fn remove_project_file_record(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
    file_id: String,
) -> Result<(), String> {
    file_service_from_workspace(&runtime)?.remove_project_file_record(&project_id, &file_id)
}

#[tauri::command]
pub async fn delete_managed_project_file(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
    file_id: String,
) -> Result<(), String> {
    file_service_from_workspace(&runtime)?.delete_managed_project_file(&project_id, &file_id)
}

#[tauri::command]
pub async fn mark_main_document(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
    file_id: Option<String>,
) -> Result<(), String> {
    file_service_from_workspace(&runtime)?.mark_main_document(&project_id, file_id.as_deref())
}

#[tauri::command]
pub async fn mark_main_budget_file(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
    file_id: Option<String>,
) -> Result<(), String> {
    file_service_from_workspace(&runtime)?.mark_main_budget_file(&project_id, file_id.as_deref())
}

#[tauri::command]
pub async fn open_project_folder(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
) -> Result<(), String> {
    file_service_from_workspace(&runtime)?.open_project_folder(&project_id)
}

#[tauri::command]
pub async fn open_project_file(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    file_id: String,
) -> Result<(), String> {
    file_service_from_workspace(&runtime)?.open_project_file(&file_id)
}

#[tauri::command]
pub async fn reveal_project_file(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    file_id: String,
) -> Result<(), String> {
    file_service_from_workspace(&runtime)?.reveal_project_file(&file_id)
}

use tauri_plugin_dialog::DialogExt;

#[tauri::command]
pub async fn select_local_folder(app: AppHandle) -> Result<Option<String>, String> {
    let folder = app.dialog().file().blocking_pick_folder();
    Ok(folder.map(|f| f.to_string()))
}

#[tauri::command]
pub async fn select_local_file(
    app: AppHandle,
    title: String,
    extensions: Option<Vec<String>>,
) -> Result<Option<String>, String> {
    let mut dialog = app.dialog().file().set_title(&title);
    if let Some(ref exts) = extensions {
        let exts_strs: Vec<&str> = exts.iter().map(|s| s.as_str()).collect();
        dialog = dialog.add_filter("Files", &exts_strs);
    }
    let file = dialog.blocking_pick_file();
    Ok(file.map(|f| f.to_string()))
}

#[tauri::command]
pub async fn save_template_asset(
    app: AppHandle,
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
    template_name: String,
    asset_type: String,
    usage: Option<String>,
    original_file_name: Option<String>,
    base64_data: String,
    width: Option<i32>,
    height: Option<i32>,
) -> Result<String, String> {
    let workspace = runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    super::assets::save_template_asset_internal(
        &app,
        &conn,
        &workspace.workspace_root,
        &project_id,
        &template_name,
        &asset_type,
        usage.as_deref(),
        original_file_name.as_deref(),
        &base64_data,
        width,
        height,
    )
}

#[tauri::command]
pub async fn get_template_asset_path(
    app: AppHandle,
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    asset_id: String,
) -> Result<String, String> {
    let workspace = runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    super::assets::get_template_asset_path_internal(
        &app,
        &conn,
        &workspace.workspace_root,
        &asset_id,
    )
}

#[tauri::command]
pub async fn delete_template_asset(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    asset_id: String,
) -> Result<(), String> {
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    super::assets::delete_template_asset_internal(&conn, &asset_id)
}

#[tauri::command]
pub async fn cleanup_orphan_template_assets(
    app: AppHandle,
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
) -> Result<(usize, Vec<String>), String> {
    let workspace = runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    super::assets::cleanup_orphan_template_assets_internal(
        &app,
        &conn,
        &workspace.workspace_root,
        &project_id,
    )
}

#[tauri::command]
pub async fn get_project_setting(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
    key: String,
) -> Result<Option<String>, String> {
    use crate::benefit::repository::ProjectRepository;
    let project_repo =
        crate::benefit::repository::SqliteProjectRepository::new(runtime.require_db()?);
    project_repo.get_project_setting(&project_id, &key)
}

#[tauri::command]
pub async fn save_project_setting(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
    key: String,
    value: String,
) -> Result<(), String> {
    if key == "ai_compute_quote::active" {
        return Err("LegacyAiComputeSettingReadOnly".to_string());
    }
    use crate::benefit::repository::ProjectRepository;
    let project_repo =
        crate::benefit::repository::SqliteProjectRepository::new(runtime.require_db()?);
    project_repo.save_project_setting(&project_id, &key, &value)
}

pub fn auto_import_excel_if_needed(
    runtime: &crate::workspace::WorkspaceRuntime,
    project_id: &str,
    files: &[ProjectFile],
) -> Result<(), String> {
    let db = runtime.require_db()?;
    let ws = runtime.require_workspace()?;
    auto_import_excel_if_needed_with_context(db, Path::new(&ws.workspace_root), project_id, files)
}

pub(crate) fn auto_import_excel_if_needed_with_context(
    db: Arc<Mutex<rusqlite::Connection>>,
    workspace_root: &Path,
    project_id: &str,
    files: &[ProjectFile],
) -> Result<(), String> {
    let project_repo = crate::benefit::repository::SqliteProjectRepository::new(db.clone());
    let project_service = crate::benefit::service::ProjectService::new(Box::new(project_repo));

    // 1. Check if the project already has any schemes
    let schemes = project_service.get_schemes(project_id)?;
    if !schemes.is_empty() {
        return Ok(());
    }

    // 2. Find files starting with "效益分析表" and ending with ".xlsx" or ".xls" (case-insensitive)
    let mut matching_files: Vec<&ProjectFile> = files
        .iter()
        .filter(|f| {
            let starts_with_pattern = f.file_name.starts_with("效益分析表");
            let ext = f.extension.to_lowercase();
            starts_with_pattern && (ext == "xlsx" || ext == "xls")
        })
        .collect();

    if matching_files.is_empty() {
        return Ok(());
    }

    // 3. Sort by modified_at descending (newest first)
    matching_files.sort_by(|a, b| b.modified_at.cmp(&a.modified_at));
    let target_file = matching_files[0];

    // 4. Resolve the file path against workspace root
    let file_path_buf = std::path::PathBuf::from(&target_file.file_path);
    let resolved_path = if !file_path_buf.is_absolute() {
        crate::workspace::resolve_workspace_path(workspace_root, &target_file.file_path)
    } else {
        file_path_buf
    };

    if !resolved_path.exists() {
        return Err(format!("匹配到的测算文件不存在: {:?}", resolved_path));
    }

    // 5. Parse and auto-import
    let parsed_data = crate::benefit::excel::parse_benefit_excel_internal(&resolved_path)?;
    crate::benefit::excel::auto_import_excel_calculation(
        project_id,
        &target_file.file_name,
        parsed_data,
        &project_service,
    )?;

    Ok(())
}
