use super::models::ProjectFile;
use super::repository::SqliteProjectFileRepository;
use super::service::ProjectFileService;
use std::sync::Arc;
use tauri::{AppHandle, State};

fn file_service_from_workspace(
    runtime: &crate::workspace::WorkspaceRuntime,
) -> Result<ProjectFileService, String> {
    let repo = Arc::new(SqliteProjectFileRepository::new(runtime.require_db()?));
    Ok(ProjectFileService::new(repo))
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
    file_service_from_workspace(&runtime)?.bind_project_folder(&project_id, &folder_path, force_mode)
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
    file_service_from_workspace(&runtime)?.scan_project_folder(&project_id, recursive.unwrap_or(false))
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
    file_service_from_workspace(&runtime)?.add_project_file(&app, &workspace.workspace_root, &project_id, &src_path, &storage_mode)
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
    super::assets::get_template_asset_path_internal(&app, &conn, &workspace.workspace_root, &asset_id)
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
    super::assets::cleanup_orphan_template_assets_internal(&app, &conn, &workspace.workspace_root, &project_id)
}

#[tauri::command]
pub async fn get_project_setting(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
    key: String,
) -> Result<Option<String>, String> {
    use crate::benefit::repository::ProjectRepository;
    let project_repo = crate::benefit::repository::SqliteProjectRepository::new(runtime.require_db()?);
    project_repo.get_project_setting(&project_id, &key)
}

#[tauri::command]
pub async fn save_project_setting(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
    key: String,
    value: String,
) -> Result<(), String> {
    use crate::benefit::repository::ProjectRepository;
    let project_repo = crate::benefit::repository::SqliteProjectRepository::new(runtime.require_db()?);
    project_repo.save_project_setting(&project_id, &key, &value)
}

