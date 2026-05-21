use super::models::ProjectFile;
use super::service::ProjectFileService;
use std::sync::Arc;
use tauri::{AppHandle, State};

#[tauri::command]
pub async fn get_project_files(
    state: State<'_, Arc<ProjectFileService>>,
    project_id: String,
) -> Result<Vec<ProjectFile>, String> {
    state.get_project_files(&project_id)
}

#[tauri::command]
pub async fn bind_project_folder(
    state: State<'_, Arc<ProjectFileService>>,
    project_id: String,
    folder_path: String,
) -> Result<(), String> {
    state.bind_project_folder(&project_id, &folder_path)
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
    state: State<'_, Arc<ProjectFileService>>,
    project_id: String,
) -> Result<(), String> {
    state.unbind_project_folder(&project_id)
}

#[tauri::command]
pub async fn scan_project_folder(
    state: State<'_, Arc<ProjectFileService>>,
    project_id: String,
    recursive: Option<bool>,
) -> Result<Vec<ProjectFile>, String> {
    state.scan_project_folder(&project_id, recursive.unwrap_or(false))
}

#[tauri::command]
pub async fn add_project_file(
    app: AppHandle,
    state: State<'_, Arc<ProjectFileService>>,
    project_id: String,
    src_path: String,
    storage_mode: String,
) -> Result<ProjectFile, String> {
    state.add_project_file(&app, &project_id, &src_path, &storage_mode)
}

#[tauri::command]
pub async fn remove_project_file_record(
    state: State<'_, Arc<ProjectFileService>>,
    project_id: String,
    file_id: String,
) -> Result<(), String> {
    state.remove_project_file_record(&project_id, &file_id)
}

#[tauri::command]
pub async fn delete_managed_project_file(
    state: State<'_, Arc<ProjectFileService>>,
    project_id: String,
    file_id: String,
) -> Result<(), String> {
    state.delete_managed_project_file(&project_id, &file_id)
}

#[tauri::command]
pub async fn mark_main_document(
    state: State<'_, Arc<ProjectFileService>>,
    project_id: String,
    file_id: Option<String>,
) -> Result<(), String> {
    state.mark_main_document(&project_id, file_id.as_deref())
}

#[tauri::command]
pub async fn mark_main_budget_file(
    state: State<'_, Arc<ProjectFileService>>,
    project_id: String,
    file_id: Option<String>,
) -> Result<(), String> {
    state.mark_main_budget_file(&project_id, file_id.as_deref())
}

#[tauri::command]
pub async fn open_project_folder(
    state: State<'_, Arc<ProjectFileService>>,
    project_id: String,
) -> Result<(), String> {
    state.open_project_folder(&project_id)
}

#[tauri::command]
pub async fn open_project_file(
    state: State<'_, Arc<ProjectFileService>>,
    file_id: String,
) -> Result<(), String> {
    state.open_project_file(&file_id)
}

#[tauri::command]
pub async fn reveal_project_file(
    state: State<'_, Arc<ProjectFileService>>,
    file_id: String,
) -> Result<(), String> {
    state.reveal_project_file(&file_id)
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
