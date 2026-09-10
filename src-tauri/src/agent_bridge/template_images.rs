//! User-operated image replacement. This IPC command is not an agent tool.
use serde::Deserialize;
use std::sync::Arc;
use tauri::Manager;

#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ReplaceImageRequest {
    session_id: String,
    workspace_id: String,
    project_id: String,
    template_name: String,
    asset_id: String,
    file_name: String,
    data_url: String,
    width: i32,
    height: i32,
}

#[tauri::command]
pub async fn ai_replace_template_image(app: tauri::AppHandle, request: ReplaceImageRequest) -> Result<String, String> {
    let binding = super::open_session_store(&app)?.binding(&request.session_id)?.ok_or("会话已失效")?;
    let runtime = app.state::<Arc<crate::workspace::WorkspaceRuntime>>();
    let (workspace, db) = runtime.require_context()?;
    let cwd = std::fs::canonicalize(&workspace.workspace_root).map_err(|e| e.to_string())?;
    if binding.workspace_id != request.workspace_id || workspace.workspace_id != request.workspace_id
        || binding.cwd != cwd || binding.project_id.as_deref() != Some(request.project_id.as_str()) {
        return Err("工作区或会话绑定已变更，请重新选择图片".into());
    }
    let conn = db.lock().map_err(|e| e.to_string())?;
    crate::project_files::assets::validate_replacement_image(&request.data_url, request.width, request.height)?;
    crate::project_files::assets::replace_demand_image_internal(
        &conn, &workspace.workspace_root, &request.project_id, &request.template_name, &request.asset_id,
        |tx, usage| crate::project_files::assets::save_template_asset_internal(
            &app, tx, &workspace.workspace_root, &request.project_id, &request.template_name,
            "image", Some(usage), Some(&request.file_name), &request.data_url, Some(request.width), Some(request.height)),
    )
}
