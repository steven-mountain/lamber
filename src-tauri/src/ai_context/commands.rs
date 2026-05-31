use super::dto::{
    AiProjectContextBundle, AiProjectContextRequest, AiTemplateAssetImageInput,
    AiTemplateAssetRequest, AiWorkspaceProjectIndexItem,
};
use std::sync::Arc;
use tauri::{AppHandle, State};

#[tauri::command]
pub async fn build_ai_project_context(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    request: AiProjectContextRequest,
) -> Result<AiProjectContextBundle, String> {
    let workspace = runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    super::service::build_ai_project_context(&conn, &workspace.workspace_root, request)
}

#[tauri::command]
pub async fn list_ai_workspace_projects(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
) -> Result<Vec<AiWorkspaceProjectIndexItem>, String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    super::service::list_ai_workspace_projects(&conn)
}

#[tauri::command]
pub async fn load_ai_template_asset(
    app: AppHandle,
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    request: AiTemplateAssetRequest,
) -> Result<AiTemplateAssetImageInput, String> {
    let workspace = runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    super::service::load_ai_template_asset(
        &app,
        &conn,
        &workspace.workspace_root,
        request.project_id.trim(),
        request.asset_id.trim(),
    )
}
