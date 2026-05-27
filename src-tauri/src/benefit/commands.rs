use super::models::{BenefitAnalysisScheme, BenefitAnalysisSnapshot, IctInput, IctResult, Project};
use super::repository::SqliteProjectRepository;
use super::service::ProjectService;
use std::sync::Arc;
use tauri::State;

fn service_from_workspace(
    runtime: &crate::workspace::WorkspaceRuntime,
) -> Result<ProjectService, String> {
    let conn = runtime.require_db()?;
    Ok(ProjectService::new(Box::new(SqliteProjectRepository::new(conn))))
}

#[tauri::command]
pub async fn get_projects(runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>) -> Result<Vec<Project>, String> {
    let service = service_from_workspace(&runtime)?;
    service.get_projects()
}

#[tauri::command]
pub async fn get_project(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    id: String,
) -> Result<Option<Project>, String> {
    let service = service_from_workspace(&runtime)?;
    service.get_project(&id)
}

#[tauri::command]
pub async fn create_project(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    name: String,
    customer_name: String,
) -> Result<Project, String> {
    let service = service_from_workspace(&runtime)?;
    service.create_project(name, customer_name)
}

#[tauri::command]
pub async fn update_project(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project: Project,
) -> Result<Project, String> {
    let service = service_from_workspace(&runtime)?;
    service.update_project(project)
}

#[tauri::command]
pub async fn delete_project(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    id: String,
) -> Result<(), String> {
    let service = service_from_workspace(&runtime)?;
    service.delete_project(&id)
}

#[tauri::command]
pub async fn delete_benefit_scheme(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
    scheme_id: String,
) -> Result<Project, String> {
    let service = service_from_workspace(&runtime)?;
    service.delete_benefit_scheme(project_id, scheme_id)
}

#[tauri::command]
pub async fn get_schemes(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
) -> Result<Vec<BenefitAnalysisScheme>, String> {
    let service = service_from_workspace(&runtime)?;
    service.get_schemes(&project_id)
}

#[tauri::command]
pub async fn get_snapshots(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    scheme_id: String,
) -> Result<Vec<BenefitAnalysisSnapshot>, String> {
    let service = service_from_workspace(&runtime)?;
    service.get_snapshots(&scheme_id)
}

#[tauri::command]
pub async fn save_benefit_scheme(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    project_id: String,
    scheme_id_opt: Option<String>,
    scheme_name: String,
    input_params: IctInput,
    output_metrics: IctResult,
    is_save_as_new: bool,
) -> Result<Project, String> {
    let service = service_from_workspace(&runtime)?;
    service.save_benefit_scheme(
        project_id,
        scheme_id_opt,
        scheme_name,
        input_params,
        output_metrics,
        is_save_as_new,
    )
}
