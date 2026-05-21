use super::models::{BenefitAnalysisScheme, BenefitAnalysisSnapshot, IctInput, IctResult, Project};
use super::service::ProjectService;
use std::sync::Arc;
use tauri::State;

#[tauri::command]
pub async fn get_projects(service: State<'_, Arc<ProjectService>>) -> Result<Vec<Project>, String> {
    service.get_projects()
}

#[tauri::command]
pub async fn get_project(
    service: State<'_, Arc<ProjectService>>,
    id: String,
) -> Result<Option<Project>, String> {
    service.get_project(&id)
}

#[tauri::command]
pub async fn create_project(
    service: State<'_, Arc<ProjectService>>,
    name: String,
    customer_name: String,
) -> Result<Project, String> {
    service.create_project(name, customer_name)
}

#[tauri::command]
pub async fn update_project(
    service: State<'_, Arc<ProjectService>>,
    project: Project,
) -> Result<Project, String> {
    service.update_project(project)
}

#[tauri::command]
pub async fn delete_project(
    service: State<'_, Arc<ProjectService>>,
    id: String,
) -> Result<(), String> {
    service.delete_project(&id)
}

#[tauri::command]
pub async fn delete_benefit_scheme(
    service: State<'_, Arc<ProjectService>>,
    project_id: String,
    scheme_id: String,
) -> Result<Project, String> {
    service.delete_benefit_scheme(project_id, scheme_id)
}

#[tauri::command]
pub async fn get_schemes(
    service: State<'_, Arc<ProjectService>>,
    project_id: String,
) -> Result<Vec<BenefitAnalysisScheme>, String> {
    service.get_schemes(&project_id)
}

#[tauri::command]
pub async fn get_snapshots(
    service: State<'_, Arc<ProjectService>>,
    scheme_id: String,
) -> Result<Vec<BenefitAnalysisSnapshot>, String> {
    service.get_snapshots(&scheme_id)
}

#[tauri::command]
pub async fn save_benefit_scheme(
    service: State<'_, Arc<ProjectService>>,
    project_id: String,
    scheme_id_opt: Option<String>,
    scheme_name: String,
    input_params: IctInput,
    output_metrics: IctResult,
    is_save_as_new: bool,
) -> Result<Project, String> {
    service.save_benefit_scheme(
        project_id,
        scheme_id_opt,
        scheme_name,
        input_params,
        output_metrics,
        is_save_as_new,
    )
}
