use super::models::{BenefitAnalysisScheme, BenefitAnalysisSnapshot, IctInput, IctResult, Project};
use super::repository::{SqliteProjectRepository, ProjectRepository};
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
    let ws = runtime.require_workspace()?;
    let service = service_from_workspace(&runtime)?;
    let updated = service.update_project(project)?;

    // If the project has a relative path, update project.json in physical project directory
    if let Some(ref rel_path) = updated.relative_path {
        let ws_root = std::path::Path::new(&ws.workspace_root);
        let project_dir = crate::workspace::resolve_workspace_path(ws_root, rel_path);
        let project_json_path = project_dir.join("project.json");
        if project_json_path.exists() {
            if let Ok(content) = std::fs::read_to_string(&project_json_path) {
                if let Ok(mut json_val) = serde_json::from_str::<serde_json::Value>(&content) {
                    json_val["name"] = serde_json::Value::String(updated.name.clone());
                    json_val["updatedAt"] = serde_json::Value::String(updated.updated_at.clone());
                    if let Ok(updated_content) = serde_json::to_string_pretty(&json_val) {
                        let _ = std::fs::write(project_json_path, updated_content);
                    }
                }
            }
        }
    }

    Ok(updated)
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

#[derive(serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WorkspaceProjectInfo {
    pub project: Project,
    pub directory_exists: bool,
}

#[derive(serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub struct UnregisteredProject {
    pub project_id: String,
    pub name: String,
    pub relative_path: String,
    pub folder_name: String,
}

#[tauri::command]
pub async fn create_project_in_workspace(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    name: String,
    customer_name: String,
) -> Result<Project, String> {
    let ws = runtime.require_workspace()?;
    let conn = runtime.require_db()?;
    
    let name = name.trim().to_string();
    if name.is_empty() {
        return Err("项目名称不能为空".to_string());
    }

    let customer_name = {
        let trimmed = customer_name.trim();
        if trimmed.is_empty() {
            "未知客户".to_string()
        } else {
            trimmed.to_string()
        }
    };

    let folder_name = crate::workspace::sanitize_folder_name(&name);
    let ws_root = std::path::Path::new(&ws.workspace_root);
    let project_dir = ws_root.join("projects").join(&folder_name);

    if project_dir.exists() {
        return Err("项目目录已存在，请更换项目名称".to_string());
    }

    // 1. Create directories
    crate::workspace::ensure_project_dirs(ws_root, &folder_name)?;

    // 2. Write project.json
    let timestamp = chrono::Utc::now().to_rfc3339();
    let project_id = format!("id_{}", uuid::Uuid::new_v4().simple());
    
    #[derive(serde::Serialize)]
    #[serde(rename_all = "camelCase")]
    struct ProjectJson {
        project_id: String,
        name: String,
        relative_path: String,
        created_at: String,
        updated_at: String,
    }

    let relative_path = format!("projects/{}", folder_name);
    let project_json = ProjectJson {
        project_id: project_id.clone(),
        name: name.clone(),
        relative_path: relative_path.clone(),
        created_at: timestamp.clone(),
        updated_at: timestamp.clone(),
    };

    let project_json_str = serde_json::to_string_pretty(&project_json)
        .map_err(|e| {
            let _ = std::fs::remove_dir_all(&project_dir);
            format!("序列化 project.json 失败: {}", e)
        })?;

    if let Err(e) = std::fs::write(project_dir.join("project.json"), project_json_str) {
        let _ = std::fs::remove_dir_all(&project_dir);
        return Err(format!("写入 project.json 失败: {}", e));
    }

    // 3. Save to database
    let mut project = Project {
        id: project_id,
        name,
        customer_name,
        status: "需求导入".to_string(),
        benefit_status: "not_started".to_string(),
        default_scheme_id: None,
        created_at: timestamp.clone(),
        updated_at: timestamp,
        total_revenue_incl: 0.0,
        total_cost_incl: 0.0,
        project_years: 1,
        discount_rate: 0.055,
        cashflow_model: "model_a".to_string(),
        summary_metrics: None,
        folder_path: Some(relative_path.clone()),
        main_document_path: None,
        main_budget_file_path: None,
        note: None,
        logs: vec![],
        folder_name: Some(folder_name.clone()),
        relative_path: Some(relative_path.clone()),
        progress: 0.0,
        deadline: None,
        linked_folder_type: Some("internal".to_string()),
        linked_folder_relative_path: Some(relative_path),
        linked_folder_external_path: None,
    };

    let repo = SqliteProjectRepository::new(conn);
    if let Err(e) = repo.save_project(&project) {
        let _ = std::fs::remove_dir_all(&project_dir);
        return Err(format!("保存项目到数据库失败，已回滚文件夹: {}", e));
    }

    Ok(project)
}

#[tauri::command]
pub async fn list_workspace_projects(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
) -> Result<Vec<WorkspaceProjectInfo>, String> {
    let ws = runtime.require_workspace()?;
    let repo = SqliteProjectRepository::new(runtime.require_db()?);
    let projects = repo.get_projects()?;
    let ws_root = std::path::Path::new(&ws.workspace_root);
    
    let mut list = Vec::new();
    for mut p in projects {
        crate::workspace::normalize_project_paths(ws_root, &mut p);
        
        let directory_exists = if let Some(ref folder_path) = p.folder_path {
            let path = crate::workspace::resolve_workspace_path(ws_root, folder_path);
            path.exists() && path.is_dir()
        } else {
            false
        };
        
        list.push(WorkspaceProjectInfo {
            project: p,
            directory_exists,
        });
    }
    Ok(list)
}

#[tauri::command]
pub async fn inspect_workspace_projects(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
) -> Result<Vec<UnregisteredProject>, String> {
    let ws = runtime.require_workspace()?;
    let conn = runtime.require_db()?;
    let ws_root = std::path::Path::new(&ws.workspace_root);
    let projects_dir = ws_root.join("projects");
    
    let mut unregistered = Vec::new();
    if !projects_dir.exists() || !projects_dir.is_dir() {
        return Ok(unregistered);
    }
    
    let entries = std::fs::read_dir(projects_dir)
        .map_err(|e| format!("无法读取 projects 目录: {}", e))?;
        
    for entry in entries {
        let entry = match entry {
            Ok(e) => e,
            Err(_) => continue,
        };
        let path = entry.path();
        if path.is_dir() {
            let project_json_path = path.join("project.json");
            if project_json_path.exists() {
                if let Ok(content) = std::fs::read_to_string(&project_json_path) {
                    #[derive(serde::Deserialize)]
                    #[serde(rename_all = "camelCase")]
                    struct MiniProjectJson {
                        project_id: String,
                        name: String,
                        relative_path: String,
                    }
                    if let Ok(json) = serde_json::from_str::<MiniProjectJson>(&content) {
                        let exists: bool = conn.lock()
                            .map_err(|e| e.to_string())?
                            .query_row(
                                "SELECT EXISTS(SELECT 1 FROM projects WHERE id = ?1)",
                                [&json.project_id],
                                |row| row.get(0),
                            )
                            .unwrap_or(false);
                            
                        if !exists {
                            let folder_name = path.file_name()
                                .map(|n| n.to_string_lossy().to_string())
                                .unwrap_or_default();
                            unregistered.push(UnregisteredProject {
                                project_id: json.project_id,
                                name: json.name,
                                relative_path: json.relative_path,
                                folder_name,
                            });
                        }
                    }
                }
            }
        }
    }
    
    Ok(unregistered)
}
