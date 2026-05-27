use std::sync::Arc;
use std::path::Path;
use serde::{Deserialize, Serialize};
use tauri::State;
use super::service::ProjectFileService;
use super::repository::ProjectFileRepository;

#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct HealthReport {
    pub total_files: usize,
    pub healthy_files: usize,
    pub missing_files: usize,
    pub recoverable_files: usize,
    pub details: Vec<FileHealthDetail>,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct FileHealthDetail {
    pub file_id: String,
    pub project_id: String,
    pub file_name: String,
    pub current_path: String,
    pub status: String, // "healthy" | "recoverable" | "missing"
    pub recovered_path: Option<String>,
}

pub struct FileLinkHealthService {
    file_service: Arc<ProjectFileService>,
    repository: Arc<dyn ProjectFileRepository + Send + Sync>,
}

impl FileLinkHealthService {
    pub fn new(
        file_service: Arc<ProjectFileService>,
        repository: Arc<dyn ProjectFileRepository + Send + Sync>,
    ) -> Self {
        Self {
            file_service,
            repository,
        }
    }

    pub fn run_health_check(&self) -> Result<HealthReport, String> {
        let files = self.repository.get_all_files()?;
        let mut details = Vec::new();
        let mut healthy_files = 0;
        let mut recoverable_files = 0;
        let mut missing_files = 0;

        for mut file in files {
            let path = Path::new(&file.file_path);
            if path.exists() {
                healthy_files += 1;
                details.push(FileHealthDetail {
                    file_id: file.id.clone(),
                    project_id: file.project_id.clone(),
                    file_name: file.file_name.clone(),
                    current_path: file.file_path.clone(),
                    status: "healthy".to_string(),
                    recovered_path: None,
                });
            } else {
                // Try to resolve using path resilience priorities
                match self.file_service.resolve_file_path(&file) {
                    Ok(recovered) => {
                        // Recoverable! Auto-heal it by writing back to database
                        let old_path = file.file_path.clone();
                        file.file_path = recovered.clone();
                        file.exists = true;
                        file.updated_at = chrono::Utc::now().to_rfc3339();
                        let _ = self.repository.save_file(&file);

                        recoverable_files += 1;
                        details.push(FileHealthDetail {
                            file_id: file.id.clone(),
                            project_id: file.project_id.clone(),
                            file_name: file.file_name.clone(),
                            current_path: old_path,
                            status: "recoverable".to_string(),
                            recovered_path: Some(recovered),
                        });
                    }
                    Err(_) => {
                        // Missing
                        missing_files += 1;
                        if file.exists {
                            file.exists = false;
                            file.updated_at = chrono::Utc::now().to_rfc3339();
                            let _ = self.repository.save_file(&file);
                        }
                        details.push(FileHealthDetail {
                            file_id: file.id.clone(),
                            project_id: file.project_id.clone(),
                            file_name: file.file_name.clone(),
                            current_path: file.file_path.clone(),
                            status: "missing".to_string(),
                            recovered_path: None,
                        });
                    }
                }
            }
        }

        Ok(HealthReport {
            total_files: details.len(),
            healthy_files,
            missing_files,
            recoverable_files,
            details,
        })
    }
}

#[tauri::command]
pub async fn run_file_health_check(
    service: State<'_, Arc<FileLinkHealthService>>,
) -> Result<HealthReport, String> {
    service.run_health_check()
}
