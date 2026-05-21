use super::models::ProjectFile;
use crate::benefit::models::StoreData;
use std::fs;
use std::path::PathBuf;

pub trait ProjectFileRepository {
    fn get_project_files(&self, project_id: &str) -> Result<Vec<ProjectFile>, String>;
    fn get_file(&self, id: &str) -> Result<Option<ProjectFile>, String>;
    fn save_file(&self, file: &ProjectFile) -> Result<(), String>;
    fn delete_file(&self, id: &str) -> Result<(), String>;
    fn save_files(&self, files: &[ProjectFile]) -> Result<(), String>;
    fn update_project_fields(
        &self,
        project_id: &str,
        folder_path: Option<String>,
        main_doc: Option<String>,
        main_budget: Option<String>,
    ) -> Result<(), String>;
    fn get_project_folder(&self, project_id: &str) -> Result<Option<String>, String>;
}

pub struct JsonProjectFileRepository {
    file_path: PathBuf,
}

impl JsonProjectFileRepository {
    pub fn new(file_path: PathBuf) -> Self {
        Self { file_path }
    }

    fn read_store(&self) -> Result<StoreData, String> {
        if !self.file_path.exists() {
            return Ok(StoreData {
                schema_version: 1,
                projects: Vec::new(),
                schemes: Vec::new(),
                snapshots: Vec::new(),
                project_files: Vec::new(),
            });
        }
        let content = fs::read_to_string(&self.file_path)
            .map_err(|e| format!("Failed to read store file: {}", e))?;
        if content.trim().is_empty() {
            return Ok(StoreData {
                schema_version: 1,
                projects: Vec::new(),
                schemes: Vec::new(),
                snapshots: Vec::new(),
                project_files: Vec::new(),
            });
        }
        let store: StoreData = serde_json::from_str(&content)
            .map_err(|e| format!("Failed to parse store data: {}", e))?;
        Ok(store)
    }

    fn write_store(&self, store: &StoreData) -> Result<(), String> {
        if let Some(parent) = self.file_path.parent() {
            if !parent.exists() {
                fs::create_dir_all(parent)
                    .map_err(|e| format!("Failed to create directories: {}", e))?;
            }
        }
        let content = serde_json::to_string_pretty(store)
            .map_err(|e| format!("Failed to serialize store data: {}", e))?;
        fs::write(&self.file_path, content)
            .map_err(|e| format!("Failed to write store file: {}", e))?;
        Ok(())
    }
}

impl ProjectFileRepository for JsonProjectFileRepository {
    fn get_project_files(&self, project_id: &str) -> Result<Vec<ProjectFile>, String> {
        let store = self.read_store()?;
        Ok(store
            .project_files
            .into_iter()
            .filter(|f| f.project_id == project_id)
            .collect())
    }

    fn get_file(&self, id: &str) -> Result<Option<ProjectFile>, String> {
        let store = self.read_store()?;
        Ok(store.project_files.into_iter().find(|f| f.id == id))
    }

    fn save_file(&self, file: &ProjectFile) -> Result<(), String> {
        let mut store = self.read_store()?;
        if let Some(idx) = store.project_files.iter().position(|f| f.id == file.id) {
            store.project_files[idx] = file.clone();
        } else {
            store.project_files.push(file.clone());
        }
        self.write_store(&store)
    }

    fn delete_file(&self, id: &str) -> Result<(), String> {
        let mut store = self.read_store()?;
        store.project_files.retain(|f| f.id != id);
        self.write_store(&store)
    }

    fn save_files(&self, files: &[ProjectFile]) -> Result<(), String> {
        let mut store = self.read_store()?;
        for file in files {
            if let Some(idx) = store.project_files.iter().position(|f| f.id == file.id) {
                store.project_files[idx] = file.clone();
            } else {
                store.project_files.push(file.clone());
            }
        }
        self.write_store(&store)
    }

    fn update_project_fields(
        &self,
        project_id: &str,
        folder_path: Option<String>,
        main_doc: Option<String>,
        main_budget: Option<String>,
    ) -> Result<(), String> {
        let mut store = self.read_store()?;
        if let Some(idx) = store.projects.iter().position(|p| p.id == project_id) {
            if let Some(path) = folder_path {
                if path.is_empty() {
                    store.projects[idx].folder_path = None;
                } else {
                    store.projects[idx].folder_path = Some(path);
                }
            }
            if let Some(doc) = main_doc {
                if doc.is_empty() {
                    store.projects[idx].main_document_path = None;
                } else {
                    store.projects[idx].main_document_path = Some(doc);
                }
            }
            if let Some(budget) = main_budget {
                if budget.is_empty() {
                    store.projects[idx].main_budget_file_path = None;
                } else {
                    store.projects[idx].main_budget_file_path = Some(budget);
                }
            }
            self.write_store(&store)?;
            Ok(())
        } else {
            Err(format!("Project with ID {} not found", project_id))
        }
    }

    fn get_project_folder(&self, project_id: &str) -> Result<Option<String>, String> {
        let store = self.read_store()?;
        Ok(store
            .projects
            .into_iter()
            .find(|p| p.id == project_id)
            .and_then(|p| p.folder_path))
    }
}
