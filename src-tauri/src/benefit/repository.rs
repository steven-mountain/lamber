use super::models::{BenefitAnalysisScheme, BenefitAnalysisSnapshot, Project, StoreData};
use std::fs;
use std::path::PathBuf;

pub trait ProjectRepository {
    fn get_projects(&self) -> Result<Vec<Project>, String>;
    fn get_project(&self, id: &str) -> Result<Option<Project>, String>;
    fn save_project(&self, project: &Project) -> Result<(), String>;
    fn delete_project(&self, id: &str) -> Result<(), String>;

    fn get_schemes(&self, project_id: &str) -> Result<Vec<BenefitAnalysisScheme>, String>;
    fn save_scheme(&self, scheme: &BenefitAnalysisScheme) -> Result<(), String>;
    fn delete_scheme(&self, project_id: &str, scheme_id: &str) -> Result<(), String>;

    fn get_snapshots(&self, scheme_id: &str) -> Result<Vec<BenefitAnalysisSnapshot>, String>;
    fn save_snapshot(&self, snapshot: &BenefitAnalysisSnapshot) -> Result<(), String>;
}

pub struct JsonProjectRepository {
    file_path: PathBuf,
}

impl JsonProjectRepository {
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

impl ProjectRepository for JsonProjectRepository {
    fn get_projects(&self) -> Result<Vec<Project>, String> {
        let store = self.read_store()?;
        Ok(store.projects)
    }

    fn get_project(&self, id: &str) -> Result<Option<Project>, String> {
        let store = self.read_store()?;
        Ok(store.projects.into_iter().find(|p| p.id == id))
    }

    fn save_project(&self, project: &Project) -> Result<(), String> {
        let mut store = self.read_store()?;
        if let Some(idx) = store.projects.iter().position(|p| p.id == project.id) {
            store.projects[idx] = project.clone();
        } else {
            store.projects.push(project.clone());
        }
        self.write_store(&store)
    }

    fn delete_project(&self, id: &str) -> Result<(), String> {
        let mut store = self.read_store()?;
        store.projects.retain(|p| p.id != id);
        store.schemes.retain(|s| s.project_id != id);
        store.snapshots.retain(|sn| sn.project_id != id);
        self.write_store(&store)
    }

    fn get_schemes(&self, project_id: &str) -> Result<Vec<BenefitAnalysisScheme>, String> {
        let store = self.read_store()?;
        Ok(store
            .schemes
            .into_iter()
            .filter(|s| s.project_id == project_id)
            .collect())
    }

    fn save_scheme(&self, scheme: &BenefitAnalysisScheme) -> Result<(), String> {
        let mut store = self.read_store()?;
        if let Some(idx) = store.schemes.iter().position(|s| s.id == scheme.id) {
            store.schemes[idx] = scheme.clone();
        } else {
            store.schemes.push(scheme.clone());
        }
        self.write_store(&store)
    }

    fn delete_scheme(&self, project_id: &str, scheme_id: &str) -> Result<(), String> {
        let mut store = self.read_store()?;
        store
            .schemes
            .retain(|s| !(s.project_id == project_id && s.id == scheme_id));
        store
            .snapshots
            .retain(|sn| !(sn.project_id == project_id && sn.scheme_id == scheme_id));
        self.write_store(&store)
    }

    fn get_snapshots(&self, scheme_id: &str) -> Result<Vec<BenefitAnalysisSnapshot>, String> {
        let store = self.read_store()?;
        let mut list: Vec<BenefitAnalysisSnapshot> = store
            .snapshots
            .into_iter()
            .filter(|sn| sn.scheme_id == scheme_id)
            .collect();
        list.sort_by(|a, b| b.version.cmp(&a.version));
        Ok(list)
    }

    fn save_snapshot(&self, snapshot: &BenefitAnalysisSnapshot) -> Result<(), String> {
        let mut store = self.read_store()?;
        if let Some(idx) = store.snapshots.iter().position(|sn| sn.id == snapshot.id) {
            store.snapshots[idx] = snapshot.clone();
        } else {
            store.snapshots.push(snapshot.clone());
        }
        self.write_store(&store)
    }
}
