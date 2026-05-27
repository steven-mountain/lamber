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
    fn get_project_setting(&self, project_id: &str, key: &str) -> Result<Option<String>, String>;
    fn save_project_setting(&self, project_id: &str, key: &str, value: &str) -> Result<(), String>;
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

    fn get_project_setting(&self, _project_id: &str, _key: &str) -> Result<Option<String>, String> {
        Ok(None)
    }

    fn save_project_setting(&self, _project_id: &str, _key: &str, _value: &str) -> Result<(), String> {
        Ok(())
    }
}

use std::sync::{Arc, Mutex, RwLock};

pub struct SqliteProjectRepository {
    conn: Arc<Mutex<rusqlite::Connection>>,
}

impl SqliteProjectRepository {
    pub fn new(conn: Arc<Mutex<rusqlite::Connection>>) -> Self {
        Self { conn }
    }
}

fn row_to_project(row: &rusqlite::Row) -> Result<Project, rusqlite::Error> {
    let summary_metrics_str: Option<String> = row.get(13)?;
    let summary_metrics = summary_metrics_str
        .and_then(|s| serde_json::from_str(&s).ok());
    
    let logs_str: Option<String> = row.get(18)?;
    let logs = logs_str
        .and_then(|s| serde_json::from_str(&s).ok())
        .unwrap_or_default();

    Ok(Project {
        id: row.get(0)?,
        name: row.get(1)?,
        customer_name: row.get(2)?,
        status: row.get(3)?,
        benefit_status: row.get(4)?,
        default_scheme_id: row.get(5)?,
        created_at: row.get(6)?,
        updated_at: row.get(7)?,
        total_revenue_incl: row.get(8)?,
        total_cost_incl: row.get(9)?,
        project_years: row.get(10)?,
        discount_rate: row.get(11)?,
        cashflow_model: row.get(12)?,
        summary_metrics,
        folder_path: row.get(14)?,
        main_document_path: row.get(15)?,
        main_budget_file_path: row.get(16)?,
        note: row.get(17)?,
        logs,
    })
}

impl ProjectRepository for SqliteProjectRepository {
    fn get_projects(&self) -> Result<Vec<Project>, String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;
        let mut stmt = conn
            .prepare("SELECT id, name, customer_name, status, benefit_status, default_scheme_id, created_at, updated_at, total_revenue_incl, total_cost_incl, project_years, discount_rate, cashflow_model, summary_metrics, folder_path, main_document_path, main_budget_file_path, note, logs FROM projects")
            .map_err(|e| e.to_string())?;
        
        let project_iter = stmt.query_map([], row_to_project).map_err(|e| e.to_string())?;

        let mut list = Vec::new();
        for p in project_iter {
            list.push(p.map_err(|e| e.to_string())?);
        }
        Ok(list)
    }

    fn get_project(&self, id: &str) -> Result<Option<Project>, String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;
        let mut stmt = conn
            .prepare("SELECT id, name, customer_name, status, benefit_status, default_scheme_id, created_at, updated_at, total_revenue_incl, total_cost_incl, project_years, discount_rate, cashflow_model, summary_metrics, folder_path, main_document_path, main_budget_file_path, note, logs FROM projects WHERE id = ?1")
            .map_err(|e| e.to_string())?;
        
        let mut rows = stmt.query([id]).map_err(|e| e.to_string())?;
        if let Some(row) = rows.next().map_err(|e| e.to_string())? {
            let p = row_to_project(row).map_err(|e| e.to_string())?;
            Ok(Some(p))
        } else {
            Ok(None)
        }
    }

    fn save_project(&self, project: &Project) -> Result<(), String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;
        let summary_metrics_str = project.summary_metrics.as_ref()
            .and_then(|m| serde_json::to_string(m).ok());
        let logs_str = serde_json::to_string(&project.logs).unwrap_or_default();

        let exists: bool = conn.query_row(
            "SELECT EXISTS(SELECT 1 FROM projects WHERE id = ?1)",
            [&project.id],
            |row| row.get(0),
        ).map_err(|e| e.to_string())?;

        if exists {
            conn.execute(
                "UPDATE projects SET 
                    name = ?1, customer_name = ?2, status = ?3, benefit_status = ?4, 
                    default_scheme_id = ?5, updated_at = ?6, total_revenue_incl = ?7, 
                    total_cost_incl = ?8, project_years = ?9, discount_rate = ?10, 
                    cashflow_model = ?11, summary_metrics = ?12, folder_path = ?13, 
                    main_document_path = ?14, main_budget_file_path = ?15, note = ?16, logs = ?17
                 WHERE id = ?18",
                rusqlite::params![
                    project.name,
                    project.customer_name,
                    project.status,
                    project.benefit_status,
                    project.default_scheme_id,
                    project.updated_at,
                    project.total_revenue_incl,
                    project.total_cost_incl,
                    project.project_years,
                    project.discount_rate,
                    project.cashflow_model,
                    summary_metrics_str,
                    project.folder_path,
                    project.main_document_path,
                    project.main_budget_file_path,
                    project.note,
                    logs_str,
                    project.id,
                ],
            ).map_err(|e| e.to_string())?;
        } else {
            conn.execute(
                "INSERT INTO projects (
                    id, name, customer_name, status, benefit_status, default_scheme_id, created_at, updated_at,
                    total_revenue_incl, total_cost_incl, project_years, discount_rate, cashflow_model,
                    summary_metrics, folder_path, main_document_path, main_budget_file_path, note, logs
                ) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10, ?11, ?12, ?13, ?14, ?15, ?16, ?17, ?18, ?19)",
                rusqlite::params![
                    project.id,
                    project.name,
                    project.customer_name,
                    project.status,
                    project.benefit_status,
                    project.default_scheme_id,
                    project.created_at,
                    project.updated_at,
                    project.total_revenue_incl,
                    project.total_cost_incl,
                    project.project_years,
                    project.discount_rate,
                    project.cashflow_model,
                    summary_metrics_str,
                    project.folder_path,
                    project.main_document_path,
                    project.main_budget_file_path,
                    project.note,
                    logs_str,
                ],
            ).map_err(|e| e.to_string())?;
        }
        Ok(())
    }

    fn delete_project(&self, id: &str) -> Result<(), String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;
        conn.execute("DELETE FROM projects WHERE id = ?1", [id])
            .map_err(|e| e.to_string())?;
        Ok(())
    }

    fn get_schemes(&self, project_id: &str) -> Result<Vec<BenefitAnalysisScheme>, String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;
        let mut stmt = conn
            .prepare("SELECT id, project_id, name, created_at, updated_at FROM benefit_schemes WHERE project_id = ?1")
            .map_err(|e| e.to_string())?;
        
        let scheme_iter = stmt.query_map([project_id], |row| {
            Ok(BenefitAnalysisScheme {
                id: row.get(0)?,
                project_id: row.get(1)?,
                name: row.get(2)?,
                created_at: row.get(3)?,
                updated_at: row.get(4)?,
            })
        }).map_err(|e| e.to_string())?;

        let mut list = Vec::new();
        for s in scheme_iter {
            list.push(s.map_err(|e| e.to_string())?);
        }
        Ok(list)
    }

    fn save_scheme(&self, scheme: &BenefitAnalysisScheme) -> Result<(), String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;
        conn.execute(
            "INSERT OR REPLACE INTO benefit_schemes (id, project_id, name, created_at, updated_at) VALUES (?1, ?2, ?3, ?4, ?5)",
            rusqlite::params![
                scheme.id,
                scheme.project_id,
                scheme.name,
                scheme.created_at,
                scheme.updated_at,
            ],
        ).map_err(|e| e.to_string())?;
        Ok(())
    }

    fn delete_scheme(&self, project_id: &str, scheme_id: &str) -> Result<(), String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;
        conn.execute(
            "DELETE FROM benefit_schemes WHERE project_id = ?1 AND id = ?2",
            [project_id, scheme_id],
        ).map_err(|e| e.to_string())?;
        Ok(())
    }

    fn get_snapshots(&self, scheme_id: &str) -> Result<Vec<BenefitAnalysisSnapshot>, String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;
        let mut stmt = conn
            .prepare("SELECT id, scheme_id, project_id, version, input_params, output_metrics, fingerprint, created_at FROM benefit_snapshots WHERE scheme_id = ?1")
            .map_err(|e| e.to_string())?;
        
        let snap_iter = stmt.query_map([scheme_id], |row| {
            let input_params_str: String = row.get(4)?;
            let input_params = serde_json::from_str(&input_params_str)
                .map_err(|e| rusqlite::Error::FromSqlConversionFailure(4, rusqlite::types::Type::Text, Box::new(e)))?;
            
            let output_metrics_str: String = row.get(5)?;
            let output_metrics = serde_json::from_str(&output_metrics_str)
                .map_err(|e| rusqlite::Error::FromSqlConversionFailure(5, rusqlite::types::Type::Text, Box::new(e)))?;

            Ok(BenefitAnalysisSnapshot {
                id: row.get(0)?,
                scheme_id: row.get(1)?,
                project_id: row.get(2)?,
                version: row.get(3)?,
                input_params,
                output_metrics,
                fingerprint: row.get(6)?,
                created_at: row.get(7)?,
            })
        }).map_err(|e| e.to_string())?;

        let mut list = Vec::new();
        for s in snap_iter {
            list.push(s.map_err(|e| e.to_string())?);
        }
        list.sort_by(|a, b| b.version.cmp(&a.version));
        Ok(list)
    }

    fn save_snapshot(&self, snapshot: &BenefitAnalysisSnapshot) -> Result<(), String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;
        let input_params_str = serde_json::to_string(&snapshot.input_params).map_err(|e| e.to_string())?;
        let output_metrics_str = serde_json::to_string(&snapshot.output_metrics).map_err(|e| e.to_string())?;

        conn.execute(
            "INSERT OR REPLACE INTO benefit_snapshots (id, scheme_id, project_id, version, input_params, output_metrics, fingerprint, created_at) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8)",
            rusqlite::params![
                snapshot.id,
                snapshot.scheme_id,
                snapshot.project_id,
                snapshot.version,
                input_params_str,
                output_metrics_str,
                snapshot.fingerprint,
                snapshot.created_at,
            ],
        ).map_err(|e| e.to_string())?;
        Ok(())
    }

    fn get_project_setting(&self, project_id: &str, key: &str) -> Result<Option<String>, String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;
        let mut stmt = conn
            .prepare("SELECT value FROM project_settings WHERE project_id = ?1 AND key = ?2")
            .map_err(|e| e.to_string())?;
        let mut rows = stmt.query([project_id, key]).map_err(|e| e.to_string())?;
        if let Some(row) = rows.next().map_err(|e| e.to_string())? {
            let val: String = row.get(0).map_err(|e| e.to_string())?;
            Ok(Some(val))
        } else {
            Ok(None)
        }
    }

    fn save_project_setting(&self, project_id: &str, key: &str, value: &str) -> Result<(), String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;
        let now = chrono::Utc::now().to_rfc3339();
        conn.execute(
            "INSERT OR REPLACE INTO project_settings (project_id, key, value, updated_at) VALUES (?1, ?2, ?3, ?4)",
            rusqlite::params![project_id, key, value, now],
        ).map_err(|e| e.to_string())?;
        Ok(())
    }
}

pub enum RepoBackend {
    Json(JsonProjectRepository),
    Sqlite(SqliteProjectRepository),
}

#[derive(Clone)]
pub struct DualProjectRepository {
    backend: Arc<RwLock<RepoBackend>>,
}

impl DualProjectRepository {
    pub fn new(backend: RepoBackend) -> Self {
        Self {
            backend: Arc::new(RwLock::new(backend)),
        }
    }

    pub fn switch_to_sqlite(&self, sqlite_repo: SqliteProjectRepository) {
        let mut backend = self.backend.write().unwrap();
        *backend = RepoBackend::Sqlite(sqlite_repo);
    }
}

impl ProjectRepository for DualProjectRepository {
    fn get_projects(&self) -> Result<Vec<Project>, String> {
        match &*self.backend.read().unwrap() {
            RepoBackend::Json(r) => r.get_projects(),
            RepoBackend::Sqlite(r) => r.get_projects(),
        }
    }

    fn get_project(&self, id: &str) -> Result<Option<Project>, String> {
        match &*self.backend.read().unwrap() {
            RepoBackend::Json(r) => r.get_project(id),
            RepoBackend::Sqlite(r) => r.get_project(id),
        }
    }

    fn save_project(&self, project: &Project) -> Result<(), String> {
        match &*self.backend.read().unwrap() {
            RepoBackend::Json(r) => r.save_project(project),
            RepoBackend::Sqlite(r) => r.save_project(project),
        }
    }

    fn delete_project(&self, id: &str) -> Result<(), String> {
        match &*self.backend.read().unwrap() {
            RepoBackend::Json(r) => r.delete_project(id),
            RepoBackend::Sqlite(r) => r.delete_project(id),
        }
    }

    fn get_schemes(&self, project_id: &str) -> Result<Vec<BenefitAnalysisScheme>, String> {
        match &*self.backend.read().unwrap() {
            RepoBackend::Json(r) => r.get_schemes(project_id),
            RepoBackend::Sqlite(r) => r.get_schemes(project_id),
        }
    }

    fn save_scheme(&self, scheme: &BenefitAnalysisScheme) -> Result<(), String> {
        match &*self.backend.read().unwrap() {
            RepoBackend::Json(r) => r.save_scheme(scheme),
            RepoBackend::Sqlite(r) => r.save_scheme(scheme),
        }
    }

    fn delete_scheme(&self, project_id: &str, scheme_id: &str) -> Result<(), String> {
        match &*self.backend.read().unwrap() {
            RepoBackend::Json(r) => r.delete_scheme(project_id, scheme_id),
            RepoBackend::Sqlite(r) => r.delete_scheme(project_id, scheme_id),
        }
    }

    fn get_snapshots(&self, scheme_id: &str) -> Result<Vec<BenefitAnalysisSnapshot>, String> {
        match &*self.backend.read().unwrap() {
            RepoBackend::Json(r) => r.get_snapshots(scheme_id),
            RepoBackend::Sqlite(r) => r.get_snapshots(scheme_id),
        }
    }

    fn save_snapshot(&self, snapshot: &BenefitAnalysisSnapshot) -> Result<(), String> {
        match &*self.backend.read().unwrap() {
            RepoBackend::Json(r) => r.save_snapshot(snapshot),
            RepoBackend::Sqlite(r) => r.save_snapshot(snapshot),
        }
    }

    fn get_project_setting(&self, project_id: &str, key: &str) -> Result<Option<String>, String> {
        match &*self.backend.read().unwrap() {
            RepoBackend::Json(r) => r.get_project_setting(project_id, key),
            RepoBackend::Sqlite(r) => r.get_project_setting(project_id, key),
        }
    }

    fn save_project_setting(&self, project_id: &str, key: &str, value: &str) -> Result<(), String> {
        match &*self.backend.read().unwrap() {
            RepoBackend::Json(r) => r.save_project_setting(project_id, key, value),
            RepoBackend::Sqlite(r) => r.save_project_setting(project_id, key, value),
        }
    }
}
