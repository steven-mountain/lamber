use chrono::Utc;
use rusqlite::params;
use serde::{Deserialize, Serialize};
use std::sync::{Arc, Mutex};
use tauri::State;

#[derive(Serialize, Deserialize, Clone, Debug)]
#[serde(rename_all = "camelCase")]
pub struct ProjectRoot {
    pub id: String,
    pub name: String,
    pub root_path: String,
    pub root_alias: Option<String>,
    pub is_default: bool,
    pub created_at: String,
    pub updated_at: String,
}

pub trait ProjectRootRepository {
    fn get_roots(&self) -> Result<Vec<ProjectRoot>, String>;
    fn get_root(&self, id: &str) -> Result<Option<ProjectRoot>, String>;
    fn save_root(&self, root: &ProjectRoot) -> Result<(), String>;
    fn delete_root(&self, id: &str) -> Result<(), String>;
    fn get_default_root(&self) -> Result<Option<ProjectRoot>, String>;
    fn clear_defaults(&self) -> Result<(), String>;
    fn check_references(&self, root_id: &str) -> Result<(usize, usize), String>;
}

pub struct SqliteProjectRootRepository {
    conn: Arc<Mutex<rusqlite::Connection>>,
}

impl SqliteProjectRootRepository {
    pub fn new(conn: Arc<Mutex<rusqlite::Connection>>) -> Self {
        Self { conn }
    }
}

fn row_to_root(row: &rusqlite::Row) -> Result<ProjectRoot, rusqlite::Error> {
    let is_default_int: i32 = row.get(4)?;
    Ok(ProjectRoot {
        id: row.get(0)?,
        name: row.get(1)?,
        root_path: row.get(2)?,
        root_alias: row.get(3)?,
        is_default: is_default_int != 0,
        created_at: row.get(5)?,
        updated_at: row.get(6)?,
    })
}

impl ProjectRootRepository for SqliteProjectRootRepository {
    fn get_roots(&self) -> Result<Vec<ProjectRoot>, String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;
        let mut stmt = conn
            .prepare("SELECT id, name, root_path, root_alias, is_default, created_at, updated_at FROM project_roots ORDER BY created_at DESC")
            .map_err(|e| e.to_string())?;

        let root_iter = stmt.query_map([], row_to_root).map_err(|e| e.to_string())?;
        let mut list = Vec::new();
        for r in root_iter {
            list.push(r.map_err(|e| e.to_string())?);
        }
        Ok(list)
    }

    fn get_root(&self, id: &str) -> Result<Option<ProjectRoot>, String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;
        let mut stmt = conn
            .prepare("SELECT id, name, root_path, root_alias, is_default, created_at, updated_at FROM project_roots WHERE id = ?1")
            .map_err(|e| e.to_string())?;

        let mut rows = stmt.query([id]).map_err(|e| e.to_string())?;
        if let Some(row) = rows.next().map_err(|e| e.to_string())? {
            let root = row_to_root(row).map_err(|e| e.to_string())?;
            Ok(Some(root))
        } else {
            Ok(None)
        }
    }

    fn save_root(&self, root: &ProjectRoot) -> Result<(), String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;
        conn.execute(
            "INSERT OR REPLACE INTO project_roots (id, name, root_path, root_alias, is_default, created_at, updated_at) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7)",
            params![
                root.id,
                root.name,
                root.root_path,
                root.root_alias,
                if root.is_default { 1 } else { 0 },
                root.created_at,
                root.updated_at,
            ],
        ).map_err(|e| e.to_string())?;
        Ok(())
    }

    fn delete_root(&self, id: &str) -> Result<(), String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;
        conn.execute("DELETE FROM project_roots WHERE id = ?1", [id])
            .map_err(|e| e.to_string())?;
        Ok(())
    }

    fn get_default_root(&self) -> Result<Option<ProjectRoot>, String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;
        let mut stmt = conn
            .prepare("SELECT id, name, root_path, root_alias, is_default, created_at, updated_at FROM project_roots WHERE is_default = 1")
            .map_err(|e| e.to_string())?;

        let mut rows = stmt.query([]).map_err(|e| e.to_string())?;
        if let Some(row) = rows.next().map_err(|e| e.to_string())? {
            let root = row_to_root(row).map_err(|e| e.to_string())?;
            Ok(Some(root))
        } else {
            Ok(None)
        }
    }

    fn clear_defaults(&self) -> Result<(), String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;
        conn.execute("UPDATE project_roots SET is_default = 0", [])
            .map_err(|e| e.to_string())?;
        Ok(())
    }

    fn check_references(&self, root_id: &str) -> Result<(usize, usize), String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;

        let dirs_count: usize = conn
            .query_row(
                "SELECT COUNT(*) FROM project_directories WHERE root_id = ?1",
                [root_id],
                |r| r.get(0),
            )
            .map_err(|e| e.to_string())?;

        let files_count: usize = conn
            .query_row(
                "SELECT COUNT(*) FROM project_files WHERE root_id = ?1",
                [root_id],
                |r| r.get(0),
            )
            .map_err(|e| e.to_string())?;

        Ok((dirs_count, files_count))
    }
}

pub struct ProjectRootService {
    repository: Arc<dyn ProjectRootRepository + Send + Sync>,
}

impl ProjectRootService {
    pub fn new(repository: Arc<dyn ProjectRootRepository + Send + Sync>) -> Self {
        Self { repository }
    }

    pub fn get_roots(&self) -> Result<Vec<ProjectRoot>, String> {
        self.repository.get_roots()
    }

    pub fn create_root(
        &self,
        name: String,
        root_path: String,
        root_alias: Option<String>,
        is_default: bool,
    ) -> Result<ProjectRoot, String> {
        let name = name.trim().to_string();
        if name.is_empty() {
            return Err("根目录名称不能为空".to_string());
        }

        let root_path = root_path.trim().to_string();
        if root_path.is_empty() {
            return Err("根目录物理路径不能为空".to_string());
        }

        // Normalize slashes for comparison
        let normalized_path = root_path.replace("\\", "/");
        let path = std::path::Path::new(&root_path);
        if !path.exists() || !path.is_dir() {
            return Err("指定的物理路径不存在或不是文件夹".to_string());
        }

        // Check if root path already exists
        let existing = self.repository.get_roots()?;
        if existing
            .iter()
            .any(|r| r.root_path.replace("\\", "/") == normalized_path)
        {
            return Err("该物理路径已被登记为项目根目录".to_string());
        }

        if is_default {
            self.repository.clear_defaults()?;
        }

        let id = format!("root_{}", Utc::now().timestamp_millis());
        let now = Utc::now().to_rfc3339();

        let root = ProjectRoot {
            id,
            name,
            root_path,
            root_alias,
            is_default,
            created_at: now.clone(),
            updated_at: now,
        };

        self.repository.save_root(&root)?;
        Ok(root)
    }

    pub fn update_root(&self, mut root: ProjectRoot) -> Result<ProjectRoot, String> {
        root.name = root.name.trim().to_string();
        if root.name.is_empty() {
            return Err("根目录名称不能为空".to_string());
        }

        root.root_path = root.root_path.trim().to_string();
        if root.root_path.is_empty() {
            return Err("根目录物理路径不能为空".to_string());
        }

        let path = std::path::Path::new(&root.root_path);
        if !path.exists() || !path.is_dir() {
            return Err("指定的物理路径不存在".to_string());
        }

        let normalized_path = root.root_path.replace("\\", "/");
        let existing = self.repository.get_roots()?;
        if existing
            .iter()
            .any(|r| r.id != root.id && r.root_path.replace("\\", "/") == normalized_path)
        {
            return Err("另一个根目录已登记了相同的物理路径".to_string());
        }

        if root.is_default {
            self.repository.clear_defaults()?;
        }

        root.updated_at = Utc::now().to_rfc3339();
        self.repository.save_root(&root)?;
        Ok(root)
    }

    pub fn delete_root(&self, id: &str) -> Result<(), String> {
        let (dirs, files) = self.repository.check_references(id)?;
        if dirs > 0 || files > 0 {
            return Err(format!(
                "无法删除此根目录，因为它目前仍被 {} 个项目文件夹和 {} 个文件记录引用。请先解除关联或执行重定位。",
                dirs, files
            ));
        }

        // If we delete default root, try to set another default if available
        let root_opt = self.repository.get_root(id)?;
        if let Some(root) = root_opt {
            if root.is_default {
                let existing = self.repository.get_roots()?;
                if let Some(other) = existing.iter().find(|r| r.id != id) {
                    let mut updated = other.clone();
                    updated.is_default = true;
                    self.repository.save_root(&updated)?;
                }
            }
        }

        self.repository.delete_root(id)
    }

    pub fn set_default_root(&self, id: &str) -> Result<(), String> {
        let root_opt = self.repository.get_root(id)?;
        if let Some(mut root) = root_opt {
            self.repository.clear_defaults()?;
            root.is_default = true;
            root.updated_at = Utc::now().to_rfc3339();
            self.repository.save_root(&root)?;
            Ok(())
        } else {
            Err("未找到指定的项目根目录".to_string())
        }
    }
}

// Tauri commands
#[tauri::command]
pub async fn get_project_roots(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
) -> Result<Vec<ProjectRoot>, String> {
    let repo = Arc::new(SqliteProjectRootRepository::new(runtime.require_db()?));
    let service = ProjectRootService::new(repo);
    service.get_roots()
}

#[tauri::command]
pub async fn create_project_root(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    name: String,
    root_path: String,
    root_alias: Option<String>,
    is_default: bool,
) -> Result<ProjectRoot, String> {
    let repo = Arc::new(SqliteProjectRootRepository::new(runtime.require_db()?));
    let service = ProjectRootService::new(repo);
    service.create_root(name, root_path, root_alias, is_default)
}

#[tauri::command]
pub async fn update_project_root(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    root: ProjectRoot,
) -> Result<ProjectRoot, String> {
    let repo = Arc::new(SqliteProjectRootRepository::new(runtime.require_db()?));
    let service = ProjectRootService::new(repo);
    service.update_root(root)
}

#[tauri::command]
pub async fn delete_project_root(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    id: String,
) -> Result<(), String> {
    let repo = Arc::new(SqliteProjectRootRepository::new(runtime.require_db()?));
    let service = ProjectRootService::new(repo);
    service.delete_root(&id)
}

#[tauri::command]
pub async fn set_default_project_root(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    id: String,
) -> Result<(), String> {
    let repo = Arc::new(SqliteProjectRootRepository::new(runtime.require_db()?));
    let service = ProjectRootService::new(repo);
    service.set_default_root(&id)
}
