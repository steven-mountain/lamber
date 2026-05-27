use std::sync::{Arc, Mutex};
use rusqlite::params;
use serde::{Deserialize, Serialize};
use std::path::Path;
use tauri::State;

#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct RelocationPreview {
    pub total_items: usize,
    pub matched_items: usize,
    pub missing_items: usize,
    pub details: Vec<RelocationPreviewDetail>,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct RelocationPreviewDetail {
    pub item_id: String,
    pub item_type: String, // "file" | "directory"
    pub name: String,
    pub old_path: String,
    pub new_path: String,
    pub exists: bool,
}

pub struct BulkRelocationService {
    conn: Arc<Mutex<rusqlite::Connection>>,
}

impl BulkRelocationService {
    pub fn new(conn: Arc<Mutex<rusqlite::Connection>>) -> Self {
        Self { conn }
    }

    pub fn get_relocation_preview(&self, old_root_id: &str, new_root_path: &str) -> Result<RelocationPreview, String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;

        // Fetch old root path
        let old_root_path: String = conn.query_row(
            "SELECT root_path FROM project_roots WHERE id = ?1",
            [old_root_id],
            |r| r.get(0),
        ).map_err(|e| format!("未找到指定的根目录: {}", e))?;

        let mut details = Vec::new();
        let mut matched_items = 0;
        let mut missing_items = 0;

        // Preview directories under this root
        let mut stmt_dirs = conn.prepare(
            "SELECT id, relative_path, dir_name FROM project_directories WHERE root_id = ?1"
        ).map_err(|e| e.to_string())?;
        
        let dir_iter = stmt_dirs.query_map([old_root_id], |r| {
            Ok((r.get::<_, String>(0)?, r.get::<_, String>(1)?, r.get::<_, String>(2)?))
        }).map_err(|e| e.to_string())?;

        for d in dir_iter {
            let (id, relative_path, dir_name) = d.map_err(|e| e.to_string())?;
            let old_path = Path::new(&old_root_path).join(&relative_path).to_string_lossy().to_string();
            let new_path = Path::new(new_root_path).join(&relative_path);
            let exists = new_path.exists();
            if exists { matched_items += 1; } else { missing_items += 1; }
            details.push(RelocationPreviewDetail {
                item_id: id,
                item_type: "directory".to_string(),
                name: dir_name,
                old_path,
                new_path: new_path.to_string_lossy().to_string(),
                exists,
            });
        }

        // Preview files under this root
        let mut stmt_files = conn.prepare(
            "SELECT id, relative_path, file_name, file_path FROM project_files WHERE root_id = ?1"
        ).map_err(|e| e.to_string())?;

        let file_iter = stmt_files.query_map([old_root_id], |r| {
            Ok((r.get::<_, String>(0)?, r.get::<_, Option<String>>(1)?, r.get::<_, String>(2)?, r.get::<_, String>(3)?))
        }).map_err(|e| e.to_string())?;

        for f in file_iter {
            let (id, relative_path, file_name, old_path) = f.map_err(|e| e.to_string())?;
            let relative_path = relative_path.unwrap_or_default();
            let new_path = Path::new(new_root_path).join(&relative_path);
            let exists = new_path.exists();
            if exists { matched_items += 1; } else { missing_items += 1; }
            details.push(RelocationPreviewDetail {
                item_id: id,
                item_type: "file".to_string(),
                name: file_name,
                old_path,
                new_path: new_path.to_string_lossy().to_string(),
                exists,
            });
        }

        Ok(RelocationPreview {
            total_items: details.len(),
            matched_items,
            missing_items,
            details,
        })
    }

    pub fn execute_bulk_relocation(&self, old_root_id: &str, new_root_path: &str) -> Result<(), String> {
        let mut conn = self.conn.lock().map_err(|e| e.to_string())?;
        let tx = conn.transaction().map_err(|e| e.to_string())?;

        // 1. Update root path
        tx.execute(
            "UPDATE project_roots SET root_path = ?1, updated_at = ?2 WHERE id = ?3",
            params![new_root_path, chrono::Utc::now().to_rfc3339(), old_root_id],
        ).map_err(|e| e.to_string())?;

        // 2. Relocate project directories and update projects.folder_path
        let mut stmt_dirs = tx.prepare(
            "SELECT id, project_id, relative_path FROM project_directories WHERE root_id = ?1"
        ).map_err(|e| e.to_string())?;

        let dir_iter = stmt_dirs.query_map([old_root_id], |r| {
            Ok((r.get::<_, String>(0)?, r.get::<_, String>(1)?, r.get::<_, String>(2)?))
        }).map_err(|e| e.to_string())?;

        let mut dirs_to_update = Vec::new();
        for d in dir_iter {
            dirs_to_update.push(d.map_err(|e| e.to_string())?);
        }
        drop(stmt_dirs);

        for (_id, project_id, relative_path) in dirs_to_update {
            let new_path = Path::new(new_root_path).join(&relative_path).to_string_lossy().to_string();
            tx.execute(
                "UPDATE projects SET folder_path = ?1 WHERE id = ?2",
                params![new_path, project_id],
            ).map_err(|e| e.to_string())?;
        }

        // 3. Relocate project files
        let mut stmt_files = tx.prepare(
            "SELECT id, relative_path FROM project_files WHERE root_id = ?1"
        ).map_err(|e| e.to_string())?;

        let file_iter = stmt_files.query_map([old_root_id], |r| {
            Ok((r.get::<_, String>(0)?, r.get::<_, Option<String>>(1)?))
        }).map_err(|e| e.to_string())?;

        let mut files_to_update = Vec::new();
        for f in file_iter {
            files_to_update.push(f.map_err(|e| e.to_string())?);
        }
        drop(stmt_files);

        for (id, relative_path) in files_to_update {
            let relative_path = relative_path.unwrap_or_default();
            let new_path_buf = Path::new(new_root_path).join(&relative_path);
            let new_path = new_path_buf.to_string_lossy().to_string();
            let exists = new_path_buf.exists();
            tx.execute(
                "UPDATE project_files SET file_path = ?1, absolute_path_snapshot = ?1, \"exists\" = ?2, updated_at = ?3 WHERE id = ?4",
                params![new_path, if exists { 1 } else { 0 }, chrono::Utc::now().to_rfc3339(), id],
            ).map_err(|e| e.to_string())?;
        }

        tx.commit().map_err(|e| e.to_string())?;
        Ok(())
    }
}

// Tauri commands
#[tauri::command]
pub async fn get_relocation_preview(
    service: State<'_, Arc<BulkRelocationService>>,
    old_root_id: String,
    new_root_path: String,
) -> Result<RelocationPreview, String> {
    service.get_relocation_preview(&old_root_id, &new_root_path)
}

#[tauri::command]
pub async fn execute_bulk_relocation(
    service: State<'_, Arc<BulkRelocationService>>,
    old_root_id: String,
    new_root_path: String,
) -> Result<(), String> {
    service.execute_bulk_relocation(&old_root_id, &new_root_path)
}
