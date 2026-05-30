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
        folder_name: Option<String>,
        relative_path: Option<String>,
        linked_folder_type: Option<String>,
        linked_folder_relative_path: Option<String>,
        linked_folder_external_path: Option<String>,
    ) -> Result<(), String>;
    fn get_project_folder(&self, project_id: &str) -> Result<Option<String>, String>;
    fn get_project_folder_name(&self, project_id: &str) -> Result<Option<String>, String>;
    fn find_matching_root(&self, path: &str) -> Result<Option<(String, String)>, String>;
    fn save_project_directory(
        &self,
        id: &str,
        project_id: &str,
        root_id: &str,
        relative_path: &str,
        name: &str,
    ) -> Result<(), String>;
    fn delete_project_directory(&self, project_id: &str) -> Result<(), String>;
    fn get_root_path(&self, root_id: &str) -> Result<Option<String>, String>;
    fn get_all_roots(&self) -> Result<Vec<(String, String)>, String>;
    fn create_root_direct(&self, name: &str, path: &str) -> Result<String, String>;
    fn get_project_directory(&self, project_id: &str) -> Result<Option<(String, String)>, String>;
    fn get_all_files(&self) -> Result<Vec<ProjectFile>, String>;
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
        folder_name: Option<String>,
        relative_path: Option<String>,
        linked_folder_type: Option<String>,
        linked_folder_relative_path: Option<String>,
        linked_folder_external_path: Option<String>,
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
            if let Some(f_name) = folder_name {
                if f_name.is_empty() {
                    store.projects[idx].folder_name = None;
                } else {
                    store.projects[idx].folder_name = Some(f_name);
                }
            }
            if let Some(rel_path) = relative_path {
                if rel_path.is_empty() {
                    store.projects[idx].relative_path = None;
                } else {
                    store.projects[idx].relative_path = Some(rel_path);
                }
            }
            if let Some(lf_type) = linked_folder_type {
                if lf_type.is_empty() {
                    store.projects[idx].linked_folder_type = None;
                } else {
                    store.projects[idx].linked_folder_type = Some(lf_type);
                }
            }
            if let Some(lf_rel) = linked_folder_relative_path {
                if lf_rel.is_empty() {
                    store.projects[idx].linked_folder_relative_path = None;
                } else {
                    store.projects[idx].linked_folder_relative_path = Some(lf_rel);
                }
            }
            if let Some(lf_ext) = linked_folder_external_path {
                if lf_ext.is_empty() {
                    store.projects[idx].linked_folder_external_path = None;
                } else {
                    store.projects[idx].linked_folder_external_path = Some(lf_ext);
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

    fn get_project_folder_name(&self, project_id: &str) -> Result<Option<String>, String> {
        let store = self.read_store()?;
        Ok(store
            .projects
            .into_iter()
            .find(|p| p.id == project_id)
            .and_then(|p| p.folder_name))
    }

    fn find_matching_root(&self, _path: &str) -> Result<Option<(String, String)>, String> {
        Ok(None)
    }

    fn save_project_directory(
        &self,
        _id: &str,
        _project_id: &str,
        _root_id: &str,
        _relative_path: &str,
        _name: &str,
    ) -> Result<(), String> {
        Ok(())
    }

    fn delete_project_directory(&self, _project_id: &str) -> Result<(), String> {
        Ok(())
    }

    fn get_root_path(&self, _root_id: &str) -> Result<Option<String>, String> {
        Ok(None)
    }

    fn get_all_roots(&self) -> Result<Vec<(String, String)>, String> {
        Ok(Vec::new())
    }

    fn create_root_direct(&self, _name: &str, _path: &str) -> Result<String, String> {
        Ok(format!("root_{}", chrono::Utc::now().timestamp_millis()))
    }

    fn get_project_directory(&self, _project_id: &str) -> Result<Option<(String, String)>, String> {
        Ok(None)
    }

    fn get_all_files(&self) -> Result<Vec<ProjectFile>, String> {
        let store = self.read_store()?;
        Ok(store.project_files)
    }
}

use std::sync::{Arc, Mutex, RwLock};

pub struct SqliteProjectFileRepository {
    conn: Arc<Mutex<rusqlite::Connection>>,
}

impl SqliteProjectFileRepository {
    pub fn new(conn: Arc<Mutex<rusqlite::Connection>>) -> Self {
        Self { conn }
    }
}

fn row_to_project_file(row: &rusqlite::Row) -> Result<ProjectFile, rusqlite::Error> {
    let exists_int: i32 = row.get(9)?;
    let is_main_doc_int: i32 = row.get(13)?;
    let is_main_budget_int: i32 = row.get(14)?;

    Ok(ProjectFile {
        id: row.get(0)?,
        project_id: row.get(1)?,
        file_name: row.get(2)?,
        file_path: row.get(3)?,
        original_path: row.get(4)?,
        managed_path: row.get(5)?,
        file_type: row.get(6)?,
        extension: row.get(7)?,
        size: row.get(8)?,
        exists: exists_int != 0,
        last_scanned_at: row.get(10)?,
        modified_at: row.get(11)?,
        storage_mode: row.get(12)?,
        is_main_document: is_main_doc_int != 0,
        is_main_budget_file: is_main_budget_int != 0,
        note: row.get(15)?,
        created_at: row.get(16)?,
        updated_at: row.get(17)?,
        root_id: row.get(18)?,
        directory_id: row.get(19)?,
        relative_path: row.get(20)?,
        absolute_path_snapshot: row.get(21)?,
        file_hash: row.get(22)?,
        file_role: row.get(23)?,
    })
}

impl ProjectFileRepository for SqliteProjectFileRepository {
    fn get_project_files(&self, project_id: &str) -> Result<Vec<ProjectFile>, String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;
        let mut stmt = conn
            .prepare("SELECT id, project_id, file_name, file_path, original_path, managed_path, file_type, extension, size, \"exists\", last_scanned_at, modified_at, storage_mode, is_main_document, is_main_budget_file, note, created_at, updated_at, root_id, directory_id, relative_path, absolute_path_snapshot, file_hash, file_role FROM project_files WHERE project_id = ?1")
            .map_err(|e| e.to_string())?;

        let file_iter = stmt
            .query_map([project_id], row_to_project_file)
            .map_err(|e| e.to_string())?;

        let mut list = Vec::new();
        for f in file_iter {
            list.push(f.map_err(|e| e.to_string())?);
        }
        Ok(list)
    }

    fn get_file(&self, id: &str) -> Result<Option<ProjectFile>, String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;
        let mut stmt = conn
            .prepare("SELECT id, project_id, file_name, file_path, original_path, managed_path, file_type, extension, size, \"exists\", last_scanned_at, modified_at, storage_mode, is_main_document, is_main_budget_file, note, created_at, updated_at, root_id, directory_id, relative_path, absolute_path_snapshot, file_hash, file_role FROM project_files WHERE id = ?1")
            .map_err(|e| e.to_string())?;

        let mut rows = stmt.query([id]).map_err(|e| e.to_string())?;
        if let Some(row) = rows.next().map_err(|e| e.to_string())? {
            let file = row_to_project_file(row).map_err(|e| e.to_string())?;
            Ok(Some(file))
        } else {
            Ok(None)
        }
    }

    fn save_file(&self, file: &ProjectFile) -> Result<(), String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;
        conn.execute(
            "INSERT OR REPLACE INTO project_files (
                id, project_id, file_name, file_path, original_path, managed_path, file_type, extension,
                size, \"exists\", last_scanned_at, modified_at, storage_mode, is_main_document, is_main_budget_file,
                note, created_at, updated_at, root_id, directory_id, relative_path, absolute_path_snapshot,
                file_hash, file_role
            ) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10, ?11, ?12, ?13, ?14, ?15, ?16, ?17, ?18, ?19, ?20, ?21, ?22, ?23, ?24)",
            rusqlite::params![
                file.id,
                file.project_id,
                file.file_name,
                file.file_path,
                file.original_path,
                file.managed_path,
                file.file_type,
                file.extension,
                file.size,
                if file.exists { 1 } else { 0 },
                file.last_scanned_at,
                file.modified_at,
                file.storage_mode,
                if file.is_main_document { 1 } else { 0 },
                if file.is_main_budget_file { 1 } else { 0 },
                file.note,
                file.created_at,
                file.updated_at,
                file.root_id,
                file.directory_id,
                file.relative_path,
                file.absolute_path_snapshot,
                file.file_hash,
                file.file_role,
            ],
        ).map_err(|e| e.to_string())?;
        Ok(())
    }

    fn delete_file(&self, id: &str) -> Result<(), String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;
        conn.execute("DELETE FROM project_files WHERE id = ?1", [id])
            .map_err(|e| e.to_string())?;
        Ok(())
    }

    fn save_files(&self, files: &[ProjectFile]) -> Result<(), String> {
        let mut conn = self.conn.lock().map_err(|e| e.to_string())?;
        let tx = conn.transaction().map_err(|e| e.to_string())?;
        for file in files {
            tx.execute(
                "INSERT OR REPLACE INTO project_files (
                    id, project_id, file_name, file_path, original_path, managed_path, file_type, extension,
                    size, \"exists\", last_scanned_at, modified_at, storage_mode, is_main_document, is_main_budget_file,
                    note, created_at, updated_at, root_id, directory_id, relative_path, absolute_path_snapshot,
                    file_hash, file_role
                ) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10, ?11, ?12, ?13, ?14, ?15, ?16, ?17, ?18, ?19, ?20, ?21, ?22, ?23, ?24)",
                rusqlite::params![
                    file.id,
                    file.project_id,
                    file.file_name,
                    file.file_path,
                    file.original_path,
                    file.managed_path,
                    file.file_type,
                    file.extension,
                    file.size,
                    if file.exists { 1 } else { 0 },
                    file.last_scanned_at,
                    file.modified_at,
                    file.storage_mode,
                    if file.is_main_document { 1 } else { 0 },
                    if file.is_main_budget_file { 1 } else { 0 },
                    file.note,
                    file.created_at,
                    file.updated_at,
                    file.root_id,
                    file.directory_id,
                    file.relative_path,
                    file.absolute_path_snapshot,
                    file.file_hash,
                    file.file_role,
                ],
            ).map_err(|e| e.to_string())?;
        }
        tx.commit().map_err(|e| e.to_string())?;
        Ok(())
    }

    fn update_project_fields(
        &self,
        project_id: &str,
        folder_path: Option<String>,
        main_doc: Option<String>,
        main_budget: Option<String>,
        folder_name: Option<String>,
        relative_path: Option<String>,
        linked_folder_type: Option<String>,
        linked_folder_relative_path: Option<String>,
        linked_folder_external_path: Option<String>,
    ) -> Result<(), String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;

        let mut query = String::from("UPDATE projects SET ");
        let mut params: Vec<Box<dyn rusqlite::ToSql>> = Vec::new();
        let mut updates = Vec::new();

        if let Some(path) = folder_path {
            updates.push(format!("folder_path = ?{}", updates.len() + 1));
            if path.is_empty() {
                params.push(Box::new(Option::<String>::None));
            } else {
                params.push(Box::new(path));
            }
        }

        if let Some(doc) = main_doc {
            updates.push(format!("main_document_path = ?{}", updates.len() + 1));
            if doc.is_empty() {
                params.push(Box::new(Option::<String>::None));
            } else {
                params.push(Box::new(doc));
            }
        }

        if let Some(budget) = main_budget {
            updates.push(format!("main_budget_file_path = ?{}", updates.len() + 1));
            if budget.is_empty() {
                params.push(Box::new(Option::<String>::None));
            } else {
                params.push(Box::new(budget));
            }
        }

        if let Some(f_name) = folder_name {
            updates.push(format!("folder_name = ?{}", updates.len() + 1));
            if f_name.is_empty() {
                params.push(Box::new(Option::<String>::None));
            } else {
                params.push(Box::new(f_name));
            }
        }

        if let Some(rel_p) = relative_path {
            updates.push(format!("relative_path = ?{}", updates.len() + 1));
            if rel_p.is_empty() {
                params.push(Box::new(Option::<String>::None));
            } else {
                params.push(Box::new(rel_p));
            }
        }

        if let Some(lf_type) = linked_folder_type {
            updates.push(format!("linked_folder_type = ?{}", updates.len() + 1));
            if lf_type.is_empty() {
                params.push(Box::new(Option::<String>::None));
            } else {
                params.push(Box::new(lf_type));
            }
        }

        if let Some(lf_rel) = linked_folder_relative_path {
            updates.push(format!(
                "linked_folder_relative_path = ?{}",
                updates.len() + 1
            ));
            if lf_rel.is_empty() {
                params.push(Box::new(Option::<String>::None));
            } else {
                params.push(Box::new(lf_rel));
            }
        }

        if let Some(lf_ext) = linked_folder_external_path {
            updates.push(format!(
                "linked_folder_external_path = ?{}",
                updates.len() + 1
            ));
            if lf_ext.is_empty() {
                params.push(Box::new(Option::<String>::None));
            } else {
                params.push(Box::new(lf_ext));
            }
        }

        if updates.is_empty() {
            return Ok(());
        }

        query.push_str(&updates.join(", "));
        query.push_str(&format!(" WHERE id = ?{}", updates.len() + 1));
        params.push(Box::new(project_id.to_string()));

        let params_refs: Vec<&dyn rusqlite::ToSql> = params.iter().map(|p| p.as_ref()).collect();

        conn.execute(&query, params_refs.as_slice())
            .map_err(|e| e.to_string())?;
        Ok(())
    }

    fn get_project_folder(&self, project_id: &str) -> Result<Option<String>, String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;
        let mut stmt = conn
            .prepare("SELECT folder_path FROM projects WHERE id = ?1")
            .map_err(|e| e.to_string())?;

        let mut rows = stmt.query([project_id]).map_err(|e| e.to_string())?;
        if let Some(row) = rows.next().map_err(|e| e.to_string())? {
            Ok(row.get(0).map_err(|e| e.to_string())?)
        } else {
            Ok(None)
        }
    }

    fn get_project_folder_name(&self, project_id: &str) -> Result<Option<String>, String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;
        let mut stmt = conn
            .prepare("SELECT folder_name FROM projects WHERE id = ?1")
            .map_err(|e| e.to_string())?;

        let mut rows = stmt.query([project_id]).map_err(|e| e.to_string())?;
        if let Some(row) = rows.next().map_err(|e| e.to_string())? {
            Ok(row.get(0).map_err(|e| e.to_string())?)
        } else {
            Ok(None)
        }
    }

    fn find_matching_root(&self, path: &str) -> Result<Option<(String, String)>, String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;
        let mut stmt = conn
            .prepare("SELECT id, root_path FROM project_roots")
            .map_err(|e| e.to_string())?;

        let root_iter = stmt
            .query_map([], |r| {
                let id: String = r.get(0)?;
                let root_path: String = r.get(1)?;
                Ok((id, root_path))
            })
            .map_err(|e| e.to_string())?;

        // Normalize target path
        let target = path.replace("\\", "/").to_lowercase();

        let mut matched = None;
        let mut longest_match_len = 0;

        for r in root_iter {
            let (id, root_path) = r.map_err(|e| e.to_string())?;
            let root_norm = root_path.replace("\\", "/").to_lowercase();

            // Check if target starts with root_norm
            if target.starts_with(&root_norm) {
                if root_norm.len() > longest_match_len {
                    longest_match_len = root_norm.len();

                    // Compute relative path
                    let rel = &path[root_path.len()..];
                    let rel_clean = rel
                        .trim_start_matches('\\')
                        .trim_start_matches('/')
                        .to_string();
                    matched = Some((id, rel_clean));
                }
            }
        }
        Ok(matched)
    }

    fn save_project_directory(
        &self,
        id: &str,
        project_id: &str,
        root_id: &str,
        relative_path: &str,
        name: &str,
    ) -> Result<(), String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;
        let now = chrono::Utc::now().to_rfc3339();
        conn.execute(
            "INSERT OR REPLACE INTO project_directories (id, project_id, root_id, relative_path, dir_name, created_at, updated_at) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7)",
            rusqlite::params![id, project_id, root_id, relative_path, name, now, now]
        ).map_err(|e| e.to_string())?;
        Ok(())
    }

    fn delete_project_directory(&self, project_id: &str) -> Result<(), String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;
        conn.execute(
            "DELETE FROM project_directories WHERE project_id = ?1",
            [project_id],
        )
        .map_err(|e| e.to_string())?;
        Ok(())
    }

    fn get_root_path(&self, root_id: &str) -> Result<Option<String>, String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;
        let mut stmt = conn
            .prepare("SELECT root_path FROM project_roots WHERE id = ?1")
            .map_err(|e| e.to_string())?;
        let mut rows = stmt.query([root_id]).map_err(|e| e.to_string())?;
        if let Some(row) = rows.next().map_err(|e| e.to_string())? {
            Ok(Some(row.get(0).map_err(|e| e.to_string())?))
        } else {
            Ok(None)
        }
    }

    fn get_all_roots(&self) -> Result<Vec<(String, String)>, String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;
        let mut stmt = conn
            .prepare("SELECT id, root_path FROM project_roots")
            .map_err(|e| e.to_string())?;
        let iter = stmt
            .query_map([], |r| Ok((r.get::<_, String>(0)?, r.get::<_, String>(1)?)))
            .map_err(|e| e.to_string())?;
        let mut list = Vec::new();
        for item in iter {
            list.push(item.map_err(|e| e.to_string())?);
        }
        Ok(list)
    }

    fn create_root_direct(&self, name: &str, path: &str) -> Result<String, String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;
        let id = format!("root_{}", chrono::Utc::now().timestamp_millis());
        let now = chrono::Utc::now().to_rfc3339();
        conn.execute(
            "INSERT INTO project_roots (id, name, root_path, root_alias, is_default, created_at, updated_at) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7)",
            rusqlite::params![id, name, path, Option::<String>::None, 0, now, now]
        ).map_err(|e| e.to_string())?;
        Ok(id)
    }

    fn get_project_directory(&self, project_id: &str) -> Result<Option<(String, String)>, String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;
        let mut stmt = conn
            .prepare("SELECT root_id, relative_path FROM project_directories WHERE project_id = ?1")
            .map_err(|e| e.to_string())?;
        let mut rows = stmt.query([project_id]).map_err(|e| e.to_string())?;
        if let Some(row) = rows.next().map_err(|e| e.to_string())? {
            Ok(Some((
                row.get(0).map_err(|e| e.to_string())?,
                row.get(1).map_err(|e| e.to_string())?,
            )))
        } else {
            Ok(None)
        }
    }

    fn get_all_files(&self) -> Result<Vec<ProjectFile>, String> {
        let conn = self.conn.lock().map_err(|e| e.to_string())?;
        let mut stmt = conn
            .prepare("SELECT id, project_id, file_name, file_path, original_path, managed_path, file_type, extension, size, \"exists\", last_scanned_at, modified_at, storage_mode, is_main_document, is_main_budget_file, note, created_at, updated_at, root_id, directory_id, relative_path, absolute_path_snapshot, file_hash, file_role FROM project_files")
            .map_err(|e| e.to_string())?;
        let file_iter = stmt
            .query_map([], row_to_project_file)
            .map_err(|e| e.to_string())?;
        let mut list = Vec::new();
        for f in file_iter {
            list.push(f.map_err(|e| e.to_string())?);
        }
        Ok(list)
    }
}

pub enum FileRepoBackend {
    Json(JsonProjectFileRepository),
    Sqlite(SqliteProjectFileRepository),
}

#[derive(Clone)]
pub struct DualProjectFileRepository {
    backend: Arc<RwLock<FileRepoBackend>>,
}

impl DualProjectFileRepository {
    pub fn new(backend: FileRepoBackend) -> Self {
        Self {
            backend: Arc::new(RwLock::new(backend)),
        }
    }

    pub fn switch_to_sqlite(&self, sqlite_repo: SqliteProjectFileRepository) {
        let mut backend = self.backend.write().unwrap();
        *backend = FileRepoBackend::Sqlite(sqlite_repo);
    }
}

impl ProjectFileRepository for DualProjectFileRepository {
    fn get_project_files(&self, project_id: &str) -> Result<Vec<ProjectFile>, String> {
        match &*self.backend.read().unwrap() {
            FileRepoBackend::Json(r) => r.get_project_files(project_id),
            FileRepoBackend::Sqlite(r) => r.get_project_files(project_id),
        }
    }

    fn get_file(&self, id: &str) -> Result<Option<ProjectFile>, String> {
        match &*self.backend.read().unwrap() {
            FileRepoBackend::Json(r) => r.get_file(id),
            FileRepoBackend::Sqlite(r) => r.get_file(id),
        }
    }

    fn save_file(&self, file: &ProjectFile) -> Result<(), String> {
        match &*self.backend.read().unwrap() {
            FileRepoBackend::Json(r) => r.save_file(file),
            FileRepoBackend::Sqlite(r) => r.save_file(file),
        }
    }

    fn delete_file(&self, id: &str) -> Result<(), String> {
        match &*self.backend.read().unwrap() {
            FileRepoBackend::Json(r) => r.delete_file(id),
            FileRepoBackend::Sqlite(r) => r.delete_file(id),
        }
    }

    fn save_files(&self, files: &[ProjectFile]) -> Result<(), String> {
        match &*self.backend.read().unwrap() {
            FileRepoBackend::Json(r) => r.save_files(files),
            FileRepoBackend::Sqlite(r) => r.save_files(files),
        }
    }

    fn update_project_fields(
        &self,
        project_id: &str,
        folder_path: Option<String>,
        main_doc: Option<String>,
        main_budget: Option<String>,
        folder_name: Option<String>,
        relative_path: Option<String>,
        linked_folder_type: Option<String>,
        linked_folder_relative_path: Option<String>,
        linked_folder_external_path: Option<String>,
    ) -> Result<(), String> {
        match &*self.backend.read().unwrap() {
            FileRepoBackend::Json(r) => r.update_project_fields(
                project_id,
                folder_path,
                main_doc,
                main_budget,
                folder_name,
                relative_path,
                linked_folder_type,
                linked_folder_relative_path,
                linked_folder_external_path,
            ),
            FileRepoBackend::Sqlite(r) => r.update_project_fields(
                project_id,
                folder_path,
                main_doc,
                main_budget,
                folder_name,
                relative_path,
                linked_folder_type,
                linked_folder_relative_path,
                linked_folder_external_path,
            ),
        }
    }

    fn get_project_folder(&self, project_id: &str) -> Result<Option<String>, String> {
        match &*self.backend.read().unwrap() {
            FileRepoBackend::Json(r) => r.get_project_folder(project_id),
            FileRepoBackend::Sqlite(r) => r.get_project_folder(project_id),
        }
    }

    fn get_project_folder_name(&self, project_id: &str) -> Result<Option<String>, String> {
        match &*self.backend.read().unwrap() {
            FileRepoBackend::Json(r) => r.get_project_folder_name(project_id),
            FileRepoBackend::Sqlite(r) => r.get_project_folder_name(project_id),
        }
    }

    fn find_matching_root(&self, path: &str) -> Result<Option<(String, String)>, String> {
        match &*self.backend.read().unwrap() {
            FileRepoBackend::Json(r) => r.find_matching_root(path),
            FileRepoBackend::Sqlite(r) => r.find_matching_root(path),
        }
    }

    fn save_project_directory(
        &self,
        id: &str,
        project_id: &str,
        root_id: &str,
        relative_path: &str,
        name: &str,
    ) -> Result<(), String> {
        match &*self.backend.read().unwrap() {
            FileRepoBackend::Json(r) => {
                r.save_project_directory(id, project_id, root_id, relative_path, name)
            }
            FileRepoBackend::Sqlite(r) => {
                r.save_project_directory(id, project_id, root_id, relative_path, name)
            }
        }
    }

    fn delete_project_directory(&self, project_id: &str) -> Result<(), String> {
        match &*self.backend.read().unwrap() {
            FileRepoBackend::Json(r) => r.delete_project_directory(project_id),
            FileRepoBackend::Sqlite(r) => r.delete_project_directory(project_id),
        }
    }

    fn get_root_path(&self, root_id: &str) -> Result<Option<String>, String> {
        match &*self.backend.read().unwrap() {
            FileRepoBackend::Json(r) => r.get_root_path(root_id),
            FileRepoBackend::Sqlite(r) => r.get_root_path(root_id),
        }
    }

    fn get_all_roots(&self) -> Result<Vec<(String, String)>, String> {
        match &*self.backend.read().unwrap() {
            FileRepoBackend::Json(r) => r.get_all_roots(),
            FileRepoBackend::Sqlite(r) => r.get_all_roots(),
        }
    }

    fn create_root_direct(&self, name: &str, path: &str) -> Result<String, String> {
        match &*self.backend.read().unwrap() {
            FileRepoBackend::Json(r) => r.create_root_direct(name, path),
            FileRepoBackend::Sqlite(r) => r.create_root_direct(name, path),
        }
    }

    fn get_project_directory(&self, project_id: &str) -> Result<Option<(String, String)>, String> {
        match &*self.backend.read().unwrap() {
            FileRepoBackend::Json(r) => r.get_project_directory(project_id),
            FileRepoBackend::Sqlite(r) => r.get_project_directory(project_id),
        }
    }

    fn get_all_files(&self) -> Result<Vec<ProjectFile>, String> {
        match &*self.backend.read().unwrap() {
            FileRepoBackend::Json(r) => r.get_all_files(),
            FileRepoBackend::Sqlite(r) => r.get_all_files(),
        }
    }
}
