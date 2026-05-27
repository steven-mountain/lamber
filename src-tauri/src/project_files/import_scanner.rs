use std::sync::{Arc, Mutex};
use std::path::Path;
use std::fs;
use serde::{Deserialize, Serialize};
use chrono::{DateTime, Utc};
use tauri::State;
use std::collections::hash_map::DefaultHasher;
use std::hash::{Hash, Hasher};

#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct ImportCandidate {
    pub folder_name: String,
    pub folder_path: String,
    pub exists_conflict: bool,
    pub files: Vec<CandidateFile>,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct CandidateFile {
    pub name: String,
    pub path: String,
    pub file_role: String, // "benefit_scheme" | "budget" | "proposal" | "other"
}

#[derive(Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct ImportSelection {
    pub folder_path: String,
    pub conflict_action: String, // "merge" | "skip" | "new"
}

pub struct ImportScanner {
    conn: Arc<Mutex<rusqlite::Connection>>,
}

impl ImportScanner {
    pub fn new(conn: Arc<Mutex<rusqlite::Connection>>) -> Self {
        Self { conn }
    }

    pub fn scan_import_candidates(&self, parent_path: &str) -> Result<Vec<ImportCandidate>, String> {
        let mut candidates = scan_for_candidates(parent_path)?;
        let conn = self.conn.lock().map_err(|e| e.to_string())?;
        
        for mut c in &mut candidates {
            let count: usize = conn.query_row(
                "SELECT COUNT(*) FROM projects WHERE name = ?1",
                [&c.folder_name],
                |r| r.get(0)
            ).map_err(|e| e.to_string())?;
            c.exists_conflict = count > 0;
        }

        Ok(candidates)
    }

    pub fn execute_bulk_import(&self, selections: Vec<ImportSelection>) -> Result<(), String> {
        let mut conn = self.conn.lock().map_err(|e| e.to_string())?;
        let tx = conn.transaction().map_err(|e| e.to_string())?;

        for sel in selections {
            let folder_path = Path::new(&sel.folder_path);
            if !folder_path.exists() || !folder_path.is_dir() {
                continue;
            }

            let folder_name = folder_path.file_name().unwrap_or_default().to_string_lossy().to_string();
            let mut final_name = folder_name.clone();
            
            let count: usize = tx.query_row(
                "SELECT COUNT(*) FROM projects WHERE name = ?1",
                [&final_name],
                |r| r.get(0)
            ).map_err(|e| e.to_string())?;

            let mut should_insert = true;
            let mut project_id = format!("proj_{}_{}", chrono::Utc::now().timestamp_millis(), calculate_hash(&sel.folder_path));

            if count > 0 {
                match sel.conflict_action.as_str() {
                    "merge" => {
                        let existing_id: String = tx.query_row(
                            "SELECT id FROM projects WHERE name = ?1",
                            [&final_name],
                            |r| r.get(0)
                        ).map_err(|e| e.to_string())?;
                        project_id = existing_id;
                        should_insert = false;
                        
                        tx.execute(
                            "UPDATE projects SET folder_path = ?1, updated_at = ?2 WHERE id = ?3",
                            rusqlite::params![sel.folder_path, chrono::Utc::now().to_rfc3339(), project_id],
                        ).map_err(|e| e.to_string())?;

                        tx.execute(
                            "DELETE FROM project_directories WHERE project_id = ?1",
                            rusqlite::params![project_id],
                        ).map_err(|e| e.to_string())?;

                        tx.execute(
                            "DELETE FROM project_files WHERE project_id = ?1 AND storage_mode = 'linked'",
                            rusqlite::params![project_id],
                        ).map_err(|e| e.to_string())?;
                    }
                    "skip" => {
                        continue;
                    }
                    "new" => {
                        let mut suffix = 1;
                        loop {
                            let candidate_name = format!("{}_{}", folder_name, suffix);
                            let sub_count: usize = tx.query_row(
                                "SELECT COUNT(*) FROM projects WHERE name = ?1",
                                [&candidate_name],
                                |r| r.get(0)
                            ).map_err(|e| e.to_string())?;
                            if sub_count == 0 {
                                final_name = candidate_name;
                                break;
                            }
                            suffix += 1;
                        }
                    }
                    _ => return Err("未知的冲突处理动作".to_string()),
                }
            }

            if should_insert {
                let now = chrono::Utc::now().to_rfc3339();
                tx.execute(
                    "INSERT INTO projects (id, name, customer_name, status, benefit_status, total_revenue_incl, total_cost_incl, project_years, discount_rate, cashflow_model, created_at, updated_at, folder_path, logs) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10, ?11, ?12, ?13, ?14)",
                    rusqlite::params![
                        project_id,
                        final_name,
                        "CMCC",
                        "立项中",
                        "未测算",
                        0.0,
                        0.0,
                        10,
                        0.06,
                        "default",
                        now,
                        now,
                        sel.folder_path,
                        "[]"
                    ]
                ).map_err(|e| e.to_string())?;
            }

            // Find matching root
            let root_opt: Option<(String, String)> = {
                let mut stmt = tx.prepare("SELECT id, root_path FROM project_roots").map_err(|e| e.to_string())?;
                let root_iter = stmt.query_map([], |r| {
                    Ok((r.get::<_, String>(0)?, r.get::<_, String>(1)?))
                }).map_err(|e| e.to_string())?;
                let target_norm = sel.folder_path.replace("\\", "/").to_lowercase();
                
                let mut matched = None;
                let mut longest_match_len = 0;
                for r in root_iter {
                    let (id, root_path) = r.map_err(|e| e.to_string())?;
                    let root_norm = root_path.replace("\\", "/").to_lowercase();
                    if target_norm.starts_with(&root_norm) && root_norm.len() > longest_match_len {
                        longest_match_len = root_norm.len();
                        let rel = &sel.folder_path[root_path.len()..];
                        let rel_clean = rel.trim_start_matches('\\').trim_start_matches('/').to_string();
                        matched = Some((id, rel_clean));
                    }
                }
                matched
            };

            let (root_id, relative_path) = match root_opt {
                Some((rid, rel)) => (Some(rid), Some(rel)),
                None => (None, None),
            };

            if let (Some(rid), Some(rel)) = (&root_id, &relative_path) {
                let dir_id = format!("dir_{}", calculate_hash(&(project_id.clone(), &sel.folder_path)));
                let now = chrono::Utc::now().to_rfc3339();
                tx.execute(
                    "INSERT OR REPLACE INTO project_directories (id, project_id, root_id, relative_path, dir_name, created_at, updated_at) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7)",
                    rusqlite::params![dir_id, project_id, rid, rel, final_name, now, now]
                ).map_err(|e| e.to_string())?;
            }

            let candidate_files = get_candidate_files(folder_path)?;
            let dir_id = if root_id.is_some() {
                Some(format!("dir_{}", calculate_hash(&(project_id.clone(), &sel.folder_path))))
            } else {
                None
            };
            
            for file in candidate_files {
                let now = chrono::Utc::now().to_rfc3339();
                let file_path_norm = file.path.replace("\\", "/");
                let folder_path_norm = sel.folder_path.replace("\\", "/");
                let file_rel = if file_path_norm.starts_with(&folder_path_norm) {
                    let sub = &file.path[sel.folder_path.len()..];
                    sub.trim_start_matches('\\').trim_start_matches('/').to_string()
                } else {
                    "".to_string()
                };

                let relative_path_val = if let Some(ref base_rel) = relative_path {
                    if base_rel.is_empty() {
                        Some(file_rel.clone())
                    } else {
                        Some(format!("{}/{}", base_rel, file_rel).replace("\\", "/"))
                    }
                } else {
                    None
                };

                let (size, modified_time) = if let Ok(meta) = fs::metadata(&file.path) {
                    let size = meta.len();
                    let modified: DateTime<Utc> = meta.modified()
                        .map(DateTime::from)
                        .unwrap_or_else(|_| Utc::now());
                    (size, modified.to_rfc3339())
                } else {
                    (0, now.clone())
                };

                let file_hash = match calculate_lightweight_hash(&file.path) {
                    Ok(h) => Some(h),
                    Err(_) => None,
                };

                let hash_val = calculate_hash(&(project_id.clone(), &file.path));
                let file_id = format!("{}-{:x}", project_id, hash_val);
                let ext = Path::new(&file.path).extension().unwrap_or_default().to_string_lossy().to_string();

                let file_type = match ext.to_lowercase().as_str() {
                    "doc" | "docx" => "word",
                    "xls" | "xlsx" => "excel",
                    "pdf" => "pdf",
                    "ppt" | "pptx" => "ppt",
                    _ => "other",
                };

                tx.execute(
                    "INSERT OR REPLACE INTO project_files (id, project_id, file_name, file_path, original_path, managed_path, file_type, extension, size, \"exists\", last_scanned_at, modified_at, storage_mode, is_main_document, is_main_budget_file, note, created_at, updated_at, root_id, directory_id, relative_path, absolute_path_snapshot, file_hash, file_role) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10, ?11, ?12, ?13, ?14, ?15, ?16, ?17, ?18, ?19, ?20, ?21, ?22, ?23, ?24)",
                    rusqlite::params![
                        file_id,
                        project_id,
                        file.name,
                        file.path.clone(),
                        Option::<String>::None,
                        Option::<String>::None,
                        file_type,
                        ext,
                        size,
                        1,
                        Some(now.clone()),
                        modified_time,
                        "linked",
                        0,
                        0,
                        Option::<String>::None,
                        now.clone(),
                        now,
                        root_id.clone(),
                        dir_id.clone(),
                        relative_path_val,
                        Some(file.path),
                        file_hash,
                        Some(file.file_role)
                    ]
                ).map_err(|e| e.to_string())?;
            }
        }

        tx.commit().map_err(|e| e.to_string())?;
        Ok(())
    }
}

// Helpers
fn scan_for_candidates(parent_path: &str) -> Result<Vec<ImportCandidate>, String> {
    let parent = Path::new(parent_path);
    if !parent.exists() || !parent.is_dir() {
        return Err("指定的路径不存在或不是一个有效的目录".to_string());
    }

    let mut candidates = Vec::new();
    let entries = fs::read_dir(parent).map_err(|e| e.to_string())?;
    
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_dir() {
            if let Some(name) = path.file_name() {
                let name_str = name.to_string_lossy();
                if name_str.starts_with('.') {
                    continue;
                }
                
                let candidate_files = get_candidate_files(&path)?;
                if !candidate_files.is_empty() {
                    candidates.push(ImportCandidate {
                        folder_name: name_str.to_string(),
                        folder_path: path.to_string_lossy().to_string(),
                        exists_conflict: false,
                        files: candidate_files,
                    });
                }

                let sub_entries = fs::read_dir(&path).map_err(|e| e.to_string())?;
                for sub_entry in sub_entries.flatten() {
                    let sub_path = sub_entry.path();
                    if sub_path.is_dir() {
                        if let Some(sub_name) = sub_path.file_name() {
                            let sub_name_str = sub_name.to_string_lossy();
                            if sub_name_str.starts_with('.') {
                                continue;
                            }
                            let sub_candidate_files = get_candidate_files(&sub_path)?;
                            if !sub_candidate_files.is_empty() {
                                candidates.push(ImportCandidate {
                                    folder_name: format!("{}/{}", name_str, sub_name_str),
                                    folder_path: sub_path.to_string_lossy().to_string(),
                                    exists_conflict: false,
                                    files: sub_candidate_files,
                                });
                            }
                        }
                    }
                }
            }
        }
    }

    Ok(candidates)
}

fn get_candidate_files(dir: &Path) -> Result<Vec<CandidateFile>, String> {
    let mut files = Vec::new();
    let entries = fs::read_dir(dir).map_err(|e| e.to_string())?;
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_file() {
            let name = path.file_name().unwrap_or_default().to_string_lossy().to_string();
            if name.starts_with("~$") || name.starts_with('.') {
                continue;
            }
            let ext = path.extension().unwrap_or_default().to_string_lossy().to_lowercase();
            if ext == "docx" || ext == "xlsx" || ext == "pdf" || ext == "pptx" {
                let file_role = detect_file_role(&name);
                files.push(CandidateFile {
                    name,
                    path: path.to_string_lossy().to_string(),
                    file_role,
                });
            }
        }
    }
    Ok(files)
}

fn detect_file_role(file_name: &str) -> String {
    let lower = file_name.to_lowercase();
    if lower.contains("效益分析") || lower.contains("效益") || lower.contains("benefit") {
        "benefit_scheme".to_string()
    } else if lower.contains("预算") || lower.contains("报价") || lower.contains("budget") || lower.contains("cost") {
        "budget".to_string()
    } else if lower.contains("方案") || lower.contains("规划") || lower.contains("proposal") || lower.contains("design") {
        "proposal".to_string()
    } else {
        "other".to_string()
    }
}

fn calculate_lightweight_hash(path_str: &str) -> Result<String, std::io::Error> {
    use std::io::Read;
    let path = Path::new(path_str);
    let metadata = fs::metadata(path)?;
    let size = metadata.len();
    
    let modified = metadata.modified()
        .map(|t| t.duration_since(std::time::UNIX_EPOCH).unwrap_or_default().as_secs())
        .unwrap_or(0);

    let mut file = fs::File::open(path)?;
    let mut buffer = [0u8; 8192];
    let bytes_read = file.read(&mut buffer)?;
    
    let mut hasher = DefaultHasher::new();
    hasher.write(&buffer[..bytes_read]);
    let content_hash = hasher.finish();

    Ok(format!("{}:{}:{:x}", size, modified, content_hash))
}

fn calculate_hash<T: Hash>(t: &T) -> u64 {
    let mut s = DefaultHasher::new();
    t.hash(&mut s);
    s.finish()
}

// Tauri commands
#[tauri::command]
pub async fn scan_import_candidates(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    parent_path: String,
) -> Result<Vec<ImportCandidate>, String> {
    let service = ImportScanner::new(runtime.require_db()?);
    service.scan_import_candidates(&parent_path)
}

#[tauri::command]
pub async fn execute_bulk_import(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    selections: Vec<ImportSelection>,
) -> Result<(), String> {
    let service = ImportScanner::new(runtime.require_db()?);
    service.execute_bulk_import(selections)
}
