use super::models::ProjectFile;
use super::repository::ProjectFileRepository;
use super::scanner;
use chrono::{DateTime, Utc};
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::Arc;
use tauri::Manager;

pub struct ProjectFileService {
    repository: Arc<dyn ProjectFileRepository + Send + Sync>,
}

impl ProjectFileService {
    pub fn new(repository: Arc<dyn ProjectFileRepository + Send + Sync>) -> Self {
        Self { repository }
    }

    pub fn get_project_files(&self, project_id: &str) -> Result<Vec<ProjectFile>, String> {
        self.repository.get_project_files(project_id)
    }

    pub fn bind_project_folder(
        &self,
        project_id: &str,
        folder_path: &str,
        force_mode: Option<String>,
    ) -> Result<(), String> {
        let path = Path::new(folder_path);
        if !path.exists() || !path.is_dir() {
            return Err("指定的路径不存在或不是一个有效的文件夹".to_string());
        }

        let folder_name = path
            .file_name()
            .unwrap_or_default()
            .to_string_lossy()
            .to_string();

        let old_folder_path = self.get_project_folder_path(project_id)?;
        if old_folder_path.as_deref() != Some(folder_path) {
            self.clear_previous_folder_links(project_id, old_folder_path.as_deref())?;
            self.repository.delete_project_directory(project_id)?;
        }

        // Find if folder belongs to any registered project roots
        let matched = self.repository.find_matching_root(folder_path)?;

        let (root_id, relative_path) = match matched {
            Some((rid, rel)) => (Some(rid), Some(rel)),
            None => {
                match force_mode.as_deref() {
                    Some("create_root") => {
                        let (root_path_str, root_name_str, rel_path_str) = if let Some(parent) = path.parent() {
                            let p_str = parent.to_string_lossy().to_string();
                            let p_name = parent.file_name()
                                .map(|n| n.to_string_lossy().to_string())
                                .unwrap_or_else(|| p_str.clone());
                            (p_str, p_name, folder_name.clone())
                        } else {
                            (folder_path.to_string(), folder_name.clone(), "".to_string())
                        };
                        let new_rid = self.repository.create_root_direct(&root_name_str, &root_path_str)?;
                        (Some(new_rid), Some(rel_path_str))
                    }
                    Some("absolute_only") => {
                        (None, None)
                    }
                    _ => {
                        return Err("NOT_IN_ROOT".to_string());
                    }
                }
            }
        };

        if let (Some(rid), Some(rel)) = (&root_id, &relative_path) {
            let dir_id = format!("dir_{}", calculate_hash(&(project_id, folder_path)));
            self.repository.save_project_directory(&dir_id, project_id, rid, rel, &folder_name)?;
        }

        // Update project folder path in repository
        self.repository.update_project_fields(
            project_id,
            Some(folder_path.to_string()),
            None,
            None,
        )?;

        // Auto scan after binding
        self.scan_project_folder(project_id, false)?;
        Ok(())
    }

    pub fn create_project_folder(parent_path: &str, folder_name: &str) -> Result<String, String> {
        let parent = Path::new(parent_path);
        if !parent.exists() || !parent.is_dir() {
            return Err("父级目录不存在或不是有效文件夹".to_string());
        }

        let clean_name = folder_name.trim();
        if clean_name.is_empty() {
            return Err("文件夹名称不能为空".to_string());
        }
        if clean_name.contains('/')
            || clean_name.contains('\\')
            || clean_name.contains(':')
            || clean_name.contains('*')
            || clean_name.contains('?')
            || clean_name.contains('"')
            || clean_name.contains('<')
            || clean_name.contains('>')
            || clean_name.contains('|')
        {
            return Err("文件夹名称包含非法字符".to_string());
        }

        let new_folder = parent.join(clean_name);
        if new_folder.exists() {
            return Err(format!("文件夹已存在: {}", new_folder.display()));
        }

        fs::create_dir_all(&new_folder).map_err(|e| format!("创建项目文件夹失败: {}", e))?;

        Ok(new_folder.to_string_lossy().to_string())
    }

    pub fn unbind_project_folder(&self, project_id: &str) -> Result<(), String> {
        // Clear project fields
        self.repository.update_project_fields(
            project_id,
            Some("".to_string()),
            Some("".to_string()),
            Some("".to_string()),
        )?;

        // Remove linked file records
        let files = self.repository.get_project_files(project_id)?;
        for file in files {
            if file.storage_mode == "linked" {
                self.repository.delete_file(&file.id)?;
            }
        }

        // Delete project directory records
        self.repository.delete_project_directory(project_id)?;

        Ok(())
    }

    pub fn scan_project_folder(
        &self,
        project_id: &str,
        recursive: bool,
    ) -> Result<Vec<ProjectFile>, String> {
        let folder_path_opt = self.get_project_folder_path(project_id)?;
        let folder_path = match folder_path_opt {
            Some(path) => path,
            None => return Err("该项目未绑定任何文件夹".to_string()),
        };

        // Determine if there is a project root and relative path matching the folder
        let matched = self.repository.find_matching_root(&folder_path)?;
        let (root_id, base_relative_path) = match matched {
            Some((rid, rel)) => (Some(rid), Some(rel)),
            None => (None, None),
        };

        // If matched a root, dynamically ensure the directory record exists in DB
        if let (Some(ref rid), Some(ref base_rel)) = (&root_id, &base_relative_path) {
            let dir_id = format!("dir_{}", calculate_hash(&(project_id, &folder_path)));
            let folder_name = Path::new(&folder_path)
                .file_name()
                .unwrap_or_default()
                .to_string_lossy()
                .to_string();
            self.repository.save_project_directory(&dir_id, project_id, rid, base_rel, &folder_name)?;
        }

        let scanned_files = scanner::scan_directory(project_id, &folder_path, recursive)?;
        let mut existing_files = self.repository.get_project_files(project_id)?;
        let now = Utc::now().to_rfc3339();

        let mut files_to_save = Vec::new();

        // 1. Process scanned files
        for mut scanned in scanned_files {
            scanned.absolute_path_snapshot = Some(scanned.file_path.clone());
            
            // Calculate relative path for this file
            if let (Some(ref rid), Some(ref base_rel)) = (&root_id, &base_relative_path) {
                scanned.root_id = Some(rid.clone());
                let file_path_norm = scanned.file_path.replace("\\", "/");
                let folder_path_norm = folder_path.replace("\\", "/");
                if file_path_norm.starts_with(&folder_path_norm) {
                    let sub_rel = &scanned.file_path[folder_path.len()..];
                    let sub_rel_clean = sub_rel.trim_start_matches('\\').trim_start_matches('/');
                    let relative_path_val = if base_rel.is_empty() {
                        sub_rel_clean.to_string()
                    } else if sub_rel_clean.is_empty() {
                        base_rel.clone()
                    } else {
                        format!("{}/{}", base_rel, sub_rel_clean).replace("\\", "/")
                    };
                    scanned.relative_path = Some(relative_path_val);
                }
            }

            if root_id.is_some() {
                scanned.directory_id = Some(format!("dir_{}", calculate_hash(&(project_id, &folder_path))));
            } else {
                scanned.directory_id = None;
            }

            // Generate lightweight file hash
            scanned.file_hash = match calculate_lightweight_hash(&scanned.file_path) {
                Ok(h) => Some(h),
                Err(_) => None,
            };

            // Detect file role
            scanned.file_role = Some(detect_file_role(&scanned.file_name));

            if let Some(existing_idx) = existing_files
                .iter()
                .position(|f| f.file_path == scanned.file_path)
            {
                // Update existing record
                let mut existing = existing_files.remove(existing_idx);
                existing.size = scanned.size;
                existing.modified_at = scanned.modified_at;
                existing.exists = true;
                existing.last_scanned_at = Some(now.clone());
                existing.updated_at = now.clone();
                
                // Copy Phase 2 fields
                existing.root_id = scanned.root_id;
                existing.directory_id = scanned.directory_id;
                existing.relative_path = scanned.relative_path;
                existing.absolute_path_snapshot = scanned.absolute_path_snapshot;
                existing.file_hash = scanned.file_hash;
                existing.file_role = scanned.file_role;

                files_to_save.push(existing);
            } else {
                // Insert new linked file record
                scanned.last_scanned_at = Some(now.clone());
                files_to_save.push(scanned);
            }
        }

        // 2. Any remaining files in existing_files that have storage_mode == "linked"
        // and are located in the folder_path but weren't found in scanned files:
        // mark them as exists = false, last_scanned_at = now.
        for mut remaining in existing_files {
            if remaining.storage_mode == "linked" && remaining.file_path.starts_with(&folder_path) {
                remaining.exists = false;
                remaining.last_scanned_at = Some(now.clone());
                remaining.updated_at = now.clone();
                files_to_save.push(remaining);
            } else {
                // Keep copied files or files outside the folder intact
                files_to_save.push(remaining);
            }
        }

        self.repository.save_files(&files_to_save)?;
        self.repository.get_project_files(project_id)
    }

    fn clear_previous_folder_links(
        &self,
        project_id: &str,
        old_folder_path: Option<&str>,
    ) -> Result<(), String> {
        let files = self.repository.get_project_files(project_id)?;
        let old_folder = old_folder_path.map(PathBuf::from);

        for file in files {
            if file.storage_mode != "linked" {
                continue;
            }

            let is_scanned_folder_record =
                file.original_path.is_none() && file.managed_path.is_none();
            let is_inside_old_folder = old_folder
                .as_ref()
                .map(|old| Path::new(&file.file_path).starts_with(old))
                .unwrap_or(false);

            if is_scanned_folder_record || is_inside_old_folder {
                self.repository.delete_file(&file.id)?;
            }
        }

        Ok(())
    }

    pub fn add_project_file(
        &self,
        app_handle: &tauri::AppHandle,
        project_id: &str,
        src_path: &str,
        storage_mode: &str,
    ) -> Result<ProjectFile, String> {
        let src_file_path = Path::new(src_path);
        if !src_file_path.exists() || !src_file_path.is_file() {
            return Err("源文件不存在或不是有效的文件".to_string());
        }

        let file_name = src_file_path
            .file_name()
            .unwrap_or_default()
            .to_string_lossy()
            .to_string();

        let ext = src_file_path
            .extension()
            .unwrap_or_default()
            .to_string_lossy()
            .to_lowercase();

        let file_type = match ext.as_str() {
            "doc" | "docx" => "word",
            "xls" | "xlsx" => "excel",
            "pdf" => "pdf",
            "ppt" | "pptx" => "ppt",
            "png" | "jpg" | "jpeg" | "gif" | "bmp" => "image",
            _ => "other",
        };

        let metadata =
            fs::metadata(src_file_path).map_err(|e| format!("无法读取文件属性: {}", e))?;
        let size = metadata.len();
        let modified: DateTime<Utc> = metadata
            .modified()
            .map(chrono::DateTime::from)
            .unwrap_or_else(|_| Utc::now());

        let now = Utc::now().to_rfc3339();

        let mut project_file = ProjectFile {
            id: "".to_string(), // Generated below
            project_id: project_id.to_string(),
            file_name: file_name.clone(),
            file_path: "".to_string(), // Filled below
            original_path: Some(src_path.to_string()),
            managed_path: None,
            file_type: file_type.to_string(),
            extension: ext,
            size,
            exists: true,
            last_scanned_at: Some(now.clone()),
            modified_at: modified.to_rfc3339(),
            storage_mode: storage_mode.to_string(),
            is_main_document: false,
            is_main_budget_file: false,
            note: None,
            created_at: now.clone(),
            updated_at: now,
            root_id: None,
            directory_id: None,
            relative_path: None,
            absolute_path_snapshot: Some(src_path.to_string()),
            file_hash: match calculate_lightweight_hash(src_path) {
                Ok(h) => Some(h),
                Err(_) => None,
            },
            file_role: Some(detect_file_role(&file_name)),
        };

        if storage_mode == "copied" {
            let app_data_dir = app_handle
                .path()
                .app_data_dir()
                .map_err(|e| format!("无法获取App数据目录: {}", e))?;

            let dest_dir = app_data_dir.join("projects").join(project_id).join("files");

            if !dest_dir.exists() {
                fs::create_dir_all(&dest_dir).map_err(|e| format!("无法创建托管文件夹: {}", e))?;
            }

            let dest_file_path = dest_dir.join(&file_name);
            fs::copy(src_file_path, &dest_file_path)
                .map_err(|e| format!("复制托管文件失败: {}", e))?;

            let final_path = dest_file_path.to_string_lossy().to_string();
            project_file.file_path = final_path.clone();
            project_file.managed_path = Some(final_path);
        } else {
            project_file.file_path = src_path.to_string();
            if let Ok(Some((rid, rel))) = self.repository.find_matching_root(src_path) {
                project_file.root_id = Some(rid);
                project_file.relative_path = Some(rel);
            }
            if let Ok(Some(folder_path)) = self.get_project_folder_path(project_id) {
                if let Ok(Some((rid, rel))) = self.repository.find_matching_root(&folder_path) {
                    let dir_id = format!("dir_{}", calculate_hash(&(project_id, &folder_path)));
                    let folder_name = Path::new(&folder_path)
                        .file_name()
                        .unwrap_or_default()
                        .to_string_lossy()
                        .to_string();
                    let _ = self.repository.save_project_directory(&dir_id, project_id, &rid, &rel, &folder_name);
                    project_file.directory_id = Some(dir_id);
                } else {
                    project_file.directory_id = None;
                }
            }
        }

        // Generate deterministic ID to avoid duplicates
        let hash_val = calculate_hash(&(project_id, &project_file.file_path));
        project_file.id = format!("{}-{:x}", project_id, hash_val);

        self.repository.save_file(&project_file)?;
        Ok(project_file)
    }

    pub fn remove_project_file_record(
        &self,
        project_id: &str,
        file_id: &str,
    ) -> Result<(), String> {
        let file_opt = self.repository.get_file(file_id)?;
        let file = match file_opt {
            Some(f) => f,
            None => return Err("未找到指定的文件记录".to_string()),
        };

        if file.project_id != project_id {
            return Err("文件记录不属于该项目".to_string());
        }

        // Check if this was a main document or main budget file and clear those flags in the project
        if file.is_main_document {
            self.repository
                .update_project_fields(project_id, None, Some("".to_string()), None)?;
        }
        if file.is_main_budget_file {
            self.repository
                .update_project_fields(project_id, None, None, Some("".to_string()))?;
        }

        self.repository.delete_file(file_id)
    }

    pub fn delete_managed_project_file(
        &self,
        project_id: &str,
        file_id: &str,
    ) -> Result<(), String> {
        let file_opt = self.repository.get_file(file_id)?;
        let file = match file_opt {
            Some(f) => f,
            None => return Err("未找到指定的文件记录".to_string()),
        };

        if file.project_id != project_id {
            return Err("文件记录不属于该项目".to_string());
        }

        // Delete physical file if it was copied/managed
        if file.storage_mode == "copied" {
            let path = Path::new(&file.file_path);
            if path.exists() {
                fs::remove_file(path).map_err(|e| format!("物理删除托管文件失败: {}", e))?;
            }
        }

        // Clear main indicators if necessary
        if file.is_main_document {
            self.repository
                .update_project_fields(project_id, None, Some("".to_string()), None)?;
        }
        if file.is_main_budget_file {
            self.repository
                .update_project_fields(project_id, None, None, Some("".to_string()))?;
        }

        self.repository.delete_file(file_id)
    }

    pub fn mark_main_document(
        &self,
        project_id: &str,
        file_id: Option<&str>,
    ) -> Result<(), String> {
        let all_files = self.repository.get_project_files(project_id)?;
        let mut files_to_save = Vec::new();
        let mut target_path = None;

        for mut file in all_files {
            if let Some(fid) = file_id {
                if file.id == fid {
                    file.is_main_document = true;
                    target_path = Some(file.file_path.clone());
                } else {
                    file.is_main_document = false;
                }
            } else {
                file.is_main_document = false;
            }
            files_to_save.push(file);
        }

        self.repository.save_files(&files_to_save)?;

        // Update Project main document path
        self.repository.update_project_fields(
            project_id,
            None,
            Some(target_path.unwrap_or_default()),
            None,
        )?;
        Ok(())
    }

    pub fn mark_main_budget_file(
        &self,
        project_id: &str,
        file_id: Option<&str>,
    ) -> Result<(), String> {
        let all_files = self.repository.get_project_files(project_id)?;
        let mut files_to_save = Vec::new();
        let mut target_path = None;

        for mut file in all_files {
            if let Some(fid) = file_id {
                if file.id == fid {
                    file.is_main_budget_file = true;
                    target_path = Some(file.file_path.clone());
                } else {
                    file.is_main_budget_file = false;
                }
            } else {
                file.is_main_budget_file = false;
            }
            files_to_save.push(file);
        }

        self.repository.save_files(&files_to_save)?;

        // Update Project main budget path
        self.repository.update_project_fields(
            project_id,
            None,
            None,
            Some(target_path.unwrap_or_default()),
        )?;
        Ok(())
    }

    pub fn open_project_folder(&self, project_id: &str) -> Result<(), String> {
        let folder_opt = self.resolve_project_folder_path(project_id)?;
        let folder = match folder_opt {
            Some(path) => path,
            None => return Err("该项目未绑定任何文件夹".to_string()),
        };

        let path = Path::new(&folder);
        if !path.exists() {
            return Err("绑定的文件夹在磁盘中已不存在".to_string());
        }

        open_path(&folder)
    }

    pub fn open_project_file(&self, file_id: &str) -> Result<(), String> {
        let file_opt = self.repository.get_file(file_id)?;
        let mut file = match file_opt {
            Some(f) => f,
            None => return Err("未找到指定的文件记录".to_string()),
        };

        let resolved_path = self.resolve_file_path(&file)?;
        if resolved_path != file.file_path {
            file.file_path = resolved_path.clone();
            file.exists = true;
            file.updated_at = Utc::now().to_rfc3339();
            self.repository.save_file(&file)?;
        }

        open_path(&resolved_path)
    }

    pub fn reveal_project_file(&self, file_id: &str) -> Result<(), String> {
        let file_opt = self.repository.get_file(file_id)?;
        let mut file = match file_opt {
            Some(f) => f,
            None => return Err("未找到指定的文件记录".to_string()),
        };

        let resolved_path = self.resolve_file_path(&file)?;
        if resolved_path != file.file_path {
            file.file_path = resolved_path.clone();
            file.exists = true;
            file.updated_at = Utc::now().to_rfc3339();
            self.repository.save_file(&file)?;
        }

        reveal_path(&resolved_path)
    }

    pub fn resolve_file_path(&self, file: &ProjectFile) -> Result<String, String> {
        // Priority 1: Current root path + relative path
        if let (Some(ref root_id), Some(ref rel_path)) = (&file.root_id, &file.relative_path) {
            if let Ok(Some(root_path)) = self.repository.get_root_path(root_id) {
                let path = Path::new(&root_path).join(rel_path);
                if path.exists() {
                    return Ok(path.to_string_lossy().to_string());
                }
            }
        }

        // Priority 2: Absolute path snapshot
        if let Some(ref abs_path) = file.absolute_path_snapshot {
            let path = Path::new(abs_path);
            if path.exists() {
                return Ok(abs_path.clone());
            }
        }

        // Priority 3 & 4: Search across all registered project roots
        if let Some(found_path) = self.search_in_roots(&file.file_name, file.size, &file.modified_at, file.file_hash.as_deref()) {
            return Ok(found_path);
        }

        // Priority 5: User manual resolution is handled on frontend, so we return FILE_NOT_FOUND
        Err("FILE_NOT_FOUND".to_string())
    }

    pub fn resolve_project_folder_path(&self, project_id: &str) -> Result<Option<String>, String> {
        let folder_opt = self.get_project_folder_path(project_id)?;
        if let Some(ref folder) = folder_opt {
            if Path::new(folder).exists() {
                return Ok(Some(folder.clone()));
            }
        }

        // Try to resolve folder path using bound project directory
        if let Ok(Some((root_id, relative_path))) = self.repository.get_project_directory(project_id) {
            if let Ok(Some(root_path)) = self.repository.get_root_path(&root_id) {
                let path = Path::new(&root_path).join(&relative_path);
                if path.exists() {
                    let path_str = path.to_string_lossy().to_string();
                    self.repository.update_project_fields(project_id, Some(path_str.clone()), None, None)?;
                    return Ok(Some(path_str));
                }
            }
        }

        Ok(folder_opt)
    }

    fn search_in_roots(&self, target_name: &str, target_size: u64, target_modified: &str, target_hash: Option<&str>) -> Option<String> {
        let roots = self.repository.get_all_roots().ok()?;
        for root in roots {
            let root_path = Path::new(&root.1);
            if root_path.exists() && root_path.is_dir() {
                if let Some(path) = search_dir_for_match(root_path, target_name, target_size, target_modified, target_hash, 0) {
                    return Some(path);
                }
            }
        }
        None
    }

    fn get_project_folder_path(&self, project_id: &str) -> Result<Option<String>, String> {
        self.repository.get_project_folder(project_id)
    }
}

// Deterministic ID generator helper
use std::collections::hash_map::DefaultHasher;
use std::hash::{Hash, Hasher};

fn calculate_hash<T: Hash>(t: &T) -> u64 {
    let mut s = DefaultHasher::new();
    t.hash(&mut s);
    s.finish()
}

#[cfg(target_os = "windows")]
fn open_path(path: &str) -> Result<(), String> {
    std::process::Command::new("cmd")
        .args(&["/C", "start", "", path])
        .spawn()
        .map(|_| ())
        .map_err(|e| format!("Windows无法打开路径: {}", e))
}

#[cfg(any(target_os = "macos", target_os = "ios"))]
fn open_path(path: &str) -> Result<(), String> {
    std::process::Command::new("open")
        .arg(path)
        .spawn()
        .map(|_| ())
        .map_err(|e| format!("macOS无法打开路径: {}", e))
}

#[cfg(not(any(target_os = "windows", target_os = "macos", target_os = "ios")))]
fn open_path(path: &str) -> Result<(), String> {
    std::process::Command::new("xdg-open")
        .arg(path)
        .spawn()
        .map(|_| ())
        .map_err(|e| format!("Linux无法打开路径: {}", e))
}

#[cfg(target_os = "windows")]
fn reveal_path(path: &str) -> Result<(), String> {
    std::process::Command::new("explorer")
        .arg(format!("/select,\"{}\"", path))
        .spawn()
        .map(|_| ())
        .map_err(|e| format!("Windows无法定位文件: {}", e))
}

#[cfg(any(target_os = "macos", target_os = "ios"))]
fn reveal_path(path: &str) -> Result<(), String> {
    std::process::Command::new("open")
        .arg("-R")
        .arg(path)
        .spawn()
        .map(|_| ())
        .map_err(|e| format!("macOS无法定位文件: {}", e))
}

#[cfg(not(any(target_os = "windows", target_os = "macos", target_os = "ios")))]
fn reveal_path(path: &str) -> Result<(), String> {
    let p = Path::new(path);
    let parent = p.parent().unwrap_or(p);
    open_path(&parent.to_string_lossy())
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

fn search_dir_for_match(
    dir: &Path,
    target_name: &str,
    target_size: u64,
    target_modified: &str,
    target_hash: Option<&str>,
    depth: usize,
) -> Option<String> {
    if depth > 4 {
        return None;
    }
    let entries = fs::read_dir(dir).ok()?;
    let mut subdirs = Vec::new();

    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_file() {
            if let Some(file_name) = path.file_name() {
                if file_name.to_string_lossy() == target_name {
                    if let Ok(meta) = path.metadata() {
                        let size = meta.len();
                        if size == target_size {
                            let modified: DateTime<Utc> = meta.modified()
                                .map(DateTime::from)
                                .unwrap_or_else(|_| Utc::now());
                            if modified.to_rfc3339() == target_modified {
                                return Some(path.to_string_lossy().to_string());
                            }
                            if let Some(h) = target_hash {
                                if let Ok(computed_h) = calculate_lightweight_hash(&path.to_string_lossy()) {
                                    if computed_h == h {
                                        return Some(path.to_string_lossy().to_string());
                                    }
                                }
                            }
                        }
                    }
                }
            }
        } else if path.is_dir() {
            if let Some(name) = path.file_name() {
                if !name.to_string_lossy().starts_with('.') {
                    subdirs.push(path);
                }
            }
        }
    }

    for subdir in subdirs {
        if let Some(found) = search_dir_for_match(&subdir, target_name, target_size, target_modified, target_hash, depth + 1) {
            return Some(found);
        }
    }

    None
}
