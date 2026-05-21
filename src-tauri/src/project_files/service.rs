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

    pub fn bind_project_folder(&self, project_id: &str, folder_path: &str) -> Result<(), String> {
        let path = Path::new(folder_path);
        if !path.exists() || !path.is_dir() {
            return Err("指定的路径不存在或不是一个有效的文件夹".to_string());
        }

        let old_folder_path = self.get_project_folder_path(project_id)?;
        if old_folder_path.as_deref() != Some(folder_path) {
            self.clear_previous_folder_links(project_id, old_folder_path.as_deref())?;
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

        Ok(())
    }

    pub fn scan_project_folder(
        &self,
        project_id: &str,
        recursive: bool,
    ) -> Result<Vec<ProjectFile>, String> {
        // We need to retrieve the project's folder_path.
        // We can find it by getting the store data or querying via a project repo,
        // but since we read StoreData anyway inside the repository, let's check folder path.
        // For convenience, we can look up the project's folder_path directly by checking the store.
        // Let's implement it inside ProjectFileRepository or do a lookup.
        // Let's get the project by ID using a helper or querying the store.
        let _ = self.repository.get_project_files(project_id)?; // Wait, get_project_files just gets files.
                                                                // Let's find folder_path of the project.
                                                                // We can add a method to JsonProjectFileRepository to get the folder_path of a project, or we can just load store inside repository.
                                                                // Let's see: we have `update_project_fields` in ProjectFileRepository which accesses projects.
                                                                // Let's add `get_project_folder(&self, project_id: &str) -> Result<Option<String>, String>` to ProjectFileRepository!
                                                                // Wait, instead of modifying the trait again, let's check if we can query it or if we should add it.
                                                                // Adding `get_project_folder(&self, project_id: &str) -> Result<Option<String>, String>` to the repository is extremely clean.
                                                                // Let's modify repository.rs in a minute, but we can write it in service.rs first.
                                                                // Wait, let's see how repository is defined. In repository.rs we implemented it. We can add this method!
                                                                // Let's write the service code assuming the repository supports `get_project_folder(project_id)`.

        let folder_path_opt = self.get_project_folder_path(project_id)?;
        let folder_path = match folder_path_opt {
            Some(path) => path,
            None => return Err("该项目未绑定任何文件夹".to_string()),
        };

        let scanned_files = scanner::scan_directory(project_id, &folder_path, recursive)?;
        let mut existing_files = self.repository.get_project_files(project_id)?;
        let now = Utc::now().to_rfc3339();

        let mut files_to_save = Vec::new();

        // 1. Process scanned files
        for mut scanned in scanned_files {
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
        let folder_opt = self.get_project_folder_path(project_id)?;
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
        let file = match file_opt {
            Some(f) => f,
            None => return Err("未找到指定的文件记录".to_string()),
        };

        let path = Path::new(&file.file_path);
        if !path.exists() {
            return Err("文件在磁盘中已不存在或已被移位".to_string());
        }

        open_path(&file.file_path)
    }

    pub fn reveal_project_file(&self, file_id: &str) -> Result<(), String> {
        let file_opt = self.repository.get_file(file_id)?;
        let file = match file_opt {
            Some(f) => f,
            None => return Err("未找到指定的文件记录".to_string()),
        };

        let path = Path::new(&file.file_path);
        if !path.exists() {
            return Err("文件在磁盘中已不存在".to_string());
        }

        reveal_path(&file.file_path)
    }

    // Helper to get folder path of a project
    fn get_project_folder_path(&self, project_id: &str) -> Result<Option<String>, String> {
        // Query the JSON file directly using our read_store wrapper (since it's a JsonRepository)
        // We can expose it or read it.
        // Let's read the store directly by accessing the repository's internal files.
        // Wait, since JsonProjectFileRepository reads StoreData, we can let ProjectFileRepository
        // expose it, or we can use our repository's JsonStore wrapper if we cast it or add it.
        // Let's add a function to ProjectFileRepository trait to query this!
        // That is extremely clean:
        // `fn get_project_folder(&self, project_id: &str) -> Result<Option<String>, String>;`
        // Let's implement it in JsonProjectFileRepository and use it!
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
