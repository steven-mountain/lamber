use super::models::ProjectFile;
use chrono::{DateTime, Utc};
use std::collections::hash_map::DefaultHasher;
use std::fs;
use std::hash::{Hash, Hasher};
use std::path::Path;

fn calculate_hash<T: Hash>(t: &T) -> u64 {
    let mut s = DefaultHasher::new();
    t.hash(&mut s);
    s.finish()
}

pub fn scan_directory(
    project_id: &str,
    folder_path: &str,
    recursive: bool,
) -> Result<Vec<ProjectFile>, String> {
    let mut files = Vec::new();
    let path = Path::new(folder_path);
    if !path.exists() || !path.is_dir() {
        return Err("文件夹不存在或不是有效的目录".to_string());
    }

    scan_dir_recursive(project_id, path, path, recursive, &mut files)?;
    Ok(files)
}

fn scan_dir_recursive(
    project_id: &str,
    base_path: &Path,
    current_path: &Path,
    recursive: bool,
    files: &mut Vec<ProjectFile>,
) -> Result<(), String> {
    let entries = fs::read_dir(current_path).map_err(|e| format!("无法读取目录: {}", e))?;
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_dir() {
            if recursive {
                // Skip hidden folders starting with .
                if let Some(name) = path.file_name() {
                    if name.to_string_lossy().starts_with('.') {
                        continue;
                    }
                }
                scan_dir_recursive(project_id, base_path, &path, recursive, files)?;
            }
        } else if path.is_file() {
            let file_name = path
                .file_name()
                .unwrap_or_default()
                .to_string_lossy()
                .to_string();

            // Skip MS Office lock files starting with ~$ or temporary files starting with .~ or hidden files starting with .
            if file_name.starts_with("~$")
                || file_name.starts_with(".~")
                || file_name.starts_with('.')
            {
                continue;
            }

            let metadata = entry
                .metadata()
                .map_err(|e| format!("无法读取元数据: {}", e))?;
            let size = metadata.len();
            let modified: DateTime<Utc> = metadata
                .modified()
                .map(DateTime::from)
                .unwrap_or_else(|_| Utc::now());

            let ext = path
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

            let file_path = path.to_string_lossy().to_string();
            let hash_val = calculate_hash(&(project_id, &file_path));
            let file_id = format!("{}-{:x}", project_id, hash_val);
            let now = Utc::now().to_rfc3339();

            files.push(ProjectFile {
                id: file_id,
                project_id: project_id.to_string(),
                file_name,
                file_path,
                original_path: None,
                managed_path: None,
                file_type: file_type.to_string(),
                extension: ext,
                size,
                exists: true,
                last_scanned_at: Some(now.clone()),
                modified_at: modified.to_rfc3339(),
                storage_mode: "linked".to_string(),
                is_main_document: false,
                is_main_budget_file: false,
                note: None,
                created_at: now.clone(),
                updated_at: now,
            });
        }
    }
    Ok(())
}
