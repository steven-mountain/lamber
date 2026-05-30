use base64::Engine;
use chrono::Utc;
use rusqlite::{params, Connection};
use std::collections::hash_map::DefaultHasher;
use std::fs;
use std::hash::{Hash, Hasher};
use std::path::{Path, PathBuf};
use tauri::AppHandle;
use tauri::Manager;

fn calculate_hash<T: Hash>(t: &T) -> u64 {
    let mut s = DefaultHasher::new();
    t.hash(&mut s);
    s.finish()
}

pub struct TemplateAsset {
    pub id: String,
    pub project_id: String,
    pub template_name: String,
    pub asset_type: String,
    pub usage: Option<String>,
    pub original_file_name: Option<String>,
    pub stored_file_name: String,
    pub relative_path: String,
    pub absolute_path_snapshot: String,
    pub mime_type: Option<String>,
    pub file_size: i64,
    pub width: Option<i32>,
    pub height: Option<i32>,
    pub file_hash: Option<String>,
    pub created_at: String,
    pub updated_at: String,
    pub deleted_at: Option<String>,
}

fn parse_base64_data(input: &str) -> Result<(String, Vec<u8>), String> {
    if input.starts_with("data:") {
        let comma_idx = input
            .find(',')
            .ok_or_else(|| "Invalid data URI format".to_string())?;
        let header = &input[..comma_idx];
        let base64_str = &input[comma_idx + 1..];

        let parts: Vec<&str> = header.split(';').collect();
        if parts.len() < 2 || parts[parts.len() - 1] != "base64" {
            return Err("Unsupported encoding or format".to_string());
        }
        let mime = parts[0].strip_prefix("data:").unwrap_or("").to_string();

        let decoded = base64::engine::general_purpose::STANDARD
            .decode(base64_str.trim())
            .map_err(|e| format!("Base64 decoding failed: {}", e))?;
        Ok((mime, decoded))
    } else {
        let decoded = base64::engine::general_purpose::STANDARD
            .decode(input.trim())
            .map_err(|e| format!("Base64 decoding failed: {}", e))?;

        let mime = if decoded.starts_with(&[137, 80, 78, 71, 13, 10, 26, 10]) {
            "image/png".to_string()
        } else if decoded.starts_with(&[0xff, 0xd8, 0xff]) {
            "image/jpeg".to_string()
        } else if decoded.starts_with(b"RIFF") && decoded.len() > 8 && &decoded[8..12] == b"WEBP" {
            "image/webp".to_string()
        } else {
            return Err("Unknown image format: only PNG, JPEG, and WEBP are supported".to_string());
        };
        Ok((mime, decoded))
    }
}

fn sanitize_folder_name(name: &str) -> String {
    name.replace(|c: char| !c.is_alphanumeric() && c != '_' && c != '-', "_")
}

fn get_project_folder_info_from_db(
    conn: &Connection,
    project_id: &str,
) -> Result<Option<(String, String)>, String> {
    let mut stmt = conn
        .prepare("SELECT folder_path, name FROM projects WHERE id = ?1")
        .map_err(|e| e.to_string())?;
    let mut rows = stmt.query([project_id]).map_err(|e| e.to_string())?;
    if let Some(row) = rows.next().map_err(|e| e.to_string())? {
        let path_opt: Option<String> = row.get(0).map_err(|e| e.to_string())?;
        let name: String = row.get(1).map_err(|e| e.to_string())?;
        if let Some(path) = path_opt {
            Ok(Some((path, name)))
        } else {
            Ok(None)
        }
    } else {
        Ok(None)
    }
}

pub fn save_template_asset_internal(
    app_handle: &AppHandle,
    conn: &Connection,
    workspace_root: &str,
    project_id: &str,
    template_name: &str,
    asset_type: &str,
    usage: Option<&str>,
    original_file_name: Option<&str>,
    base64_data: &str,
    width: Option<i32>,
    height: Option<i32>,
) -> Result<String, String> {
    let (mime_type, data_bytes) = parse_base64_data(base64_data)?;

    if data_bytes.len() > 20 * 1024 * 1024 {
        return Err("IMAGE_TOO_LARGE::图片大小不能超过 20MB".to_string());
    }

    let ext = match mime_type.as_str() {
        "image/png" => "png",
        "image/jpeg" | "image/jpg" => "jpg",
        "image/webp" => "webp",
        _ => return Err("UNSUPPORTED_MIME_TYPE::仅支持 PNG, JPEG, WEBP 格式图片".to_string()),
    };

    let now_str = Utc::now().to_rfc3339();
    let hash_input = format!("{}-{}-{}", project_id, now_str, data_bytes.len());
    let hash_val = calculate_hash(&hash_input);
    let asset_id = format!("asset_{:x}", hash_val);
    let stored_file_name = format!("{}.{}", asset_id, ext);

    let _app_data_dir = app_handle
        .path()
        .app_data_dir()
        .map_err(|e| format!("无法获取 App 数据目录: {}", e))?;
    let workspace_root = Path::new(workspace_root);

    // Check if the project is bound to a folder
    let folder_info_opt = get_project_folder_info_from_db(conn, project_id)?;
    let (dest_file_path, relative_path) = if let Some((ref folder_path, ref project_name)) =
        folder_info_opt
    {
        if !folder_path.trim().is_empty() {
            let sanitized = sanitize_folder_name(project_name);
            let folder_name = format!("{}-图片", sanitized);
            let project_dir = crate::workspace::resolve_workspace_path(workspace_root, folder_path);
            let use_project_dir =
                crate::workspace::is_inside_workspace(workspace_root, &project_dir);
            let assets_dir = if use_project_dir {
                project_dir.join(&folder_name).join("assets")
            } else {
                workspace_root
                    .join(".projects")
                    .join(project_id)
                    .join("assets")
            };
            if !assets_dir.exists() {
                fs::create_dir_all(&assets_dir)
                    .map_err(|e| format!("创建项目嵌套资产目录失败: {}", e))?;
            }
            let dest = assets_dir.join(&stored_file_name);
            let rel = if use_project_dir {
                crate::workspace::to_relative_workspace_path(workspace_root, &dest)
            } else {
                format!(".projects/{}/assets/{}", project_id, stored_file_name)
            };
            (dest, rel)
        } else {
            let assets_dir = workspace_root
                .join(".projects")
                .join(project_id)
                .join("assets");
            if !assets_dir.exists() {
                fs::create_dir_all(&assets_dir).map_err(|e| format!("创建资产目录失败: {}", e))?;
            }
            let dest = assets_dir.join(&stored_file_name);
            let rel = format!(".projects/{}/assets/{}", project_id, stored_file_name);
            (dest, rel)
        }
    } else {
        let assets_dir = workspace_root
            .join(".projects")
            .join(project_id)
            .join("assets");
        if !assets_dir.exists() {
            fs::create_dir_all(&assets_dir).map_err(|e| format!("创建资产目录失败: {}", e))?;
        }
        let dest = assets_dir.join(&stored_file_name);
        let rel = format!(".projects/{}/assets/{}", project_id, stored_file_name);
        (dest, rel)
    };

    crate::workspace::mark_path_hidden_if_supported(&workspace_root.join(".projects"));

    let absolute_path_snapshot = dest_file_path.to_string_lossy().to_string();
    let file_hash = format!("{:x}", calculate_hash(&data_bytes));

    fs::write(&dest_file_path, &data_bytes).map_err(|e| format!("保存物理文件失败: {}", e))?;

    let result = conn.execute(
        "INSERT INTO project_template_assets (
            id, project_id, template_name, asset_type, usage, original_file_name, stored_file_name,
            relative_path, absolute_path_snapshot, mime_type, file_size, width, height, file_hash,
            created_at, updated_at, deleted_at
         ) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10, ?11, ?12, ?13, ?14, ?15, ?16, ?17)",
        params![
            asset_id,
            project_id,
            template_name,
            asset_type,
            usage,
            original_file_name,
            stored_file_name,
            relative_path,
            absolute_path_snapshot,
            Some(mime_type),
            data_bytes.len() as i64,
            width,
            height,
            Some(file_hash),
            now_str.clone(),
            now_str,
            None::<String>,
        ],
    );

    match result {
        Ok(_) => Ok(asset_id),
        Err(e) => {
            if dest_file_path.exists() {
                let _ = fs::remove_file(&dest_file_path);
            }
            Err(format!("保存资产元数据失败: {}", e))
        }
    }
}

pub fn get_template_assets_internal(
    conn: &Connection,
    project_id: &str,
    template_name: &str,
) -> Result<Vec<TemplateAsset>, String> {
    let mut stmt = conn
        .prepare("SELECT id, project_id, template_name, asset_type, usage, original_file_name, stored_file_name, relative_path, absolute_path_snapshot, mime_type, file_size, width, height, file_hash, created_at, updated_at, deleted_at FROM project_template_assets WHERE project_id = ?1 AND template_name = ?2 AND deleted_at IS NULL")
        .map_err(|e| e.to_string())?;
    let asset_iter = stmt
        .query_map([project_id, template_name], |row| {
            Ok(TemplateAsset {
                id: row.get(0)?,
                project_id: row.get(1)?,
                template_name: row.get(2)?,
                asset_type: row.get(3)?,
                usage: row.get(4)?,
                original_file_name: row.get(5)?,
                stored_file_name: row.get(6)?,
                relative_path: row.get(7)?,
                absolute_path_snapshot: row.get(8)?,
                mime_type: row.get(9)?,
                file_size: row.get(10)?,
                width: row.get(11)?,
                height: row.get(12)?,
                file_hash: row.get(13)?,
                created_at: row.get(14)?,
                updated_at: row.get(15)?,
                deleted_at: row.get(16)?,
            })
        })
        .map_err(|e| e.to_string())?;

    let mut list = Vec::new();
    for a in asset_iter {
        list.push(a.map_err(|e| e.to_string())?);
    }
    Ok(list)
}

pub fn delete_template_asset_internal(conn: &Connection, asset_id: &str) -> Result<(), String> {
    let now_str = Utc::now().to_rfc3339();
    conn.execute(
        "UPDATE project_template_assets SET deleted_at = ?1, updated_at = ?2 WHERE id = ?3",
        params![now_str.clone(), now_str, asset_id],
    )
    .map_err(|e| e.to_string())?;
    Ok(())
}

pub fn get_template_asset_path_internal(
    app_handle: &AppHandle,
    conn: &Connection,
    workspace_root: &str,
    asset_id: &str,
) -> Result<String, String> {
    let mut stmt = conn
        .prepare("SELECT project_id, relative_path, absolute_path_snapshot FROM project_template_assets WHERE id = ?1 AND deleted_at IS NULL")
        .map_err(|e| e.to_string())?;
    let mut rows = stmt.query([asset_id]).map_err(|e| e.to_string())?;
    if let Some(row) = rows.next().map_err(|e| e.to_string())? {
        let project_id: String = row.get(0).map_err(|e| e.to_string())?;
        let rel_path: String = row.get(1).map_err(|e| e.to_string())?;
        let abs_snap: String = row.get(2).map_err(|e| e.to_string())?;

        let app_data_dir = app_handle
            .path()
            .app_data_dir()
            .map_err(|e| format!("无法获取 App 数据目录: {}", e))?;
        let workspace_root = Path::new(workspace_root);

        let workspace_full_path = workspace_root.join(&rel_path);
        if workspace_full_path.exists() {
            return Ok(workspace_full_path.to_string_lossy().to_string());
        }

        // 1. Try resolving relative to the bound project folder if appropriate
        let folder_info_opt = get_project_folder_info_from_db(conn, &project_id)?;
        if let Some((ref folder_path, ref project_name)) = folder_info_opt {
            if !folder_path.trim().is_empty() {
                // Check direct path first
                let project_dir =
                    crate::workspace::resolve_workspace_path(workspace_root, folder_path);
                let full_path = project_dir.join(&rel_path);
                if full_path.exists() {
                    return Ok(full_path.to_string_lossy().to_string());
                }

                // Check current project name suffix path as fallback
                let sanitized = sanitize_folder_name(project_name);
                let folder_name = format!("{}-图片", sanitized);
                let suffix = if let Some(slash_idx) = rel_path.find('/') {
                    &rel_path[slash_idx + 1..]
                } else if let Some(backslash_idx) = rel_path.find('\\') {
                    &rel_path[backslash_idx + 1..]
                } else {
                    &rel_path
                };
                let fallback_path = project_dir.join(&folder_name).join(suffix);
                if fallback_path.exists() {
                    return Ok(fallback_path.to_string_lossy().to_string());
                }
            }
        }

        // 2. Try resolving relative to the current workspace
        let full_path = workspace_root.join(&rel_path);
        if full_path.exists() {
            return Ok(full_path.to_string_lossy().to_string());
        }
        let legacy_full_path = app_data_dir.join(&rel_path);
        if legacy_full_path.exists() {
            return Ok(legacy_full_path.to_string_lossy().to_string());
        }

        // 3. Fall back to absolute path snapshot
        let snap_path = if Path::new(&abs_snap).is_absolute() {
            PathBuf::from(&abs_snap)
        } else {
            workspace_root.join(&abs_snap)
        };
        if snap_path.exists() {
            return Ok(snap_path.to_string_lossy().to_string());
        }

        Err("物理图片文件已被删除或丢失".to_string())
    } else {
        Err("未找到指定的图片记录或已删除".to_string())
    }
}

pub fn cleanup_orphan_template_assets_internal(
    app_handle: &AppHandle,
    conn: &Connection,
    workspace_root: &str,
    project_id: &str,
) -> Result<(usize, Vec<String>), String> {
    let mut stmt = conn
        .prepare("SELECT value FROM project_settings WHERE project_id = ?1")
        .map_err(|e| e.to_string())?;
    let value_iter = stmt
        .query_map([project_id], |row| {
            let val: String = row.get(0)?;
            Ok(val)
        })
        .map_err(|e| e.to_string())?;

    let re = regex::Regex::new(r#""(asset_[0-9a-f]+)""#)
        .map_err(|e| format!("Failed to compile regex: {}", e))?;

    let mut active_assets = std::collections::HashSet::new();
    for val_res in value_iter {
        if let Ok(val) = val_res {
            for cap in re.captures_iter(&val) {
                if let Some(m) = cap.get(1) {
                    active_assets.insert(m.as_str().to_string());
                }
            }
        }
    }

    let mut stmt_assets = conn
        .prepare("SELECT id, relative_path, absolute_path_snapshot FROM project_template_assets WHERE project_id = ?1")
        .map_err(|e| e.to_string())?;

    let asset_iter = stmt_assets
        .query_map([project_id], |row| {
            let id: String = row.get(0)?;
            let rel: String = row.get(1)?;
            let abs: String = row.get(2)?;
            Ok((id, rel, abs))
        })
        .map_err(|e| e.to_string())?;

    let app_data_dir = app_handle
        .path()
        .app_data_dir()
        .map_err(|e| format!("无法获取 App 数据目录: {}", e))?;

    let folder_info_opt = get_project_folder_info_from_db(conn, project_id)?;
    let workspace_root_path = Path::new(workspace_root);
    let mut orphans_cleaned = 0;
    let mut cleaned_ids = Vec::new();

    for asset_res in asset_iter {
        if let Ok((id, rel, abs)) = asset_res {
            if !active_assets.contains(&id) {
                // Determine physical path using the same self-adaptive rules
                let mut deleted_physical = false;
                if let Some((ref folder_path, ref project_name)) = folder_info_opt {
                    if !folder_path.trim().is_empty() {
                        let mut target_path = None;

                        // 1. Direct path check
                        let project_dir = crate::workspace::resolve_workspace_path(
                            workspace_root_path,
                            folder_path,
                        );
                        let p1 = project_dir.join(&rel);
                        if p1.exists() {
                            target_path = Some(p1);
                        } else {
                            // 2. Fallback path check
                            let sanitized = sanitize_folder_name(project_name);
                            let folder_name = format!("{}-图片", sanitized);
                            let suffix = if let Some(slash_idx) = rel.find('/') {
                                &rel[slash_idx + 1..]
                            } else if let Some(backslash_idx) = rel.find('\\') {
                                &rel[backslash_idx + 1..]
                            } else {
                                &rel
                            };
                            let p2 = project_dir.join(&folder_name).join(suffix);
                            if p2.exists() {
                                target_path = Some(p2);
                            }
                        }

                        if let Some(path) = target_path {
                            let _ = fs::remove_file(&path);
                            deleted_physical = true;
                        }
                    }
                }

                if !deleted_physical {
                    let full_path = workspace_root_path.join(&rel);
                    if full_path.exists() {
                        let _ = fs::remove_file(&full_path);
                    } else {
                        let legacy_full_path = app_data_dir.join(&rel);
                        if legacy_full_path.exists() {
                            let _ = fs::remove_file(&legacy_full_path);
                        }
                    }
                }

                let snap_path = if Path::new(&abs).is_absolute() {
                    PathBuf::from(&abs)
                } else {
                    workspace_root_path.join(&abs)
                };
                if snap_path.exists()
                    && crate::workspace::is_inside_workspace(workspace_root_path, &snap_path)
                {
                    let _ = fs::remove_file(&snap_path);
                }

                let _ = conn.execute(
                    "DELETE FROM project_template_assets WHERE id = ?1",
                    [id.clone()],
                );

                orphans_cleaned += 1;
                cleaned_ids.push(id);
            }
        }
    }

    Ok((orphans_cleaned, cleaned_ids))
}
