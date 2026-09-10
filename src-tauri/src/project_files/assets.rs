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

/// A replacement is one asset transaction, never a UI save followed by a delete.
/// The selected old asset must still be live; concurrent replacement fails closed.
pub fn validate_replacement_image(data_url: &str, width: i32, height: i32) -> Result<(), String> {
    // The UI fully decodes the selected file before preview. Reject empty/spoofed
    // payloads here too, before an IPC request can remove the old asset.
    let (mime, bytes) = parse_base64_data(data_url)?;
    let matches = match mime.as_str() {
        "image/png" => bytes.starts_with(&[137,80,78,71,13,10,26,10]),
        "image/jpeg" => bytes.starts_with(&[0xff,0xd8,0xff]),
        "image/webp" => bytes.starts_with(b"RIFF") && bytes.get(8..12) == Some(b"WEBP"),
        _ => false,
    };
    if !matches || bytes.len() > 20 * 1024 * 1024 || width <= 0 || height <= 0 {
        return Err("替换图片格式、大小或尺寸无效，原图片保持不变".into());
    }
    Ok(())
}

pub fn replace_demand_image_internal(
    conn: &Connection, workspace_root: &str, project_id: &str, template_name: &str, old_id: &str,
    save: impl FnOnce(&Connection, &str) -> Result<String, String>,
) -> Result<String, String> {
    let template = crate::agent_bridge::template_catalog::resolve(template_name, false)?;
    if template.id != "demand" { return Err("本轮图片替换仅支持需求导入表附件".into()); }
    let tx = conn.unchecked_transaction().map_err(|e| e.to_string())?;
    let (usage, created): (String, String) = tx.query_row(
        "SELECT usage, created_at FROM project_template_assets WHERE id=?1 AND project_id=?2 AND template_name=?3 AND asset_type='image' AND deleted_at IS NULL",
        params![old_id, project_id, template_name], |r| Ok((r.get(0)?, r.get(1)?)),
    ).map_err(|_| "原图片已变更、删除或不属于此项目模板，请重新选择")?;
    if !matches!(usage.as_str(), "attach1" | "attach2") { return Err("此图片不是可替换的需求表附件".into()); }
    let new_id = save(&tx, &usage)?;
    let relative: String = tx.query_row("SELECT relative_path FROM project_template_assets WHERE id=?1", [&new_id], |r| r.get(0)).map_err(|e| e.to_string())?;
    let result = (|| {
        // Preserve document order; the old file remains available for recovery.
        tx.execute("UPDATE project_template_assets SET created_at=?1 WHERE id=?2", params![created, new_id]).map_err(|e| e.to_string())?;
        tx.execute("UPDATE project_template_assets SET deleted_at=?1, updated_at=?1 WHERE id=?2", params![Utc::now().to_rfc3339(), old_id]).map_err(|e| e.to_string())?;
        tx.commit().map_err(|e| e.to_string())?;
        Ok(new_id)
    })();
    if result.is_err() {
        let path = Path::new(workspace_root).join(relative);
        if crate::workspace::is_inside_workspace(Path::new(workspace_root), &path) { let _ = fs::remove_file(path); }
    }
    result
}

#[cfg(test)]
mod replacement_tests {
    use super::*;
    use std::cell::Cell;
    const TEMPLATE: &str = "ICT项目需求导入表.docx";
    #[test]
    fn image_replacement_rejects_empty_and_spoofed_payloads() {
        for value in ["data:image/png;base64,", "data:image/png;base64,YmFk", "data:image/gif;base64,R0lGODlh"] {
            assert!(validate_replacement_image(value, 10, 10).is_err());
        }
        let png = "data:image/png;base64,iVBORw0KGgo=";
        assert!(validate_replacement_image(png, 0, 10).is_err());
        assert!(validate_replacement_image(png, 10, -1).is_err());
    }
    fn fixture() -> (Connection, PathBuf) {
        let root = std::env::temp_dir().join(format!("lamber-image-replace-{}", uuid::Uuid::new_v4()));
        fs::create_dir_all(&root).unwrap(); fs::write(root.join("old.png"), b"original").unwrap();
        let conn = Connection::open_in_memory().unwrap();
        conn.execute_batch("CREATE TABLE project_template_assets (
            id TEXT PRIMARY KEY, project_id TEXT, template_name TEXT, asset_type TEXT,
            usage TEXT, created_at TEXT, updated_at TEXT, deleted_at TEXT, relative_path TEXT);").unwrap();
        for (id, project, usage, deleted) in [("old","p","attach1",None),("sibling","p","attach1",None),
            ("second","p","attach2",None),("foreign","q","attach1",None),
            ("deleted","p","attach1",Some("gone")),("vendor","p","vendor_0",None)] {
            conn.execute("INSERT INTO project_template_assets VALUES (?1,?2,?3,'image',?4,'old-time','old-time',?5,'old.png')",
                params![id,project,TEMPLATE,usage,deleted]).unwrap();
        }
        (conn, root)
    }
    fn snapshot(conn: &Connection) -> Vec<Vec<Option<String>>> {
        conn.prepare("SELECT * FROM project_template_assets ORDER BY id").unwrap()
            .query_map([], |r| (0..9).map(|i| r.get(i)).collect()).unwrap().collect::<Result<_,_>>().unwrap()
    }
    fn save(conn: &Connection, root: &Path, usage: &str) -> Result<String, String> {
        fs::write(root.join("new.png"), b"replacement").unwrap();
        conn.execute("INSERT INTO project_template_assets VALUES ('new','p',?1,'image',?2,'new-time','new-time',NULL,'new.png')",params![TEMPLATE,usage]).map_err(|e|e.to_string())?;
        Ok("new".into())
    }
    #[test]
    fn image_replacement_rejects_stale_foreign_and_wrong_slot_before_save() {
        let (conn,root)=fixture(); let before=snapshot(&conn); let calls=Cell::new(0);
        for (project,template,id) in [("q",TEMPLATE,"old"),("p",TEMPLATE,"foreign"),("p",TEMPLATE,"deleted"),
            ("p",TEMPLATE,"missing"),("p",TEMPLATE,"vendor"),("p","其他需求导入表.docx","old"),
            ("p","会审纪要.docx","old")] {
            assert!(replace_demand_image_internal(&conn,root.to_str().unwrap(),project,template,id,|tx,usage|{
                calls.set(calls.get()+1);save(tx,&root,usage)
            }).is_err());
            assert_eq!(calls.get(),0);assert_eq!(snapshot(&conn),before);
        }
        fs::remove_dir_all(root).unwrap();
    }
    #[test]
    fn image_replacement_failure_rolls_back_and_success_preserves_other_assets() {
        let (conn,root)=fixture();let before=snapshot(&conn);
        assert!(replace_demand_image_internal(&conn,root.to_str().unwrap(),"p",TEMPLATE,"old",|_,_|Err("save failed".into())).is_err());
        assert_eq!(snapshot(&conn),before);
        // Simulate a DB failure AFTER the new image has been saved but before commit.
        conn.execute_batch("CREATE TRIGGER fail_delete BEFORE UPDATE OF deleted_at ON project_template_assets BEGIN SELECT RAISE(ABORT,'injected delete failure'); END;").unwrap();
        assert!(replace_demand_image_internal(&conn,root.to_str().unwrap(),"p",TEMPLATE,"old",|tx,usage|save(tx,&root,usage)).is_err());
        assert_eq!(snapshot(&conn),before);assert!(!root.join("new.png").exists());
        assert_eq!(fs::read(root.join("old.png")).unwrap(),b"original");
        conn.execute_batch("DROP TRIGGER fail_delete").unwrap();
        assert_eq!(replace_demand_image_internal(&conn,root.to_str().unwrap(),"p",TEMPLATE,"old",|tx,usage|save(tx,&root,usage)).unwrap(),"new");
        let after=snapshot(&conn);
        for row in before.iter().filter(|row|row[0].as_deref()!=Some("old")) {assert!(after.contains(row));}
        assert!(after.iter().find(|r|r[0].as_deref()==Some("old")).unwrap()[7].is_some());
        let new=after.iter().find(|r|r[0].as_deref()==Some("new")).unwrap();
        assert_eq!(new[5].as_deref(),Some("old-time"));assert!(new[7].is_none());
        let calls=Cell::new(0);
        assert!(replace_demand_image_internal(&conn,root.to_str().unwrap(),"p",TEMPLATE,"old",|_,_|{calls.set(1);Ok("bad".into())}).is_err());
        assert_eq!(calls.get(),0);assert_eq!(snapshot(&conn),after);
        assert_eq!(fs::read(root.join("old.png")).unwrap(),b"original");
        fs::remove_dir_all(root).unwrap();
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

// Demand slots are asset-table owned: a chat upload need not have a saved form reference yet.
fn demand_slot_asset_ids(
    conn: &Connection,
    project_id: &str,
) -> Result<std::collections::HashSet<String>, String> {
    let mut statement = conn
        .prepare(
            "SELECT id FROM project_template_assets WHERE project_id = ?1 AND deleted_at IS NULL
         AND template_name LIKE '%需求导入表%' AND usage IN ('attach1', 'attach2')",
        )
        .map_err(|error| error.to_string())?;
    let rows = statement
        .query_map([project_id], |row| row.get::<_, String>(0))
        .map_err(|error| error.to_string())?;
    rows.collect::<Result<std::collections::HashSet<_>, _>>()
        .map_err(|error| error.to_string())
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

    let mut active_assets = demand_slot_asset_ids(conn, project_id)?;
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

#[cfg(test)]
mod demand_asset_tests {
    use super::*;

    #[test]
    fn cleanup_preserves_only_live_demand_slots_in_the_requested_project() {
        let conn = Connection::open_in_memory().unwrap();
        conn.execute_batch("CREATE TABLE project_template_assets (id TEXT, project_id TEXT, template_name TEXT, usage TEXT, deleted_at TEXT);
            INSERT INTO project_template_assets VALUES
            ('chat1', 'p1', 'ICT项目需求导入表模板.docx', 'attach1', NULL),
            ('chat2', 'p1', 'ICT项目需求导入表模板.docx', 'attach2', NULL),
            ('removed', 'p1', 'ICT项目需求导入表模板.docx', 'attach1', 'deleted'),
            ('other-project', 'p2', 'ICT项目需求导入表模板.docx', 'attach1', NULL),
            ('vendor', 'p1', '会审纪要.docx', 'attach1', NULL),
            ('unknown', 'p1', 'ICT项目需求导入表模板.docx', 'other', NULL);").unwrap();
        let ids = demand_slot_asset_ids(&conn, "p1").unwrap();
        assert_eq!(
            ids,
            ["chat1".to_string(), "chat2".to_string()]
                .into_iter()
                .collect()
        );
        assert!(demand_slot_asset_ids(&conn, "missing").unwrap().is_empty());
    }
}
