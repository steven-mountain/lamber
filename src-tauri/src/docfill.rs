use crate::config_manager;
use base64::Engine;
use regex::Regex;
use std::collections::HashMap;
use std::fs::File;
use std::io::{Read, Write};
use zip::{write::SimpleFileOptions, ZipArchive, ZipWriter};

/// Attempts to parse out `{variable}` placeholders.
/// In raw XML, tags might be fragmented like `<w:t>{</w:t> ... <w:t>name</w:t>`.
/// To handle this, we do a purely text-based extraction by stripping XML tags first.
fn internal_generate_docx(
    app_handle: Option<&tauri::AppHandle>,
    db_conn: Option<&rusqlite::Connection>,
    workspace_root: Option<&str>,
    template_path: &str,
    output_path: &str,
    variables: &HashMap<String, String>,
) -> Result<(), String> {
    let file = File::open(template_path).map_err(|e| format!("Failed to open template: {}", e))?;
    let mut archive =
        ZipArchive::new(file).map_err(|e| format!("Failed to read template zip: {}", e))?;

    let out_file =
        File::create(output_path).map_err(|e| format!("Failed to create output: {}", e))?;
    let mut zip_writer = ZipWriter::new(out_file);

    let options: SimpleFileOptions = SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Stored)
        .unix_permissions(0o755);

    let mut files: Vec<(String, Vec<u8>)> = Vec::new();
    for i in 0..archive.len() {
        let mut f = archive.by_index(i).unwrap();
        let name = f.name().to_string();
        let mut content = Vec::new();
        f.read_to_end(&mut content)
            .map_err(|e| format!("Read error: {}", e))?;
        files.push((name, content));
    }

    let mut image_map: HashMap<String, Vec<(Vec<u8>, String, String, u32, u32, String)>> =
        HashMap::new();
    for (k, v) in variables {
        if !k.contains("IMAGE") && !k.contains("SCREENSHOT") {
            continue;
        }
        let val = v.trim();
        if val.is_empty() {
            continue;
        }

        let mut processed = Vec::new();
        let mut raw_images = Vec::new();
        if val.starts_with('[') {
            // JSON array of images
            if let Ok(list) = serde_json::from_str::<Vec<serde_json::Value>>(val) {
                for item in list {
                    let data = item["assetId"].as_str().or_else(|| item["data"].as_str());
                    if let Some(data) = data {
                        let w = item["width"].as_u64().unwrap_or(0) as u32;
                        let h = item["height"].as_u64().unwrap_or(0) as u32;
                        let title = item["title"].as_str().unwrap_or("").to_string();

                        if data.starts_with("asset_") {
                            if let (Some(app), Some(conn)) = (app_handle, db_conn) {
                                if let Some(root) = workspace_root {
                                    if let Ok(physical_path) = crate::project_files::assets::get_template_asset_path_internal(app, conn, root, data) {
                                    if let Ok(bytes) = std::fs::read(&physical_path) {
                                        let path = std::path::Path::new(&physical_path);
                                        let ext = path.extension().and_then(|e| e.to_str()).unwrap_or("png").to_string();
                                        let ct = if ext == "png" {
                                            "image/png"
                                        } else if ext == "jpg" || ext == "jpeg" {
                                            "image/jpeg"
                                        } else {
                                            "image/webp"
                                        }.to_string();
                                        processed.push((bytes, ext, ct, w, h, title));
                                    }
                                    }
                                }
                            }
                        } else if data.starts_with("data:image/") {
                            raw_images.push((data.to_string(), w, h, title));
                        }
                    }
                }
            }
        } else if val.starts_with("data:image/") {
            // Single image (legacy support)
            raw_images.push((val.to_string(), 0, 0, "".to_string()));
        } else if val.starts_with("asset_") {
            // Single assetId directly
            if let (Some(app), Some(conn)) = (app_handle, db_conn) {
                if let Some(root) = workspace_root {
                    if let Ok(physical_path) =
                        crate::project_files::assets::get_template_asset_path_internal(
                            app, conn, root, val,
                        )
                    {
                        if let Ok(bytes) = std::fs::read(&physical_path) {
                            let path = std::path::Path::new(&physical_path);
                            let ext = path
                                .extension()
                                .and_then(|e| e.to_str())
                                .unwrap_or("png")
                                .to_string();
                            let ct = if ext == "png" {
                                "image/png"
                            } else if ext == "jpg" || ext == "jpeg" {
                                "image/jpeg"
                            } else {
                                "image/webp"
                            }
                            .to_string();
                            processed.push((bytes, ext, ct, 0, 0, "".to_string()));
                        }
                    }
                }
            }
        }

        for (data_url, w, h, title) in raw_images {
            let (meta, b64) = match data_url.split_once(',') {
                Some(x) => x,
                None => continue,
            };
            let ext = if meta.contains("image/png") {
                "png"
            } else {
                "jpg"
            };
            let ct = if ext == "png" {
                "image/png"
            } else {
                "image/jpeg"
            };
            if let Ok(bytes) = base64::engine::general_purpose::STANDARD.decode(b64) {
                processed.push((bytes, ext.to_string(), ct.to_string(), w, h, title));
            }
        }
        if !processed.is_empty() {
            image_map.insert(k.to_string(), processed);
        }
    }

    let mut rels_additions: Vec<(String, String)> = Vec::new();
    let mut media_additions: Vec<(String, Vec<u8>)> = Vec::new();
    let mut content_type_additions: Vec<(String, String)> = Vec::new();

    let mut docpr_id: u32 = 3000;

    for (name, content) in files.iter_mut() {
        if name == "word/document.xml" {
            let mut xml_str = String::from_utf8(content.clone()).map_err(|e| e.to_string())?;
            xml_str = clean_xml_placeholders(&xml_str);
            xml_str = normalize_signoff_project_situation_placeholders(&xml_str);

            for (k, v) in variables {
                if k.starts_with("TABLE_") {
                    if let Ok(rows_data) =
                        serde_json::from_str::<Vec<std::collections::HashMap<String, String>>>(v)
                    {
                        let first_key = if !rows_data.is_empty() {
                            rows_data[0].keys().next().cloned()
                        } else {
                            if k == "TABLE_TECH_ITEMS" {
                                Some("TECH_ITEM_NAME".to_string())
                            } else if k == "TABLE_INQ_VENDORS" {
                                Some("INQ_VENDOR_NAME".to_string())
                            } else {
                                None
                            }
                        };

                        if let Some(first_key) = first_key {
                            let pattern = format!("{{{}}}", first_key);

                            if let Some(idx) = xml_str.find(&pattern) {
                                let tr_start = xml_str[..idx]
                                    .rfind("<w:tr>")
                                    .or_else(|| xml_str[..idx].rfind("<w:tr "))
                                    .unwrap_or(0);
                                let tr_end_rel = xml_str[idx..]
                                    .find("</w:tr>")
                                    .unwrap_or(xml_str.len() - idx);
                                let tr_end = idx + tr_end_rel + 7;

                                if tr_start < tr_end && tr_end <= xml_str.len() {
                                    let row_xml = &xml_str[tr_start..tr_end];
                                    let mut new_rows = String::new();

                                    for row_data in rows_data {
                                        let mut new_row = row_xml.to_string();
                                        for (rk, rv) in &row_data {
                                            let r_pattern = format!("{{{}}}", rk);
                                            let escaped_rv = rv
                                                .replace("&", "&amp;")
                                                .replace("<", "&lt;")
                                                .replace(">", "&gt;");
                                            let docx_rv =
                                                escaped_rv.replace("\n", "</w:t><w:br/><w:t>");
                                            new_row = new_row.replace(&r_pattern, &docx_rv);
                                        }
                                        let re = regex::Regex::new(r"\{[A-Z_0-9]+\}").unwrap();
                                        new_row = re.replace_all(&new_row, "").to_string();
                                        new_rows.push_str(&new_row);
                                    }
                                    xml_str = format!(
                                        "{}{}{}",
                                        &xml_str[..tr_start],
                                        new_rows,
                                        &xml_str[tr_end..]
                                    );
                                }
                            }
                        }
                    }
                }
            }

            for (key, images) in &image_map {
                let placeholder = format!("{{{}}}", key);
                if !xml_str.contains(&placeholder) {
                    continue;
                }

                let mut combined_xml = String::new();
                // Close the current text run and paragraph in which the placeholder resides
                combined_xml.push_str("</w:t></w:r></w:p>");

                let mut prev_title = String::new();
                for (idx, (bytes, ext, ct, w, h, title)) in images.iter().enumerate() {
                    let safe_key = format!(
                        "{}_{}",
                        key.to_lowercase()
                            .replace(|c: char| !c.is_ascii_alphanumeric() && c != '_', "_"),
                        idx
                    );
                    let media_name = format!("word/media/{}.{}", safe_key, ext);
                    let rid = format!("rId{}", docpr_id);
                    docpr_id += 1;

                    rels_additions.push((rid.clone(), format!("media/{}.{}", safe_key, ext)));
                    media_additions.push((media_name, bytes.clone()));
                    content_type_additions.push((ext.clone(), ct.clone()));

                    // Default to ~6 inches width (5,500,000 EMUs)
                    let mut cx = 5_500_000;
                    let mut cy = 3_000_000;

                    if *w > 0 && *h > 0 {
                        let original_cx = (*w as u64) * 9525; // 1 pixel ~= 9525 EMUs at 96dpi
                        let original_cy = (*h as u64) * 9525;

                        if original_cx < cx {
                            cx = original_cx;
                            cy = original_cy;
                        } else {
                            // Scale down to max width
                            cy = (original_cy * cx) / original_cx;
                        }
                    }

                    let is_new_vendor = title != &prev_title && !title.is_empty();

                    if is_new_vendor {
                        if idx > 0 {
                            // Add empty line between DIFFERENT vendors
                            combined_xml
                                .push_str("<w:p><w:pPr><w:jc w:val=\"left\"/></w:pPr></w:p>");
                        }

                        // Open left-aligned paragraph for the vendor
                        combined_xml.push_str("<w:p><w:pPr><w:jc w:val=\"left\"/></w:pPr>");

                        let escaped_title = title
                            .replace("&", "&amp;")
                            .replace("<", "&lt;")
                            .replace(">", "&gt;");
                        // Render Title as a bolded text run, then a break
                        combined_xml.push_str(&format!(
                            r#"<w:r><w:rPr><w:b/></w:rPr><w:t>{}</w:t></w:r><w:r><w:br/></w:r>"#,
                            escaped_title
                        ));
                        prev_title = title.clone();
                    } else {
                        // Same vendor (or empty title), no empty line, no title, just open a new paragraph
                        combined_xml.push_str("<w:p><w:pPr><w:jc w:val=\"left\"/></w:pPr>");
                    }

                    // Render the drawing run inside the same paragraph
                    combined_xml.push_str(&format!(
                        r#"<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="{docid}" name="Picture {docid}"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="0" name="{safe_key}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#,
                        cx = cx,
                        cy = cy,
                        docid = docpr_id,
                        safe_key = safe_key,
                        rid = rid
                    ));

                    // Close the paragraph
                    combined_xml.push_str("</w:p>");
                }

                // Re-open a paragraph and run for subsequent content
                combined_xml.push_str("<w:p><w:r><w:t>");

                // Safely replace the placeholder within its run, preserving the paragraph and surrounding text
                xml_str = xml_str.replace(&placeholder, &combined_xml);
            }

            for (k, v) in variables {
                if k.starts_with("TABLE_") {
                    continue;
                }
                if image_map.contains_key(k) {
                    continue;
                }
                if (k.contains("IMAGE") || k.contains("SCREENSHOT"))
                    && v.trim_start().starts_with('[')
                {
                    continue;
                }
                let pattern = format!("{{{}}}", k);
                let escaped_v = v
                    .replace("&", "&amp;")
                    .replace("<", "&lt;")
                    .replace(">", "&gt;");

                let docx_v = if k == "VENDOR_SCREENSHOT_LIST" {
                    let mut formatted = String::new();
                    // Close original paragraph
                    formatted.push_str("</w:t></w:r></w:p>");
                    for line in escaped_v.lines() {
                        if line.is_empty() {
                            formatted.push_str("<w:p><w:pPr><w:jc w:val=\"left\"/></w:pPr></w:p>");
                        } else {
                            formatted.push_str(&format!(
                                r#"<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:t>{}</w:t></w:r></w:p>"#,
                                line
                            ));
                        }
                    }
                    // Re-open paragraph
                    formatted.push_str("<w:p><w:r><w:t>");
                    formatted
                } else {
                    escaped_v.replace("\n", "</w:t><w:br/><w:t>")
                };

                xml_str = xml_str.replace(&pattern, &docx_v);
            }

            let unresolved_re = Regex::new(r"\{[a-zA-Z0-9_\u4e00-\u9fa5（）]+\}").unwrap();
            xml_str = unresolved_re.replace_all(&xml_str, "").to_string();

            *content = xml_str.into_bytes();
        }
    }

    if !rels_additions.is_empty() {
        for (name, content) in files.iter_mut() {
            if name == "word/_rels/document.xml.rels" {
                let mut xml = String::from_utf8(content.clone()).map_err(|e| e.to_string())?;
                if let Some(idx) = xml.rfind("</Relationships>") {
                    let mut insert = String::new();
                    for (rid, target) in &rels_additions {
                        insert.push_str(&format!(r#"<Relationship Id="{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{}"/>"#, rid, target));
                    }
                    xml = format!("{}{}{}", &xml[..idx], insert, &xml[idx..]);
                }
                *content = xml.into_bytes();
            }
            if name == "[Content_Types].xml" {
                let mut xml = String::from_utf8(content.clone()).map_err(|e| e.to_string())?;
                for (ext, ct) in &content_type_additions {
                    let pat = format!(r#"Extension="{}""#, ext);
                    if xml.contains(&pat) {
                        continue;
                    }
                    if let Some(idx) = xml.rfind("</Types>") {
                        let insert =
                            format!(r#"<Default Extension="{}" ContentType="{}"/>"#, ext, ct);
                        xml = format!("{}{}{}", &xml[..idx], insert, &xml[idx..]);
                    }
                }
                *content = xml.into_bytes();
            }
        }
    }

    for (name, content) in &files {
        zip_writer
            .start_file(name, options)
            .map_err(|e| e.to_string())?;
        zip_writer.write_all(content).map_err(|e| e.to_string())?;
    }
    for (name, content) in &media_additions {
        zip_writer
            .start_file(name, options)
            .map_err(|e| e.to_string())?;
        zip_writer.write_all(content).map_err(|e| e.to_string())?;
    }

    zip_writer
        .finish()
        .map_err(|e| format!("Finish zip error: {}", e))?;
    Ok(())
}

/// Robustly joins fragmented placeholders in Word XML.
/// e.g. {<w:t>VAR</w:t>} -> {VAR}
fn clean_xml_placeholders(xml: &str) -> String {
    let re = Regex::new(r"\{(<[^>]+>|[^}])*?\}").unwrap();
    let tag_re = Regex::new(r"<[^>]+>").unwrap();

    re.replace_all(xml, |caps: &regex::Captures| {
        let matched = &caps[0];
        if matched.contains('<') {
            // Fragmented placeholder detected. Strip internal tags.
            let stripped = tag_re.replace_all(matched, "");
            stripped.to_string()
        } else {
            matched.to_string()
        }
    })
    .to_string()
}

fn normalize_signoff_project_situation_placeholders(xml: &str) -> String {
    xml.replace(
        "项目投入（不含税）：总投入{PROJECT_TOTAL_INVESTMENT}元：其中{SUBJECT_IT_COST}投入{IT_INVESTMENT}元，{SUBJECT_CT_COST}投入{CT_INVESTMENT}元，此IT部分费用参考三家询价",
        "项目投入（不含税）：{PROJECT_INVESTMENT_SITUATION}",
    )
    .replace("最低价，申请立项后甄选，最终费用不超过上述总投入。", "")
    .replace(
        "项目收入（不含税）：总收入{PROJECT_TOTAL_REVENUE}元；其中{SUBJECT_IT_REV}收入{IT_REVENUE}元，{SUBJECT_CT_REV}收入{CT_REVENUE}元。",
        "项目收入（不含税）：{PROJECT_REVENUE_SITUATION}",
    )
}

#[tauri::command]
pub fn get_available_templates(
    state: tauri::State<'_, std::sync::Mutex<config_manager::AppConfig>>,
    module_id: String,
) -> Result<Vec<String>, String> {
    use std::fs;
    let config = state.lock().unwrap();
    let module_path = config
        .module_paths
        .get(&module_id)
        .ok_or("未设置工作目录")?;
    let template_dir = std::path::Path::new(module_path).join("templates");

    if !template_dir.exists() {
        return Ok(vec![]);
    }

    let mut templates = Vec::new();
    if let Ok(entries) = fs::read_dir(&template_dir) {
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_file() {
                if let Some(ext) = path.extension() {
                    let ext_str = ext.to_string_lossy().to_lowercase();
                    let file_name = path.file_name().unwrap().to_string_lossy().to_string();
                    if (ext_str == "docx" || ext_str == "xlsx")
                        && !file_name.starts_with("~$")
                        && !file_name.starts_with(".~")
                    {
                        templates.push(file_name);
                    }
                }
            }
        }
    }

    Ok(templates)
}

fn resolve_lifecycle_output_dir(
    conn: &rusqlite::Connection,
    workspace_root: &std::path::Path,
    project_id: Option<&str>,
    requested_output_dir: Option<&str>,
    base_path: &std::path::Path,
) -> Result<std::path::PathBuf, String> {
    if let Some(path) = requested_output_dir
        .map(str::trim)
        .filter(|path| !path.is_empty())
    {
        return Ok(crate::workspace::resolve_workspace_path(
            workspace_root,
            path,
        ));
    }

    if let Some(project_id) = project_id.map(str::trim).filter(|id| !id.is_empty()) {
        let mut stmt = conn
            .prepare(
                "SELECT relative_path, linked_folder_relative_path, folder_path, folder_name, linked_folder_external_path FROM projects WHERE id = ?1",
            )
            .map_err(|e| e.to_string())?;
        let mut rows = stmt.query([project_id]).map_err(|e| e.to_string())?;
        let row = rows
            .next()
            .map_err(|e| e.to_string())?
            .ok_or_else(|| format!("未找到项目: {}", project_id))?;

        let candidates: [Option<String>; 5] = [
            row.get(0).map_err(|e| e.to_string())?,
            row.get(1).map_err(|e| e.to_string())?,
            row.get(2).map_err(|e| e.to_string())?,
            row.get(3).map_err(|e| e.to_string())?,
            row.get(4).map_err(|e| e.to_string())?,
        ];

        if let Some(path) = candidates
            .iter()
            .flatten()
            .map(|path| path.trim())
            .find(|path| !path.is_empty())
        {
            return Ok(crate::workspace::resolve_workspace_path(
                workspace_root,
                path,
            ));
        }
    }

    Ok(base_path.join("output"))
}

#[tauri::command]
pub fn generate_lifecycle_docs(
    app: tauri::AppHandle,
    runtime: tauri::State<'_, std::sync::Arc<crate::workspace::WorkspaceRuntime>>,
    state: tauri::State<'_, std::sync::Mutex<config_manager::AppConfig>>,
    module_id: String,
    variables: HashMap<String, String>,
    selected_templates: Vec<String>,
    output_dir: Option<String>,
    project_id: Option<String>,
    overwrite_existing: Option<bool>,
) -> Result<String, String> {
    let workspace = runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    use std::fs;

    let module_path = {
        let config = state.lock().unwrap();
        config
            .module_paths
            .get(&module_id)
            .ok_or("未设置工作目录")?
            .clone()
    };
    let base_path = std::path::Path::new(&module_path);

    let template_dir = base_path.join("templates");

    if !template_dir.exists() {
        return Err(format!("未找到模板目录: {}", template_dir.display()));
    }

    let workspace_root = std::path::Path::new(&workspace.workspace_root);
    let output_dir = resolve_lifecycle_output_dir(
        &conn,
        workspace_root,
        project_id.as_deref(),
        output_dir.as_deref(),
        base_path,
    )?;
    if !output_dir.exists() {
        fs::create_dir_all(&output_dir).map_err(|e| format!("创建输出目录失败: {}", e))?;
    }
    let overwrite_existing = overwrite_existing.unwrap_or(false);

    let mut generated_count = 0;

    // Iterate over files in template directory
    let entries = fs::read_dir(&template_dir).map_err(|e| e.to_string())?;
    for entry_result in entries {
        let entry = entry_result.map_err(|e| e.to_string())?;
        let path = entry.path();

        if path.is_file() {
            if let Some(ext) = path.extension() {
                let ext_str = ext.to_string_lossy().to_lowercase();
                let file_name = path.file_name().unwrap().to_string_lossy().to_string();

                // Ignore temporary files created by MS Word/Excel (starting with ~$)
                if (ext_str == "docx" || ext_str == "xlsx")
                    && !file_name.starts_with("~$")
                    && !file_name.starts_with(".~")
                {
                    // Only generate files that the user explicitly selected
                    if !selected_templates.contains(&file_name) {
                        continue;
                    }

                    let proj_name = variables
                        .get("PROJECT_NAME")
                        .cloned()
                        .unwrap_or_else(|| "未命名".to_string());
                    let safe_proj_name = proj_name
                        .chars()
                        .filter(|c| !r#"\/:*?"<>|"#.contains(*c))
                        .collect::<String>();

                    // First clean up some generic template markings
                    let mut clean_name = file_name
                        .replace("模板", "")
                        .replace("【2024版】", "")
                        .replace("【2025版】", "")
                        .replace("_变量版", "");
                    // Remove extension
                    if let Some(dot_idx) = clean_name.rfind('.') {
                        clean_name = clean_name[..dot_idx].to_string();
                    }
                    // Trim trailing hyphens or underscores
                    clean_name = clean_name
                        .trim_end_matches('-')
                        .trim_end_matches('_')
                        .to_string();

                    // Reconstruct: clean_name-project_name.extension
                    let out_name = format!("{}-{}.{}", clean_name, safe_proj_name, ext_str);

                    let out_path = output_dir.join(&out_name);
                    if out_path.exists() && !overwrite_existing {
                        return Err(format!("FILE_EXISTS::{}", out_path.to_string_lossy()));
                    }

                    if ext_str == "docx" {
                        if let Err(e) = internal_generate_docx(
                            Some(&app),
                            Some(&conn),
                            Some(&workspace.workspace_root),
                            path.to_str().unwrap(),
                            out_path.to_str().unwrap(),
                            &variables,
                        ) {
                            println!(
                                "Warning: failed to process docx template {}: {}",
                                file_name, e
                            );
                            continue;
                        }
                    } else if ext_str == "xlsx" {
                        // Create a copy of the excel file to output path
                        if let Err(e) = fs::copy(&path, &out_path) {
                            println!("Warning: failed to copy xlsx template {}: {}", file_name, e);
                            continue;
                        }
                        if let Err(e) =
                            internal_generate_xlsx(out_path.to_str().unwrap(), &variables)
                        {
                            println!(
                                "Warning: failed to process xlsx template {}: {}",
                                file_name, e
                            );
                            continue;
                        }
                    }

                    generated_count += 1;
                }
            }
        }
    }

    if generated_count == 0 {
        return Err("模板目录中未找到任何可生成的 .docx 或 .xlsx 模板文件。".into());
    }

    Ok(output_dir.to_string_lossy().to_string())
}

fn internal_generate_xlsx(
    output_path: &str,
    variables: &HashMap<String, String>,
) -> Result<(), String> {
    use umya_spreadsheet::*;
    let mut book = reader::xlsx::read(std::path::Path::new(output_path))
        .map_err(|e| format!("无法读取 Excel: {}", e))?;

    if let Some(sheet) = book.get_sheet_by_name_mut("3-直接经济效益评估表") {
        if let Some(v) = variables.get("PROJECT_NAME") {
            let c = sheet.get_cell_mut("D2");
            c.set_value(v);
            c.set_formula("");
        }

        let mut set_cell = |cell: &str, key: &str, as_text: bool| {
            if let Some(v) = variables.get(key) {
                let cell_obj = sheet.get_cell_mut(cell);
                cell_obj.set_formula("");
                if as_text {
                    cell_obj.set_value(v);
                    return;
                }

                let mut num_str = v.replace(",", "");
                let mut is_pct = false;
                if num_str.ends_with('%') {
                    num_str = num_str.trim_end_matches('%').to_string();
                    is_pct = true;
                }

                if let Ok(mut num) = num_str.parse::<f64>() {
                    if is_pct {
                        num /= 100.0;
                    }
                    cell_obj.set_value_number(num);
                } else {
                    cell_obj.set_value(v);
                }
            }
        };

        let subject_mappings = [
            ("D3", "G3", "Q3", "EXCEL_REV_IT_INTEGRATION"),
            ("D4", "G4", "Q4", "EXCEL_REV_IT_MAINTENANCE"),
            ("D5", "G5", "Q5", "EXCEL_REV_IT_DEVICE_SALES"),
            ("D6", "G6", "Q6", "EXCEL_REV_IT_DEVICE_LEASE"),
            ("D7", "G7", "Q7", "EXCEL_REV_IT_OTHER"),
            ("D8", "G8", "Q8", "EXCEL_REV_IT_CLOUD"),
            ("D9", "G9", "Q9", "EXCEL_REV_CT_LINE"),
            ("D10", "G10", "Q10", "EXCEL_REV_CT_PRODUCT"),
            ("D11", "G11", "Q11", "EXCEL_REV_NON_IT_CT"),
            ("E13", "G13", "Q13", "EXCEL_COST_IT_DEVICE"),
            ("E14", "G14", "Q14", "EXCEL_COST_IT_CONSTRUCTION"),
            ("E15", "G15", "Q15", "EXCEL_COST_IT_SURVEY"),
            ("E16", "G16", "Q16", "EXCEL_COST_IT_INTEGRATION"),
            ("E17", "G17", "Q17", "EXCEL_COST_IT_OTHER"),
            ("E18", "G18", "Q18", "EXCEL_COST_IT_MAINTENANCE"),
            ("E19", "G19", "Q19", "EXCEL_COST_IT_RUNNING"),
            ("E20", "G20", "Q20", "EXCEL_COST_IT_BIDDING"),
            ("E21", "G21", "Q21", "EXCEL_COST_IT_DESIGN_EVAL"),
            ("E22", "G22", "Q22", "EXCEL_COST_IT_AUDIT"),
            ("E23", "G23", "Q23", "EXCEL_COST_CT_CONSTRUCTION"),
            ("E24", "G24", "Q24", "EXCEL_COST_CT_MAINTENANCE"),
            ("E25", "G25", "Q25", "EXCEL_COST_CT_OTHER"),
            ("E26", "G26", "Q26", "EXCEL_COST_CT_BANDWIDTH"),
            ("E27", "G27", "Q27", "EXCEL_COST_CT_RENEWAL"),
            ("E28", "G28", "Q28", "EXCEL_COST_NON_IT_CT"),
            ("D29", "G29", "Q29", "EXCEL_COST_MIX_MARKETING"),
            ("D30", "G30", "Q30", "EXCEL_COST_MIX_CHANNEL"),
            ("D31", "G31", "Q31", "EXCEL_COST_MIX_OTHER"),
        ];

        for (name_cell, excl_cell, incl_cell, prefix) in subject_mappings {
            set_cell(name_cell, &format!("{}_NAME", prefix), true);
            set_cell(excl_cell, &format!("{}_EXCL", prefix), false);
            set_cell(incl_cell, &format!("{}_INCL", prefix), false);
        }

        for year in 1..=10 {
            let in_cell = format!("E{}", 33 + year);
            let out_cell = format!("G{}", 33 + year);
            set_cell(&in_cell, &format!("CASH_IN_Y{}", year), false);
            set_cell(&out_cell, &format!("CASH_OUT_Y{}", year), false);
        }
    }

    if let Some(sheet2) = book.get_sheet_by_name_mut("2-ICT项目评估结果") {
        if let Some(v) = variables.get("PROJECT_NAME") {
            let c = sheet2.get_cell_mut("B4");
            c.set_value(v);
            c.set_formula("");
        }
        if let Some(v) = variables.get("CUSTOMER_NAME") {
            let c = sheet2.get_cell_mut("B5");
            c.set_value(v);
            c.set_formula("");
        }
        if let Some(v) = variables.get("RENEWAL_PROJECT_FLAG") {
            let c = sheet2.get_cell_mut("B6");
            c.set_value(v);
            c.set_formula("");
        }
        if let Some(v) = variables.get("IT_BUSINESS_MODE") {
            let c = sheet2.get_cell_mut("B7");
            c.set_value(v);
            c.set_formula("");
        }
        if let Some(v) = variables.get("CONTRACT_DURATION") {
            let c = sheet2.get_cell_mut("B8");
            c.set_value(v);
            c.set_formula("");
        }
        if let Some(v) = variables.get("IT_FUNDING_SOURCE") {
            let c = sheet2.get_cell_mut("B9");
            c.set_value(v);
            c.set_formula("");
        }
    }

    writer::xlsx::write(&book, std::path::Path::new(output_path))
        .map_err(|e| format!("保存 Excel 失败: {}", e))?;

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::internal_generate_xlsx;
    use calamine::{open_workbook, Reader, Xlsx};
    use std::collections::HashMap;
    use std::fs;
    use std::path::Path;
    use std::time::{SystemTime, UNIX_EPOCH};

    fn cell_string(range: &calamine::Range<calamine::Data>, row: u32, col: u32) -> String {
        range
            .get_value((row - 1, col - 1))
            .map(|cell| cell.to_string())
            .unwrap_or_default()
    }

    fn cell_number(range: &calamine::Range<calamine::Data>, row: u32, col: u32) -> f64 {
        range
            .get_value((row - 1, col - 1))
            .and_then(|cell| match cell {
                calamine::Data::Float(value) => Some(*value),
                calamine::Data::Int(value) => Some(*value as f64),
                calamine::Data::String(value) => value.parse::<f64>().ok(),
                _ => None,
            })
            .unwrap_or(0.0)
    }

    #[test]
    fn lifecycle_xlsx_subject_mapping_writes_names_excl_and_incl() {
        let template_path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../项目全生命周期文件模版/效益分析表 .xlsx");
        assert!(
            template_path.exists(),
            "missing test template: {}",
            template_path.display()
        );

        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let output_path =
            std::env::temp_dir().join(format!("lamber-xlsx-subjects-{}.xlsx", suffix));
        fs::copy(&template_path, &output_path).unwrap();

        let mut variables = HashMap::new();
        variables.insert(
            "EXCEL_REV_CT_PRODUCT_NAME".to_string(),
            "产品收入（视频监控）".to_string(),
        );
        variables.insert("EXCEL_REV_CT_PRODUCT_EXCL".to_string(), "41.51".to_string());
        variables.insert("EXCEL_REV_CT_PRODUCT_INCL".to_string(), "44".to_string());
        variables.insert(
            "EXCEL_COST_CT_OTHER_NAME".to_string(),
            "其他产品成本（视频监控）".to_string(),
        );
        variables.insert("EXCEL_COST_CT_OTHER_EXCL".to_string(), "41.51".to_string());
        variables.insert("EXCEL_COST_CT_OTHER_INCL".to_string(), "44".to_string());

        internal_generate_xlsx(output_path.to_str().unwrap(), &variables).unwrap();

        let mut workbook: Xlsx<_> = open_workbook(&output_path).unwrap();
        let range = workbook.worksheet_range("3-直接经济效益评估表").unwrap();

        assert_eq!(cell_string(&range, 10, 4), "产品收入（视频监控）");
        assert!((cell_number(&range, 10, 7) - 41.51).abs() < 0.001);
        assert!((cell_number(&range, 10, 17) - 44.0).abs() < 0.001);

        assert_eq!(cell_string(&range, 25, 5), "其他产品成本（视频监控）");
        assert!((cell_number(&range, 25, 7) - 41.51).abs() < 0.001);
        assert!((cell_number(&range, 25, 17) - 44.0).abs() < 0.001);

        let _ = fs::remove_file(output_path);
    }

    #[test]
    fn lifecycle_xlsx_blank_amounts_clear_formula_and_do_not_write_zero() {
        let template_path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../项目全生命周期文件模版/效益分析表 .xlsx");
        assert!(
            template_path.exists(),
            "missing test template: {}",
            template_path.display()
        );

        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let output_path =
            std::env::temp_dir().join(format!("lamber-xlsx-blank-amounts-{}.xlsx", suffix));
        fs::copy(&template_path, &output_path).unwrap();

        let mut variables = HashMap::new();
        variables.insert(
            "EXCEL_REV_CT_PRODUCT_NAME".to_string(),
            "产品收入".to_string(),
        );
        variables.insert("EXCEL_REV_CT_PRODUCT_EXCL".to_string(), "".to_string());
        variables.insert("EXCEL_REV_CT_PRODUCT_INCL".to_string(), "".to_string());

        internal_generate_xlsx(output_path.to_str().unwrap(), &variables).unwrap();

        let mut workbook: Xlsx<_> = open_workbook(&output_path).unwrap();
        let range = workbook.worksheet_range("3-直接经济效益评估表").unwrap();

        assert_eq!(cell_string(&range, 10, 4), "产品收入");
        assert_eq!(cell_string(&range, 10, 7), "");
        assert_eq!(cell_string(&range, 10, 17), "");

        let _ = fs::remove_file(output_path);
    }
}
