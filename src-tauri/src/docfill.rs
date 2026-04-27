use std::collections::{HashSet, HashMap};
use std::fs::File;
use std::io::{Read, Write};
use zip::{ZipArchive, ZipWriter, write::SimpleFileOptions};
use regex::Regex;

/// Attempts to parse out `{variable}` placeholders. 
/// In raw XML, tags might be fragmented like `<w:t>{</w:t> ... <w:t>name</w:t>`. 
/// To handle this, we do a purely text-based extraction by stripping XML tags first.
#[tauri::command]
pub fn extract_docx_variables(path: String) -> Result<Vec<String>, String> {
    let file = File::open(&path).map_err(|e| format!("Failed to open file: {}", e))?;
    let mut archive = ZipArchive::new(file).map_err(|e| format!("Failed to read zip: {}", e))?;
    
    let mut doc_xml = String::new();
    
    for i in 0..archive.len() {
        let mut file = archive.by_index(i).unwrap();
        if file.name().ends_with(".xml") {
            let mut content = String::new();
            if file.read_to_string(&mut content).is_ok() {
                doc_xml.push_str(&content);
            }
        }
    }

    if doc_xml.is_empty() {
        return Err("Could not find any xml files in the provided archive.".into());
    }

    // Strip all XML tags to find pure text content
    let tag_re = Regex::new(r"<[^>]+>").unwrap();
    let pure_text = tag_re.replace_all(&doc_xml, "");

    // Find all {var_name}
    let var_re = Regex::new(r"\{([a-zA-Z0-9_\u4e00-\u9fa5]+)\}").unwrap();
    let mut vars = HashSet::new();

    for cap in var_re.captures_iter(&pure_text) {
        if let Some(matched) = cap.get(1) {
            vars.insert(matched.as_str().to_string());
        }
    }

    let mut result: Vec<String> = vars.into_iter().collect();
    result.sort();
    Ok(result)
}

/// Generates the docx by replacing variables in the xml files.
#[tauri::command]
pub fn generate_docx(template_path: String, output_path: String, variables: HashMap<String, String>) -> Result<(), String> {
    internal_generate_docx(&template_path, &output_path, &variables)
}

fn internal_generate_docx(template_path: &str, output_path: &str, variables: &HashMap<String, String>) -> Result<(), String> {
    let file = File::open(template_path).map_err(|e| format!("Failed to open template: {}", e))?;
    let mut archive = ZipArchive::new(file).map_err(|e| format!("Failed to read template zip: {}", e))?;
    
    let out_file = File::create(output_path).map_err(|e| format!("Failed to create output: {}", e))?;
    let mut zip_writer = ZipWriter::new(out_file);
    
    let options = SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Stored)
        .unix_permissions(0o755);

    for i in 0..archive.len() {
        let mut file = archive.by_index(i).unwrap();
        let name = file.name().to_string();
        
        let mut content = Vec::new();
        file.read_to_end(&mut content).map_err(|e| format!("Read error: {}", e))?;
        
        if name.starts_with("word/") && name.ends_with(".xml") {
            if let Ok(mut xml_str) = String::from_utf8(content.clone()) {
                // Pre-process XML to join fragmented placeholders like {<w:t>V</w:t><w:t>AR</w:t>}
                xml_str = clean_xml_placeholders(&xml_str);

                // Handle dynamic tables passed as JSON array string
                for (k, v) in variables {
                    if k.starts_with("TABLE_") {
                        if let Ok(rows_data) = serde_json::from_str::<Vec<std::collections::HashMap<String, String>>>(v) {
                            // Find the first key to locate the row. We have to guess the key if it's empty.
                            // But usually, we map TABLE_TECH_ITEMS -> TECH_ITEM_NAME. 
                            // Let's deduce it from the key name if possible, or if the array is empty, we might not be able to find it easily unless we hardcode.
                            // To be safe, we always send at least one item, or we look for specific known keys.
                            let first_key = if !rows_data.is_empty() {
                                rows_data[0].keys().next().cloned()
                            } else {
                                if k == "TABLE_TECH_ITEMS" { Some("TECH_ITEM_NAME".to_string()) }
                                else if k == "TABLE_INQ_VENDORS" { Some("INQ_VENDOR_NAME".to_string()) }
                                else { None }
                            };
                            
                            if let Some(first_key) = first_key {
                                let pattern = format!("{{{}}}", first_key);
                                
                                if let Some(idx) = xml_str.find(&pattern) {
                                    let tr_start = xml_str[..idx].rfind("<w:tr>").or_else(|| xml_str[..idx].rfind("<w:tr ")).unwrap_or(0);
                                    let tr_end_rel = xml_str[idx..].find("</w:tr>").unwrap_or(xml_str.len() - idx);
                                    let tr_end = idx + tr_end_rel + 7;
                                    
                                    if tr_start < tr_end && tr_end <= xml_str.len() {
                                        let row_xml = &xml_str[tr_start..tr_end];
                                        let mut new_rows = String::new();
                                        
                                        for row_data in rows_data {
                                            let mut new_row = row_xml.to_string();
                                            for (rk, rv) in &row_data {
                                                let r_pattern = format!("{{{}}}", rk);
                                                let escaped_rv = rv.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;");
                                                let docx_rv = escaped_rv.replace("\n", "</w:t><w:br/><w:t>");
                                                new_row = new_row.replace(&r_pattern, &docx_rv);
                                            }
                                            // Also clean up any remaining unresolved {} in this row (optional, but good practice)
                                            let re = regex::Regex::new(r"\{[A-Z_0-9]+\}").unwrap();
                                            new_row = re.replace_all(&new_row, "").to_string();
                                            
                                            new_rows.push_str(&new_row);
                                        }
                                        
                                        xml_str = format!("{}{}{}", &xml_str[..tr_start], new_rows, &xml_str[tr_end..]);
                                    }
                                }
                            }
                        }
                    }
                }

                for (k, v) in variables {
                    let pattern = format!("{{{}}}", k);
                    
                    // We must escape XML characters in `v` first!
                    let escaped_v = v.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;");
                    
                    // Replace newlines with docx line breaks
                    // Since {VAR} is inside a <w:t> tag, replacing it with text</w:t><w:br/><w:t>text will create valid line breaks
                    let docx_v = escaped_v.replace("\n", "</w:t><w:br/><w:t>");
                    xml_str = xml_str.replace(&pattern, &docx_v);
                }

                let unresolved_re = Regex::new(r"\{[a-zA-Z0-9_\u4e00-\u9fa5]+\}").unwrap();
                xml_str = unresolved_re.replace_all(&xml_str, "").to_string();

                zip_writer.start_file(name, options).map_err(|e| e.to_string())?;
                zip_writer.write_all(xml_str.as_bytes()).map_err(|e| e.to_string())?;
                continue;
            }
        }
        
        zip_writer.start_file(name, options).map_err(|e| e.to_string())?;
        zip_writer.write_all(&content).map_err(|e| e.to_string())?;
    }
    
    zip_writer.finish().map_err(|e| format!("Finish zip error: {}", e))?;
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
    }).to_string()
}

#[tauri::command]
pub fn get_available_templates() -> Result<Vec<String>, String> {
    use std::fs;
    let current_dir = std::env::current_dir().map_err(|e| format!("无法获取当前目录: {}", e))?;
    let project_root = if current_dir.ends_with("src-tauri") {
        current_dir.parent().unwrap().to_path_buf()
    } else {
        current_dir.clone()
    };
    let template_dir = project_root.join("项目全生命周期文件模版");
    
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
                    if (ext_str == "docx" || ext_str == "xlsx") && !file_name.starts_with("~$") && !file_name.starts_with(".~") {
                        templates.push(file_name);
                    }
                }
            }
        }
    }
    
    Ok(templates)
}

#[tauri::command]
pub fn generate_lifecycle_docs(variables: HashMap<String, String>, selected_templates: Vec<String>) -> Result<String, String> {
    use std::fs;
    
    let current_dir = std::env::current_dir().map_err(|e| format!("无法获取当前目录: {}", e))?;
    
    let project_root = if current_dir.ends_with("src-tauri") {
        current_dir.parent().unwrap().to_path_buf()
    } else {
        current_dir.clone()
    };
    
    let template_dir = project_root.join("项目全生命周期文件模版");
    
    if !template_dir.exists() {
        return Err(format!("未找到模板目录: {}", template_dir.display()));
    }
    
    let output_dir = project_root.join("一键生成全生命周期结果");
    if !output_dir.exists() {
        fs::create_dir_all(&output_dir).map_err(|e| format!("创建输出目录失败: {}", e))?;
    }
    
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
                if (ext_str == "docx" || ext_str == "xlsx") && !file_name.starts_with("~$") && !file_name.starts_with(".~") {
                    
                    // Only generate files that the user explicitly selected
                    if !selected_templates.contains(&file_name) {
                        continue;
                    }
                    
                    let proj_name = variables.get("PROJECT_NAME").cloned().unwrap_or_else(|| "未命名".to_string());
                    let safe_proj_name = proj_name.chars().filter(|c| !r#"\/:*?"<>|"#.contains(*c)).collect::<String>();
                    
                    // First clean up some generic template markings
                    let mut clean_name = file_name.replace("模板", "").replace("【2024版】", "").replace("【2025版】", "").replace("_变量版", "");
                    // Remove extension
                    if let Some(dot_idx) = clean_name.rfind('.') {
                        clean_name = clean_name[..dot_idx].to_string();
                    }
                    // Trim trailing hyphens or underscores
                    clean_name = clean_name.trim_end_matches('-').trim_end_matches('_').to_string();
                    
                    // Reconstruct: clean_name-project_name.extension
                    let out_name = format!("{}-{}.{}", clean_name, safe_proj_name, ext_str);
                    
                    let out_path = output_dir.join(&out_name);
                    
                    if ext_str == "docx" {
                        if let Err(e) = internal_generate_docx(path.to_str().unwrap(), out_path.to_str().unwrap(), &variables) {
                            println!("Warning: failed to process docx template {}: {}", file_name, e);
                            continue;
                        }
                    } else if ext_str == "xlsx" {
                        // Create a copy of the excel file to output path
                        if let Err(e) = fs::copy(&path, &out_path) {
                            println!("Warning: failed to copy xlsx template {}: {}", file_name, e);
                            continue;
                        }
                        if let Err(e) = internal_generate_xlsx(out_path.to_str().unwrap(), &variables) {
                            println!("Warning: failed to process xlsx template {}: {}", file_name, e);
                            continue;
                        }
                    }
                    
                    generated_count += 1;
                }
            }
        }
    }
    
    if generated_count == 0 {
        return Err("模板目录中未找到任何 .docx 模板文件。".into());
    }
    
    Ok(output_dir.to_string_lossy().to_string())
}

#[tauri::command]
pub fn batch_generate_docx_from_excel(excel_path: String, template_path: String) -> Result<String, String> {
    use calamine::{open_workbook, Reader, Xlsx};
    use chrono::Local;

    let mut workbook: Xlsx<_> = open_workbook(&excel_path).map_err(|e| format!("打开Excel异常: {}", e))?;
    let sheet_names = workbook.sheet_names().to_owned();
    let sheet_name = sheet_names.first().ok_or("找不到工作表")?.clone();
    let range = workbook.worksheet_range(&sheet_name).map_err(|e| format!("读取工作表异常: {}", e))?;

    // Create output directory
    let excel_path_buf = std::path::Path::new(&excel_path);
    let parent = excel_path_buf.parent().unwrap_or(std::path::Path::new("."));
    let output_dir = parent.join("立项签批表生成结果");
    if !output_dir.exists() {
        std::fs::create_dir_all(&output_dir).map_err(|e| format!("创建输出目录失败: {}", e))?;
    }

    let mut headers = HashMap::new();
    let mut rows = range.rows();
    if let Some(header_row) = rows.next() {
        for (i, cell) in header_row.iter().enumerate() {
            let h = cell.to_string().trim().to_string();
            headers.insert(h, i);
        }
    }
    
    println!("Found Excel Headers: {:?}", headers.keys().collect::<Vec<_>>());

    // Mapping: Excel Header Label -> Docx Placeholder Name
    let mapping = [
        ("项目名称", "PROJECT_NAME"),
        ("CT产品名", "CT_PRODUCT_NAME"),
        ("项目总投入(不含税)", "TOTAL_PROJECT_INVESTMENT"),
        ("IT投入(不含税)", "IT_INVESTMENT"),
        ("CT投入(不含税)", "CT_INVESTMENT"),
        ("项目总收入(不含税)", "TOTAL_PROJECT_REVENUE"),
        ("IT收入(不含税)", "IT_REVENUE"),
        ("CT收入(不含税)", "CT_REVENUE"),
        ("项目净现值率", "NET_PRESENT_VALUE_RATE"),
        ("项目毛利率", "PROJECT_GROSS_PROFIT_MARGIN"),
        ("IT净现值率", "IT_NET_PRESENT_VALUE_RATE"),
        ("项目周期", "CONTRACT_DURATION"),
        ("收款方式", "REV_COLLECTION"),
        ("付款方式", "EXP_PAYMENT"),
    ];

    let curr_date = Local::now().format("%Y年%m月%d日").to_string();
    let mut count = 0;

    for row in rows {
        let mut vars = HashMap::new();
        vars.insert("CURR_DATE".to_string(), curr_date.clone());
        
        // Add default values for new dynamic subjects to ensure batch generation doesn't leave un-replaced placeholders
        vars.insert("SUBJECT_IT_COST".to_string(), "IT集成".to_string());
        vars.insert("SUBJECT_CT_COST".to_string(), "CT-专线及产品".to_string());
        vars.insert("SUBJECT_IT_REV".to_string(), "小微ICT业务-IoT-集成".to_string());
        vars.insert("SUBJECT_CT_REV".to_string(), "CT-专线及产品".to_string());
        vars.insert("PROJECT_BACKGROUND".to_string(), "".to_string());

        for (ch_key, en_key) in mapping {
            if let Some(&idx) = headers.get(ch_key) {
                if idx < row.len() {
                    let mut val = row[idx].to_string();
                    
                    // Format percentage rates
                    if en_key.contains("RATE") || en_key.contains("MARGIN") {
                        if let Ok(num) = val.parse::<f64>() {
                            val = format!("{:.2}%", num * 100.0);
                        }
                    }
                    
                    vars.insert(en_key.to_string(), val);
                }
            }
        }
        
        if count == 0 {
           println!("Sample Variables for first row: {:?}", vars);
        }

        let proj_name = vars.get("PROJECT_NAME").cloned().unwrap_or_else(|| format!("未命名_{}", count));
        let safe_proj_name = proj_name.chars().filter(|c| !r#"\/:*?"<>|"#.contains(*c)).collect::<String>();
        let target_name = format!("立项签批表-{}.docx", safe_proj_name);
        let target_path = output_dir.join(target_name);

        internal_generate_docx(&template_path, target_path.to_str().unwrap(), &vars)?;
        count += 1;
    }

    Ok(format!("成功生成 {} 份签批表，保存在目录：\n{}", count, output_dir.display()))
}

fn internal_generate_xlsx(output_path: &str, variables: &HashMap<String, String>) -> Result<(), String> {
    use umya_spreadsheet::*;
    let mut book = reader::xlsx::read(std::path::Path::new(output_path))
        .map_err(|e| format!("无法读取 Excel: {}", e))?;

    if let Some(sheet) = book.get_sheet_by_name_mut("3-直接经济效益评估表") {
        if let Some(v) = variables.get("PROJECT_NAME") {
            let c = sheet.get_cell_mut("D2");
            c.set_value(v);
            c.set_formula("");
        }
        
        let mut set_val = |cell: &str, key: &str| {
            if let Some(v) = variables.get(key) {
                let mut num_str = v.replace(",", "");
                let mut is_pct = false;
                if num_str.ends_with('%') {
                    num_str = num_str.trim_end_matches('%').to_string();
                    is_pct = true;
                }
                
                let cell_obj = sheet.get_cell_mut(cell);
                // Do NOT clear formula if we are just filling inputs. Wait, if it's an input cell, we should clear formula if any? No, we only target cells we know. But wait, we target Q10, Q23, Q24, Q25. If they have formulas, we overwrite. If they don't, we overwrite.
                // Wait, if the user explicitly wants to overwrite a cell, we overwrite it. But we MUST NOT overwrite C45-C50, C64-C66.
                
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

        set_val("G3", "EXCEL_REV_IT_INTEGRATION_EXCL");
        set_val("G4", "EXCEL_REV_IT_MAINTENANCE_EXCL");
        set_val("G5", "EXCEL_REV_IT_DEVICE_SALES_EXCL");
        set_val("G6", "EXCEL_REV_IT_DEVICE_LEASE_EXCL");
        set_val("G7", "EXCEL_REV_IT_OTHER_EXCL");
        set_val("G8", "EXCEL_REV_IT_CLOUD_EXCL");
        set_val("G9", "EXCEL_REV_CT_LINE_EXCL");
        set_val("Q10", "EXCEL_REV_CT_PRODUCT_INCL"); // formula in G10 uses Q10
        set_val("G11", "EXCEL_REV_NON_IT_CT_EXCL");

        set_val("G13", "EXCEL_COST_IT_DEVICE_EXCL");
        set_val("G14", "EXCEL_COST_IT_CONSTRUCTION_EXCL");
        set_val("G15", "EXCEL_COST_IT_SURVEY_EXCL");
        set_val("G16", "EXCEL_COST_IT_INTEGRATION_EXCL");
        set_val("G17", "EXCEL_COST_IT_OTHER_EXCL");
        set_val("G18", "EXCEL_COST_IT_MAINTENANCE_EXCL");
        set_val("G19", "EXCEL_COST_IT_RUNNING_EXCL");
        set_val("G20", "EXCEL_COST_IT_BIDDING_EXCL");
        set_val("G21", "EXCEL_COST_IT_DESIGN_EVAL_EXCL");
        set_val("G22", "EXCEL_COST_IT_AUDIT_EXCL");

        set_val("Q23", "EXCEL_COST_CT_CONSTRUCTION_INCL"); // formula in G23 uses Q23
        set_val("Q24", "EXCEL_COST_CT_MAINTENANCE_INCL"); // formula in G24 uses Q24
        set_val("Q25", "EXCEL_COST_CT_OTHER_INCL"); // formula in G25 uses Q25
        set_val("G26", "EXCEL_COST_CT_BANDWIDTH_EXCL");
        set_val("G27", "EXCEL_COST_CT_RENEWAL_EXCL");

        set_val("G28", "EXCEL_COST_NON_IT_CT_EXCL");
        set_val("G29", "EXCEL_COST_MIX_MARKETING_EXCL");
        set_val("G30", "EXCEL_COST_MIX_CHANNEL_EXCL");
        set_val("G31", "EXCEL_COST_MIX_OTHER_EXCL");
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

