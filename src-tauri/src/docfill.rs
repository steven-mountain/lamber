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
        if file.name() == "word/document.xml" || file.name() == "word/header1.xml" || file.name() == "word/footer1.xml" {
            let mut content = String::new();
            if file.read_to_string(&mut content).is_ok() {
                doc_xml.push_str(&content);
            }
        }
    }

    if doc_xml.is_empty() {
        return Err("Could not find document.xml in the provided docx file.".into());
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

                for (k, v) in variables {
                    let pattern = format!("{{{}}}", k);
                    xml_str = xml_str.replace(&pattern, v);
                }

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
    ];

    let curr_date = Local::now().format("%Y年%m月%d日").to_string();
    let mut count = 0;

    for row in rows {
        let mut vars = HashMap::new();
        vars.insert("CURR_DATE".to_string(), curr_date.clone());

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
