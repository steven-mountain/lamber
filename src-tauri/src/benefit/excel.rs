use super::calculator::calculate_benefit;
use super::models::CalcInput;
use crate::config_manager;
use calamine::{open_workbook, Reader, Xlsx};
use rust_xlsxwriter::{Format, Workbook};
use std::collections::HashMap;
use std::path::Path;

#[tauri::command]
pub async fn process_excel_batch(
    state: tauri::State<'_, std::sync::Mutex<config_manager::AppConfig>>,
    module_id: String,
    file_path: String,
) -> Result<String, String> {
    let module_path = {
        let config = state.lock().unwrap();
        config
            .module_paths
            .get(&module_id)
            .ok_or_else(|| "未设置工作目录".to_string())?
            .clone()
    };

    tauri::async_runtime::spawn_blocking(move || {
        let path = Path::new(&file_path);

        if !path.exists() {
            return Err("文件不存在".to_string());
        }

        let output_dir = std::path::Path::new(&module_path).join("output");

        if !output_dir.exists() {
            std::fs::create_dir_all(&output_dir).map_err(|e| format!("创建输出目录失败: {}", e))?;
        }

        let file_name = path.file_name().unwrap().to_string_lossy().to_string();
        let out_name = file_name.replace(".xlsx", "_批处理结果.xlsx");
        let out_path = output_dir.join(out_name);
        let out_path_str = out_path.to_string_lossy().to_string();

        let mut workbook: Xlsx<_> =
            open_workbook(&file_path).map_err(|e| format!("打开Excel异常: {}", e))?;
        let sheet_names = workbook.sheet_names().to_owned();
        let sheet_name = sheet_names.first().ok_or("找不到工作表")?.clone();

        let range = workbook
            .worksheet_range(&sheet_name)
            .map_err(|e| format!("读取工作表异常: {}", e))?;

        let mut out_wb = Workbook::new();
        let out_sheet = out_wb.add_worksheet();

        let mut row_idx = 0;

        let mut headers = vec![];
        let mut has_headers = false;

        let mut inc_col = None;
        let mut cost_col = None;
        let mut margin_col = None;
        let mut npv_col = None;
        let mut ct_amt_col = None;
        let mut it_tax_col = None;
        let mut ct_tax_col = None;

        let percent_format = Format::new().set_num_format("0.00%");

        for row in range.rows() {
            if !has_headers {
                for (c_idx, cell) in row.iter().enumerate() {
                    let h_str = cell.to_string();
                    headers.push(h_str.clone());
                    match h_str.trim() {
                        "项目总收入" | "含税总收入" => inc_col = Some(c_idx),
                        "项目总投入" | "含税总投入" => cost_col = Some(c_idx),
                        "目标利润率" => margin_col = Some(c_idx),
                        "目标净现值率" => npv_col = Some(c_idx),
                        "CT产品含税总额" | "CT产品名" | "CT产品" => {
                            ct_amt_col = Some(c_idx)
                        }
                        "IT税率" => it_tax_col = Some(c_idx),
                        "CT税率" => ct_tax_col = Some(c_idx),
                        _ => {}
                    }
                    out_sheet
                        .write_string(row_idx, c_idx as u16, &h_str)
                        .unwrap();
                }

                let ext_idx = headers.len() as u16;
                out_sheet
                    .write_string(row_idx, ext_idx, "项目总收入(含税)")
                    .unwrap();
                out_sheet
                    .write_string(row_idx, ext_idx + 1, "项目总收入(不含税)")
                    .unwrap();
                out_sheet
                    .write_string(row_idx, ext_idx + 2, "IT收入(不含税)")
                    .unwrap();
                out_sheet
                    .write_string(row_idx, ext_idx + 3, "CT收入(不含税)")
                    .unwrap();
                out_sheet
                    .write_string(row_idx, ext_idx + 4, "项目总投入(含税)")
                    .unwrap();
                out_sheet
                    .write_string(row_idx, ext_idx + 5, "项目总投入(不含税)")
                    .unwrap();
                out_sheet
                    .write_string(row_idx, ext_idx + 6, "IT投入(不含税)")
                    .unwrap();
                out_sheet
                    .write_string(row_idx, ext_idx + 7, "CT投入(不含税)")
                    .unwrap();
                out_sheet
                    .write_string(row_idx, ext_idx + 8, "项目毛利率")
                    .unwrap();
                out_sheet
                    .write_string(row_idx, ext_idx + 9, "项目净现值率")
                    .unwrap();
                out_sheet
                    .write_string(row_idx, ext_idx + 10, "IT净现值率")
                    .unwrap();
                out_sheet
                    .write_string(row_idx, ext_idx + 11, "算账明细/警告")
                    .unwrap();

                has_headers = true;
                row_idx += 1;
                continue;
            }

            // Write original data
            for (c_idx, cell) in row.iter().enumerate() {
                let val = cell.to_string();
                if let Ok(num) = val.parse::<f64>() {
                    out_sheet.write_number(row_idx, c_idx as u16, num).unwrap();
                } else {
                    out_sheet.write_string(row_idx, c_idx as u16, &val).unwrap();
                }
            }

            let ext_idx = headers.len() as u16;

            let get_val = |opt_col: Option<usize>| -> String {
                if let Some(col) = opt_col {
                    if col < row.len() {
                        return row[col].to_string().trim().to_string();
                    }
                }
                "".to_string()
            };

            let inc_val = get_val(inc_col);
            let cost_val = get_val(cost_col);
            let margin_val = get_val(margin_col);
            let npv_val = get_val(npv_col);
            let ct_amt_val = get_val(ct_amt_col);
            let mut it_tax_val = get_val(it_tax_col);
            let mut ct_tax_val = get_val(ct_tax_col);

            if it_tax_val.is_empty() {
                it_tax_val = "0.06".to_string();
            }
            if ct_tax_val.is_empty() {
                ct_tax_val = "0.06".to_string();
            }

            let ct_amt_opt = if ct_amt_val.is_empty() {
                None
            } else {
                Some(ct_amt_val)
            };

            let mut calc_mode = "";
            let mut target_val = "".to_string();

            if !cost_val.is_empty() {
                calc_mode = "total_cost";
                target_val = cost_val;
            } else if !margin_val.is_empty() {
                calc_mode = "margin";
                target_val = margin_val;
            } else if !npv_val.is_empty() {
                calc_mode = "npv";
                target_val = npv_val;
            }

            if inc_val.is_empty() || target_val.is_empty() {
                out_sheet
                    .write_string(row_idx, ext_idx + 3, "跳过：缺少关键参数")
                    .unwrap();
                row_idx += 1;
                continue;
            }

            let input = CalcInput {
                tax_rate_it: it_tax_val,
                tax_rate_ct: ct_tax_val,
                total_income_incl: inc_val,
                calc_mode: calc_mode.to_string(),
                target_value: target_val,
                ct_income_incl_opt: ct_amt_opt,
            };

            match calculate_benefit(input) {
                Ok(res) => {
                    // Backfill inferred values into empty columns
                    if get_val(ct_amt_col).is_empty() {
                        if let Some(col) = ct_amt_col {
                            if let Ok(c) = res.ct_income_incl.parse::<f64>() {
                                out_sheet.write_number(row_idx, col as u16, c).unwrap();
                            }
                        }
                    }
                    if get_val(it_tax_col).is_empty() {
                        if let Some(col) = it_tax_col {
                            out_sheet.write_number(row_idx, col as u16, 0.06).unwrap();
                        }
                    }
                    if get_val(ct_tax_col).is_empty() {
                        if let Some(col) = ct_tax_col {
                            out_sheet.write_number(row_idx, col as u16, 0.06).unwrap();
                        }
                    }

                    if get_val(cost_col).is_empty() {
                        if let Some(col) = cost_col {
                            if let Ok(c) = res.total_cost_incl.parse::<f64>() {
                                out_sheet.write_number(row_idx, col as u16, c).unwrap();
                            }
                        }
                    }
                    if get_val(margin_col).is_empty() {
                        if let Some(col) = margin_col {
                            if let Ok(m) = res.margin_rate.parse::<f64>() {
                                out_sheet
                                    .write_number_with_format(
                                        row_idx,
                                        col as u16,
                                        m,
                                        &percent_format,
                                    )
                                    .unwrap();
                            }
                        }
                    }
                    if get_val(npv_col).is_empty() {
                        if let Some(col) = npv_col {
                            if let Ok(n) = res.npv_rate.parse::<f64>() {
                                out_sheet
                                    .write_number_with_format(
                                        row_idx,
                                        col as u16,
                                        n,
                                        &percent_format,
                                    )
                                    .unwrap();
                            }
                        }
                    }

                    let ext_idx = headers.len() as u16;

                    if let Ok(val) = res.total_income_incl.parse::<f64>() {
                        out_sheet.write_number(row_idx, ext_idx, val).unwrap();
                    }
                    if let Ok(val) = res.total_income_excl.parse::<f64>() {
                        out_sheet.write_number(row_idx, ext_idx + 1, val).unwrap();
                    }
                    if let Ok(val) = res.it_income_excl.parse::<f64>() {
                        out_sheet.write_number(row_idx, ext_idx + 2, val).unwrap();
                    }
                    if let Ok(val) = res.ct_income_excl.parse::<f64>() {
                        out_sheet.write_number(row_idx, ext_idx + 3, val).unwrap();
                    }

                    if let Ok(val) = res.total_cost_incl.parse::<f64>() {
                        out_sheet.write_number(row_idx, ext_idx + 4, val).unwrap();
                    }
                    if let Ok(val) = res.total_cost_excl.parse::<f64>() {
                        out_sheet.write_number(row_idx, ext_idx + 5, val).unwrap();
                    }
                    if let Ok(val) = res.it_cost_excl.parse::<f64>() {
                        out_sheet.write_number(row_idx, ext_idx + 6, val).unwrap();
                    }
                    if let Ok(val) = res.ct_cost_excl.parse::<f64>() {
                        out_sheet.write_number(row_idx, ext_idx + 7, val).unwrap();
                    }

                    if let Ok(m) = res.margin_rate.parse::<f64>() {
                        out_sheet
                            .write_number_with_format(row_idx, ext_idx + 8, m, &percent_format)
                            .unwrap();
                    }
                    if let Ok(n) = res.npv_rate.parse::<f64>() {
                        out_sheet
                            .write_number_with_format(row_idx, ext_idx + 9, n, &percent_format)
                            .unwrap();
                    }
                    if let Ok(it_n) = res.it_npv_rate.parse::<f64>() {
                        out_sheet
                            .write_number_with_format(row_idx, ext_idx + 10, it_n, &percent_format)
                            .unwrap();
                    }
                    let warn = res.warning_message.unwrap_or_else(|| "正常".to_string());
                    out_sheet
                        .write_string(row_idx, ext_idx + 11, &warn)
                        .unwrap();
                }
                Err(e) => {
                    let ext_idx = headers.len() as u16;
                    out_sheet
                        .write_string(row_idx, ext_idx + 11, &format!("错误: {}", e))
                        .unwrap();
                }
            }

            row_idx += 1;
        }

        out_wb
            .save(&out_path_str)
            .map_err(|e| format!("保存文件失败: {}", e))?;
        Ok(out_path_str)
    })
    .await
    .map_err(|e| format!("异步执行异常: {}", e))?
}

#[tauri::command]
pub fn generate_excel_template(path: String) -> Result<(), String> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    let headers = [
        "项目名称",
        "含税总收入",
        "含税总投入",
        "目标利润率",
        "目标净现值率",
        "项目周期",
        "CT产品名",
        "CT产品含税总额",
        "IT税率",
        "CT税率",
        "收款方式",
        "付款方式",
    ];
    for (col, header) in headers.iter().enumerate() {
        worksheet.write_string(0, col as u16, *header).unwrap();
    }

    worksheet.write_string(1, 0, "示例项目A").unwrap();
    worksheet.write_number(1, 1, 1000000.0).unwrap();
    worksheet.write_number(1, 2, 800000.0).unwrap();
    worksheet.write_string(1, 3, "").unwrap();
    worksheet.write_string(1, 4, "").unwrap();
    worksheet.write_string(1, 5, "1").unwrap();
    worksheet.write_string(1, 6, "示例产品").unwrap();
    worksheet.write_string(1, 7, "").unwrap();
    worksheet.write_number(1, 8, 0.06).unwrap();
    worksheet.write_number(1, 9, 0.06).unwrap();
    worksheet.write_string(1, 10, "合同签订后支付XX%").unwrap();
    worksheet.write_string(1, 11, "背靠背支付").unwrap();

    workbook
        .save(&path)
        .map_err(|e| format!("写入模板失败: {}", e))?;
    Ok(())
}

#[derive(serde::Serialize, Clone)]
pub struct ExcelParsedTaxItem {
    pub incl_tax: f64,
    pub excl_tax: f64,
    pub tax_rate: f64,
}

#[derive(serde::Serialize, Clone)]
pub struct ExcelParsedData {
    pub project_name: String,
    pub customer_name: String,
    pub total_income_incl: f64,
    pub total_cost_incl: f64,
    pub target_margin: f64,
    pub target_npv: f64,
    pub project_years: i32,
    pub discount_rate: f64,
    pub ct_name: String,
    pub ct_income_incl: f64,
    pub it_tax: f64,
    pub ct_tax: f64,
    pub payment_collect: String,
    pub payment_pay: String,
    pub items: HashMap<String, ExcelParsedTaxItem>,
}

fn parse_cell_f64(cell: &calamine::Data) -> f64 {
    match cell {
        calamine::Data::Float(f) => *f,
        calamine::Data::Int(i) => *i as f64,
        calamine::Data::String(s) => {
            let s_trimmed = s
                .trim()
                .replace(",", "")
                .replace("，", "")
                .replace("元", "")
                .replace("%", "");
            if let Ok(num) = s_trimmed.parse::<f64>() {
                if s.contains('%') {
                    num / 100.0
                } else {
                    num
                }
            } else {
                0.0
            }
        }
        _ => 0.0,
    }
}

fn parse_cell_i32(cell: &calamine::Data) -> i32 {
    match cell {
        calamine::Data::Int(i) => *i as i32,
        calamine::Data::Float(f) => *f as i32,
        calamine::Data::String(s) => s.trim().parse::<i32>().unwrap_or(0),
        _ => 0,
    }
}

fn parse_cell_string(cell: &calamine::Data) -> String {
    match cell {
        calamine::Data::String(s) => s.trim().to_string(),
        calamine::Data::Float(f) => f.to_string(),
        calamine::Data::Int(i) => i.to_string(),
        calamine::Data::Bool(b) => b.to_string(),
        _ => "".to_string(),
    }
}

fn cell_ref_to_index(cell_ref: &str) -> Option<(u32, u32)> {
    let mut col: u32 = 0;
    let mut row = String::new();

    for ch in cell_ref.chars() {
        if ch.is_ascii_alphabetic() {
            col = col * 26 + (ch.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
        } else if ch.is_ascii_digit() {
            row.push(ch);
        }
    }

    let row_num = row.parse::<u32>().ok()?;
    if row_num == 0 || col == 0 {
        return None;
    }

    Some((row_num - 1, col - 1))
}

fn get_cell<'a>(
    range: &'a calamine::Range<calamine::Data>,
    cell_ref: &str,
) -> Option<&'a calamine::Data> {
    let (row, col) = cell_ref_to_index(cell_ref)?;
    range.get_value((row, col))
}

fn get_cell_f64(range: &calamine::Range<calamine::Data>, cell_ref: &str) -> f64 {
    get_cell(range, cell_ref).map(parse_cell_f64).unwrap_or(0.0)
}

fn get_cell_i32(range: &calamine::Range<calamine::Data>, cell_ref: &str) -> i32 {
    get_cell(range, cell_ref).map(parse_cell_i32).unwrap_or(0)
}

fn get_cell_string(range: &calamine::Range<calamine::Data>, cell_ref: &str) -> String {
    get_cell(range, cell_ref)
        .map(parse_cell_string)
        .unwrap_or_default()
}

fn normalize_tax_percent(tax_rate: f64) -> f64 {
    if tax_rate > 0.0 && tax_rate < 1.0 {
        tax_rate * 100.0
    } else {
        tax_rate
    }
}

fn item_from_cells(
    range: &calamine::Range<calamine::Data>,
    incl_cell: &str,
    excl_cell: &str,
    tax_rate: f64,
) -> ExcelParsedTaxItem {
    let tax = normalize_tax_percent(tax_rate);
    let incl = get_cell_f64(range, incl_cell);
    let excl = get_cell_f64(range, excl_cell);

    let (incl_tax, excl_tax) = if incl.abs() > f64::EPSILON {
        (incl, incl / (1.0 + tax / 100.0))
    } else if excl.abs() > f64::EPSILON {
        (excl * (1.0 + tax / 100.0), excl)
    } else {
        (0.0, 0.0)
    };

    ExcelParsedTaxItem {
        incl_tax,
        excl_tax,
        tax_rate: tax,
    }
}

fn parse_lifecycle_benefit_excel(
    workbook: &mut Xlsx<std::io::BufReader<std::fs::File>>,
) -> Result<Option<ExcelParsedData>, String> {
    let sheet_names = workbook.sheet_names().to_owned();
    if !sheet_names
        .iter()
        .any(|name| name == "3-直接经济效益评估表")
    {
        return Ok(None);
    }

    let econ_range = workbook
        .worksheet_range("3-直接经济效益评估表")
        .map_err(|e| format!("读取效益评估表异常: {}", e))?;

    let result_range = if sheet_names.iter().any(|name| name == "2-ICT项目评估结果") {
        Some(
            workbook
                .worksheet_range("2-ICT项目评估结果")
                .map_err(|e| format!("读取评估结果表异常: {}", e))?,
        )
    } else {
        None
    };

    let mut items = HashMap::new();
    let mappings = [
        ("rev_it_integration", "Q3", "G3", 6.0),
        ("rev_it_maintenance", "Q4", "G4", 6.0),
        ("rev_it_device_sales", "Q5", "G5", 13.0),
        ("rev_it_device_lease", "Q6", "G6", 13.0),
        ("rev_it_other", "Q7", "G7", 6.0),
        ("rev_it_cloud", "Q8", "G8", 6.0),
        ("rev_ct_line", "Q9", "G9", 9.0),
        ("rev_ct_product", "Q10", "G10", 6.0),
        ("rev_non_it_ct", "Q11", "G11", 9.0),
        ("cost_it_device", "Q13", "G13", 13.0),
        ("cost_it_construction", "Q14", "G14", 9.0),
        ("cost_it_survey", "Q15", "G15", 6.0),
        ("cost_it_integration", "Q16", "G16", 6.0),
        ("cost_it_other", "Q17", "G17", 6.0),
        ("cost_it_maintenance", "Q18", "G18", 6.0),
        ("cost_it_running", "Q19", "G19", 13.0),
        ("cost_it_bidding", "Q20", "G20", 6.0),
        ("cost_it_design_eval", "Q21", "G21", 6.0),
        ("cost_it_audit", "Q22", "G22", 6.0),
        ("cost_ct_construction", "Q23", "G23", 6.0),
        ("cost_ct_maintenance", "Q24", "G24", 9.0),
        ("cost_ct_other", "Q25", "G25", 6.0),
        ("cost_ct_bandwidth", "Q26", "G26", 9.0),
        ("cost_ct_renewal", "Q27", "G27", 9.0),
        ("cost_non_it_ct", "Q28", "G28", 9.0),
        ("cost_mix_marketing", "Q29", "G29", 6.0),
        ("cost_mix_channel", "Q30", "G30", 6.0),
        ("cost_mix_other", "Q31", "G31", 6.0),
    ];

    for (key, incl_cell, excl_cell, tax_rate) in mappings {
        items.insert(
            key.to_string(),
            item_from_cells(&econ_range, incl_cell, excl_cell, tax_rate),
        );
    }

    let total_income_incl = [
        "rev_it_integration",
        "rev_it_maintenance",
        "rev_it_device_sales",
        "rev_it_device_lease",
        "rev_it_other",
        "rev_it_cloud",
        "rev_ct_line",
        "rev_ct_product",
        "rev_non_it_ct",
    ]
    .iter()
    .filter_map(|key| items.get(*key))
    .map(|item| item.incl_tax)
    .sum();

    let total_cost_incl = [
        "cost_it_device",
        "cost_it_construction",
        "cost_it_survey",
        "cost_it_integration",
        "cost_it_other",
        "cost_it_maintenance",
        "cost_it_running",
        "cost_it_bidding",
        "cost_it_design_eval",
        "cost_it_audit",
        "cost_ct_construction",
        "cost_ct_maintenance",
        "cost_ct_other",
        "cost_ct_bandwidth",
        "cost_ct_renewal",
        "cost_non_it_ct",
        "cost_mix_marketing",
        "cost_mix_channel",
        "cost_mix_other",
    ]
    .iter()
    .filter_map(|key| items.get(*key))
    .map(|item| item.incl_tax)
    .sum();

    let project_name_from_econ = get_cell_string(&econ_range, "D2");
    let project_name_from_result = result_range
        .as_ref()
        .map(|range| get_cell_string(range, "B4"))
        .unwrap_or_default();
    let customer_name = result_range
        .as_ref()
        .map(|range| get_cell_string(range, "B5"))
        .unwrap_or_default();
    let project_years = result_range
        .as_ref()
        .map(|range| get_cell_i32(range, "B8"))
        .unwrap_or(0);
    let discount_rate = get_cell_f64(&econ_range, "D33");

    Ok(Some(ExcelParsedData {
        project_name: if project_name_from_econ.is_empty() {
            project_name_from_result
        } else {
            project_name_from_econ
        },
        customer_name,
        total_income_incl,
        total_cost_incl,
        target_margin: 0.0,
        target_npv: 0.0,
        project_years: if project_years > 0 { project_years } else { 1 },
        discount_rate: if discount_rate > 0.0 {
            discount_rate
        } else {
            0.055
        },
        ct_name: String::new(),
        ct_income_incl: items
            .get("rev_ct_product")
            .map(|item| item.incl_tax)
            .unwrap_or(0.0),
        it_tax: 0.06,
        ct_tax: 0.06,
        payment_collect: String::new(),
        payment_pay: String::new(),
        items,
    }))
}

#[tauri::command]
pub async fn parse_benefit_excel(file_path: String) -> Result<ExcelParsedData, String> {
    tauri::async_runtime::spawn_blocking(move || {
        let path = Path::new(&file_path);
        if !path.exists() {
            return Err("文件不存在".to_string());
        }

        let mut workbook: Xlsx<_> =
            open_workbook(&file_path).map_err(|e| format!("打开Excel异常: {}", e))?;
        if let Some(parsed) = parse_lifecycle_benefit_excel(&mut workbook)? {
            return Ok(parsed);
        }

        let sheet_names = workbook.sheet_names().to_owned();
        let sheet_name = sheet_names.first().ok_or("找不到工作表")?.clone();
        let range = workbook
            .worksheet_range(&sheet_name)
            .map_err(|e| format!("读取工作表异常: {}", e))?;

        let mut rows = range.rows();
        let headers_row = rows.next().ok_or("Excel表为空")?;

        let mut proj_name_col = None;
        let mut inc_col = None;
        let mut cost_col = None;
        let mut margin_col = None;
        let mut npv_col = None;
        let mut years_col = None;
        let mut ct_name_col = None;
        let mut ct_amt_col = None;
        let mut it_tax_col = None;
        let mut ct_tax_col = None;
        let mut pay_collect_col = None;
        let mut pay_pay_col = None;

        for (c_idx, cell) in headers_row.iter().enumerate() {
            let h_str = cell.to_string().trim().to_string();
            match h_str.as_str() {
                "项目名称" => proj_name_col = Some(c_idx),
                "含税总收入" | "项目总收入" => inc_col = Some(c_idx),
                "含税总投入" | "项目总投入" => cost_col = Some(c_idx),
                "目标利润率" => margin_col = Some(c_idx),
                "目标净现值率" => npv_col = Some(c_idx),
                "项目周期" | "周期" => years_col = Some(c_idx),
                "CT产品名" | "CT产品" => ct_name_col = Some(c_idx),
                "CT产品含税总额" => ct_amt_col = Some(c_idx),
                "IT税率" => it_tax_col = Some(c_idx),
                "CT税率" => ct_tax_col = Some(c_idx),
                "收款方式" => pay_collect_col = Some(c_idx),
                "付款方式" => pay_pay_col = Some(c_idx),
                _ => {}
            }
        }

        let first_data_row = rows.next().ok_or("找不到数据行")?;

        let get_string = |col_opt: Option<usize>| -> String {
            if let Some(col) = col_opt {
                if col < first_data_row.len() {
                    return parse_cell_string(&first_data_row[col]);
                }
            }
            "".to_string()
        };

        let get_f64 = |col_opt: Option<usize>| -> f64 {
            if let Some(col) = col_opt {
                if col < first_data_row.len() {
                    return parse_cell_f64(&first_data_row[col]);
                }
            }
            0.0
        };

        let get_i32 = |col_opt: Option<usize>| -> i32 {
            if let Some(col) = col_opt {
                if col < first_data_row.len() {
                    return parse_cell_i32(&first_data_row[col]);
                }
            }
            0
        };

        Ok(ExcelParsedData {
            project_name: get_string(proj_name_col),
            customer_name: String::new(),
            total_income_incl: get_f64(inc_col),
            total_cost_incl: get_f64(cost_col),
            target_margin: get_f64(margin_col),
            target_npv: get_f64(npv_col),
            project_years: get_i32(years_col),
            discount_rate: 0.055,
            ct_name: get_string(ct_name_col),
            ct_income_incl: get_f64(ct_amt_col),
            it_tax: get_f64(it_tax_col),
            ct_tax: get_f64(ct_tax_col),
            payment_collect: get_string(pay_collect_col),
            payment_pay: get_string(pay_pay_col),
            items: HashMap::new(),
        })
    })
    .await
    .map_err(|e| format!("异步执行异常: {}", e))?
}
