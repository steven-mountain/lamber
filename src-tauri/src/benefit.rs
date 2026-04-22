use rust_decimal::prelude::*;
use serde::{Deserialize, Serialize};
use std::path::Path;
use calamine::{open_workbook, Reader, Xlsx};
use rust_xlsxwriter::{Workbook, Format};

#[derive(Deserialize, Clone)]
pub struct CalcInput {
    pub tax_rate_it: String,
    pub tax_rate_ct: String,
    pub total_income_incl: String,
    pub calc_mode: String,  // "margin", "npv", or "total_cost"
    pub target_value: String,
    pub ct_income_incl_opt: Option<String>,
}

#[derive(Serialize)]
pub struct CalcResult {
    pub it_income_incl: String,
    pub ct_income_incl: String,
    pub total_income_incl: String,
    pub it_income_excl: String,
    pub ct_income_excl: String,
    pub total_income_excl: String,
    pub it_cost_incl: String,
    pub ct_cost_incl: String,
    pub total_cost_incl: String,
    pub it_cost_excl: String,
    pub ct_cost_excl: String,
    pub total_cost_excl: String,
    pub margin_rate: String,
    pub npv_rate: String,
    pub it_npv_rate: String,
    pub warning_message: Option<String>,
}

fn round_2(val: Decimal) -> Decimal {
    val.round_dp(2)
}

fn round_4(val: Decimal) -> Decimal {
    val.round_dp(4)
}

#[tauri::command]
pub fn calculate_benefit(input: CalcInput) -> Result<CalcResult, String> {
    let d1 = Decimal::ONE;
    let d72 = Decimal::new(72, 0);
    let d0_01 = Decimal::new(1, 2);

    let tax_rate_it = Decimal::from_str(&input.tax_rate_it).unwrap_or(Decimal::new(6, 2));
    let tax_rate_ct = Decimal::from_str(&input.tax_rate_ct).unwrap_or(Decimal::new(6, 2));
    let total_income_incl = Decimal::from_str(&input.total_income_incl).map_err(|e| e.to_string())?;
    let target_value = Decimal::from_str(&input.target_value).map_err(|e| e.to_string())?;

    // --- 第一步：含税盘子分配 ---
    let ct_income_incl = if let Some(ct_str) = &input.ct_income_incl_opt {
        if ct_str.trim().is_empty() {
             let ct_income_incl_min = round_2(total_income_incl * d0_01);
             let ceil_multiplier = (ct_income_incl_min / d72).ceil();
             round_2(ceil_multiplier * d72)
        } else {
             Decimal::from_str(ct_str).unwrap_or_else(|_| {
                 let ct_income_incl_min = round_2(total_income_incl * d0_01);
                 let ceil_multiplier = (ct_income_incl_min / d72).ceil();
                 round_2(ceil_multiplier * d72)
             })
        }
    } else {
        let ct_income_incl_min = round_2(total_income_incl * d0_01);
        let ceil_multiplier = (ct_income_incl_min / d72).ceil();
        round_2(ceil_multiplier * d72)
    };
    
    let ct_cost_incl = ct_income_incl;
    let it_income_incl = round_2(total_income_incl - ct_income_incl);

    // --- 第二步：价税分离 ---
    let it_income_excl = round_2(it_income_incl / (d1 + tax_rate_it));
    let ct_income_excl = round_2(ct_income_incl / (d1 + tax_rate_ct));
    let total_income_excl = it_income_excl + ct_income_excl;
    let ct_cost_excl = round_2(ct_cost_incl / (d1 + tax_rate_ct));

    // --- 第三步：测算投入 ---
    let total_cost_excl;
    let it_cost_excl;
    let it_cost_incl;
    let total_cost_incl;

    if input.calc_mode == "margin" {
        total_cost_excl = round_2(total_income_excl * (d1 - target_value));
        it_cost_excl = total_cost_excl - ct_cost_excl;
        it_cost_incl = round_2(it_cost_excl * (d1 + tax_rate_it));
        total_cost_incl = it_cost_incl + ct_cost_incl;
    } else if input.calc_mode == "npv" {
        total_cost_excl = round_2(total_income_excl / (d1 + target_value));
        it_cost_excl = total_cost_excl - ct_cost_excl;
        it_cost_incl = round_2(it_cost_excl * (d1 + tax_rate_it));
        total_cost_incl = it_cost_incl + ct_cost_incl;
    } else if input.calc_mode == "total_cost" {
        total_cost_incl = target_value; 
        it_cost_incl = total_cost_incl - ct_cost_incl;
        it_cost_excl = round_2(it_cost_incl / (d1 + tax_rate_it));
        total_cost_excl = it_cost_excl + ct_cost_excl;
    } else {
        return Err("未知的计算模式".to_string());
    }

    // --- 第四步：效益指标核算 ---
    let margin_rate = if total_income_excl.is_zero() {
        Decimal::ZERO
    } else {
        round_4((total_income_excl - total_cost_excl) / total_income_excl)
    };

    let npv_rate = if total_cost_excl.is_zero() {
        Decimal::ZERO
    } else {
        round_4((total_income_excl - total_cost_excl) / total_cost_excl)
    };

    let it_npv_rate = if it_cost_excl.is_zero() {
        Decimal::ZERO
    } else {
        round_4((it_income_excl - it_cost_excl) / it_cost_excl)
    };

    let mut warnings = Vec::new();
    if it_cost_incl < Decimal::ZERO || it_cost_excl < Decimal::ZERO {
        warnings.push("目标太高或投入太低，IT投入已被穿透为负数".to_string());
    }

    let warning_message = if warnings.is_empty() {
        None
    } else {
        Some(warnings.join(" | "))
    };

    Ok(CalcResult {
        it_income_incl: it_income_incl.to_string(),
        ct_income_incl: ct_income_incl.to_string(),
        total_income_incl: total_income_incl.to_string(),
        it_income_excl: it_income_excl.to_string(),
        ct_income_excl: ct_income_excl.to_string(),
        total_income_excl: total_income_excl.to_string(),
        it_cost_incl: it_cost_incl.to_string(),
        ct_cost_incl: ct_cost_incl.to_string(),
        total_cost_incl: total_cost_incl.to_string(),
        it_cost_excl: it_cost_excl.to_string(),
        ct_cost_excl: ct_cost_excl.to_string(),
        total_cost_excl: total_cost_excl.to_string(),
        margin_rate: margin_rate.to_string(),
        npv_rate: npv_rate.to_string(),
        it_npv_rate: it_npv_rate.to_string(),
        warning_message,
    })
}

#[tauri::command]
pub async fn process_excel_batch(file_path: String) -> Result<String, String> {
    let path = Path::new(&file_path);

    if !path.exists() {
        return Err("文件不存在".to_string());
    }

    let mut workbook: Xlsx<_> = open_workbook(&file_path).map_err(|e| format!("打开Excel异常: {}", e))?;
    let sheet_names = workbook.sheet_names().to_owned();
    let sheet_name = sheet_names.first().ok_or("找不到工作表")?.clone();

    let range = workbook.worksheet_range(&sheet_name).map_err(|e| format!("读取工作表异常: {}", e))?;

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
                    "CT产品含税总额" | "CT产品名" | "CT产品" => ct_amt_col = Some(c_idx),
                    "IT税率" => it_tax_col = Some(c_idx),
                    "CT税率" => ct_tax_col = Some(c_idx),
                    _ => {}
                }
                out_sheet.write_string(row_idx, c_idx as u16, &h_str).unwrap();
            }
            
            let ext_idx = headers.len() as u16;
            out_sheet.write_string(row_idx, ext_idx, "项目总收入(含税)").unwrap();
            out_sheet.write_string(row_idx, ext_idx + 1, "项目总收入(不含税)").unwrap();
            out_sheet.write_string(row_idx, ext_idx + 2, "IT收入(不含税)").unwrap();
            out_sheet.write_string(row_idx, ext_idx + 3, "CT收入(不含税)").unwrap();
            out_sheet.write_string(row_idx, ext_idx + 4, "项目总投入(含税)").unwrap();
            out_sheet.write_string(row_idx, ext_idx + 5, "项目总投入(不含税)").unwrap();
            out_sheet.write_string(row_idx, ext_idx + 6, "IT投入(不含税)").unwrap();
            out_sheet.write_string(row_idx, ext_idx + 7, "CT投入(不含税)").unwrap();
            out_sheet.write_string(row_idx, ext_idx + 8, "项目毛利率").unwrap();
            out_sheet.write_string(row_idx, ext_idx + 9, "项目净现值率").unwrap();
            out_sheet.write_string(row_idx, ext_idx + 10, "IT净现值率").unwrap();
            out_sheet.write_string(row_idx, ext_idx + 11, "算账明细/警告").unwrap();
            
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

        if it_tax_val.is_empty() { it_tax_val = "0.06".to_string(); }
        if ct_tax_val.is_empty() { ct_tax_val = "0.06".to_string(); }

        let ct_amt_opt = if ct_amt_val.is_empty() { None } else { Some(ct_amt_val) };

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
            out_sheet.write_string(row_idx, ext_idx + 3, "跳过：缺少关键参数").unwrap();
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
                            out_sheet.write_number_with_format(row_idx, col as u16, m, &percent_format).unwrap();
                        }
                    }
                }
                if get_val(npv_col).is_empty() {
                    if let Some(col) = npv_col {
                        if let Ok(n) = res.npv_rate.parse::<f64>() {
                            out_sheet.write_number_with_format(row_idx, col as u16, n, &percent_format).unwrap();
                        }
                    }
                }

                let ext_idx = headers.len() as u16;

                // Detail columns: 项目总收入(含税), 项目总收入(不含税), IT收入(不含税), CT收入(不含税),
                // 项目总投入(含税), 项目总投入(不含税), IT投入(不含税), CT投入(不含税),
                // 项目毛利率, 项目净现值率, IT净现值率, 算账明细/警告
                if let Ok(val) = res.total_income_incl.parse::<f64>() { out_sheet.write_number(row_idx, ext_idx, val).unwrap(); }
                if let Ok(val) = res.total_income_excl.parse::<f64>() { out_sheet.write_number(row_idx, ext_idx + 1, val).unwrap(); }
                if let Ok(val) = res.it_income_excl.parse::<f64>() { out_sheet.write_number(row_idx, ext_idx + 2, val).unwrap(); }
                if let Ok(val) = res.ct_income_excl.parse::<f64>() { out_sheet.write_number(row_idx, ext_idx + 3, val).unwrap(); }
                
                if let Ok(val) = res.total_cost_incl.parse::<f64>() { out_sheet.write_number(row_idx, ext_idx + 4, val).unwrap(); }
                if let Ok(val) = res.total_cost_excl.parse::<f64>() { out_sheet.write_number(row_idx, ext_idx + 5, val).unwrap(); }
                if let Ok(val) = res.it_cost_excl.parse::<f64>() { out_sheet.write_number(row_idx, ext_idx + 6, val).unwrap(); }
                if let Ok(val) = res.ct_cost_excl.parse::<f64>() { out_sheet.write_number(row_idx, ext_idx + 7, val).unwrap(); }
                
                if let Ok(m) = res.margin_rate.parse::<f64>() {
                    out_sheet.write_number_with_format(row_idx, ext_idx + 8, m, &percent_format).unwrap();
                }
                if let Ok(n) = res.npv_rate.parse::<f64>() {
                    out_sheet.write_number_with_format(row_idx, ext_idx + 9, n, &percent_format).unwrap();
                }
                if let Ok(it_n) = res.it_npv_rate.parse::<f64>() {
                    out_sheet.write_number_with_format(row_idx, ext_idx + 10, it_n, &percent_format).unwrap();
                }
                let warn = res.warning_message.unwrap_or("正常".to_string());
                out_sheet.write_string(row_idx, ext_idx + 11, &warn).unwrap();
            },
            Err(e) => {
                let ext_idx = headers.len() as u16;
                out_sheet.write_string(row_idx, ext_idx + 11, &format!("错误: {}", e)).unwrap();
            }
        }
        
        row_idx += 1;
    }

    let out_name = file_path.replace(".xlsx", "_批处理结果.xlsx");
    out_wb.save(&out_name).map_err(|e| format!("保存文件失败: {}", e))?;

    Ok(out_name)
}

#[tauri::command]
pub fn generate_excel_template(path: String) -> Result<(), String> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    
    let headers = ["项目名称", "含税总收入", "含税总投入", "目标利润率", "目标净现值率", "项目周期", "CT产品名", "CT产品含税总额", "IT税率", "CT税率"];
    for (col, header) in headers.iter().enumerate() {
        worksheet.write_string(0, col as u16, *header).unwrap();
    }
    
    // Add sample data for clarity
    worksheet.write_string(1, 0, "示例项目A").unwrap();
    worksheet.write_number(1, 1, 1000000.0).unwrap();
    worksheet.write_number(1, 2, 800000.0).unwrap();
    worksheet.write_string(1, 3, "").unwrap(); // 目标利润率
    worksheet.write_string(1, 4, "").unwrap(); // 目标净现值率
    worksheet.write_string(1, 5, "1").unwrap(); // 项目周期
    worksheet.write_string(1, 6, "示例产品").unwrap(); // CT产品名
    worksheet.write_string(1, 7, "").unwrap(); // CT产品含税总额
    worksheet.write_number(1, 8, 0.06).unwrap(); // IT税率
    worksheet.write_number(1, 9, 0.06).unwrap(); // CT税率
    
    workbook.save(&path).map_err(|e| format!("写入模板失败: {}", e))?;
    Ok(())
}
