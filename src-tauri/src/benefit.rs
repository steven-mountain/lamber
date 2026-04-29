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

#[derive(Deserialize, Clone)]
pub struct IctItem {
    pub incl_tax: String,
    pub tax_rate: String,
}

#[derive(Deserialize, Clone)]
pub struct IctInput {
    pub project_name: String,
    pub property_rights: String,
    pub discount_rate: String,
    
    // The revenue and cost distributions over 10 years (e.g., [1.0, 0.0, ..., 0.0])
    pub rev_distribution: Vec<f64>,
    pub cost_distribution: Vec<f64>,
    
    pub rev_it_integration: IctItem,
    pub rev_it_maintenance: IctItem,
    pub rev_it_device_sales: IctItem,
    pub rev_it_device_lease: IctItem,
    pub rev_it_other: IctItem,
    pub rev_it_cloud: IctItem,
    
    pub rev_ct_line: IctItem,
    pub rev_ct_product: IctItem,
    
    pub rev_non_it_ct: IctItem,

    pub cost_it_device: IctItem,
    pub cost_it_construction: IctItem,
    pub cost_it_survey: IctItem,
    pub cost_it_integration: IctItem,
    pub cost_it_other: IctItem,
    pub cost_it_maintenance: IctItem,
    pub cost_it_running: IctItem,
    pub cost_it_bidding: IctItem,
    pub cost_it_design_eval: IctItem,
    pub cost_it_audit: IctItem,
    
    pub cost_ct_construction: IctItem,
    pub cost_ct_maintenance: IctItem,
    pub cost_ct_other: IctItem,
    pub cost_ct_bandwidth: IctItem,
    pub cost_ct_renewal: IctItem,
    
    pub cost_non_it_ct: IctItem,
    pub cost_mix_marketing: IctItem,
    pub cost_mix_channel: IctItem,
    pub cost_mix_other: IctItem,
}

#[derive(Serialize)]
pub struct IctCashflowRow {
    pub year: i32,
    pub cash_in: String,
    pub cash_out: String,
    pub net_cash: String,
    pub cum_net_cash: String,
    pub pv: String,
    pub cum_pv: String,
}

#[derive(Serialize)]
pub struct IctResult {
    pub npv: String,
    pub npv_rate: String,
    pub margin_rate: String,
    pub dynamic_payback: String,
    pub irr: String,
    
    pub it_npv: String,
    pub it_npv_rate: String,
    pub it_margin_rate: String,
    
    pub cashflow: Vec<IctCashflowRow>,
}

fn get_excl(item: &IctItem) -> Decimal {
    let incl = Decimal::from_str(&item.incl_tax).unwrap_or(Decimal::ZERO);
    let rate = Decimal::from_str(&item.tax_rate).unwrap_or(Decimal::ZERO) / Decimal::new(100, 0);
    if incl.is_zero() {
        return Decimal::ZERO;
    }
    (incl / (Decimal::ONE + rate)).round_dp(2)
}

#[tauri::command]
pub fn calculate_ict_benefit(input: IctInput) -> Result<IctResult, String> {
    let discount_rate = Decimal::from_str(&input.discount_rate).unwrap_or(Decimal::new(55, 3)); // 0.055

    // IT Revenue
    let it_rev = get_excl(&input.rev_it_integration)
        + get_excl(&input.rev_it_maintenance)
        + get_excl(&input.rev_it_device_sales)
        + get_excl(&input.rev_it_device_lease)
        + get_excl(&input.rev_it_other)
        + get_excl(&input.rev_it_cloud);

    // CT Revenue
    let ct_rev = get_excl(&input.rev_ct_line)
        + get_excl(&input.rev_ct_product);

    // Non-IT Revenue
    let non_it_rev = get_excl(&input.rev_non_it_ct);

    let total_rev = it_rev + ct_rev + non_it_rev;

    // IT Cost
    let it_cost = get_excl(&input.cost_it_device)
        + get_excl(&input.cost_it_construction)
        + get_excl(&input.cost_it_survey)
        + get_excl(&input.cost_it_integration)
        + get_excl(&input.cost_it_other)
        + get_excl(&input.cost_it_maintenance)
        + get_excl(&input.cost_it_running)
        + get_excl(&input.cost_it_bidding)
        + get_excl(&input.cost_it_design_eval)
        + get_excl(&input.cost_it_audit);

    // CT Cost
    let ct_cost = get_excl(&input.cost_ct_construction)
        + get_excl(&input.cost_ct_maintenance)
        + get_excl(&input.cost_ct_other)
        + get_excl(&input.cost_ct_bandwidth)
        + get_excl(&input.cost_ct_renewal);

    // Non-IT Cost & Mix Cost
    let non_it_cost = get_excl(&input.cost_non_it_ct);
    let mix_cost = get_excl(&input.cost_mix_marketing)
        + get_excl(&input.cost_mix_channel)
        + get_excl(&input.cost_mix_other);

    let total_cost = it_cost + ct_cost + non_it_cost + mix_cost;

    let margin_rate = if total_rev.is_zero() { Decimal::ZERO } else { ((total_rev - total_cost) / total_rev).round_dp(4) };
    let it_margin_rate = if it_rev.is_zero() { Decimal::ZERO } else { ((it_rev - it_cost) / it_rev).round_dp(4) };

    // --- 10 Year Cashflow Simulation ---
    let mut cashflow = Vec::new();
    
    let mut cum_net_cash = Decimal::ZERO;
    let mut cum_pv = Decimal::ZERO;
    
    let mut total_pv_in = Decimal::ZERO;
    let mut total_pv_out = Decimal::ZERO;
    let mut total_it_pv_in = Decimal::ZERO;
    let mut total_it_pv_out = Decimal::ZERO;

    let mut dynamic_payback_year = 0;
    let mut payback_found = false;

    // Use provided distributions or fallback to 100% Year 1 if not provided/empty
    let default_dist = vec![1.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0];
    let rev_dist = if input.rev_distribution.len() == 10 { &input.rev_distribution } else { &default_dist };
    let cost_dist = if input.cost_distribution.len() == 10 { &input.cost_distribution } else { &default_dist };

    for year in 1..=10 {
        let rev_ratio = Decimal::from_f64_retain(rev_dist[year - 1]).unwrap_or(Decimal::ZERO);
        let cost_ratio = Decimal::from_f64_retain(cost_dist[year - 1]).unwrap_or(Decimal::ZERO);

        let cash_in = (total_rev * rev_ratio).round_dp(2);
        let cash_out = (total_cost * cost_ratio).round_dp(2);
        let net_cash = cash_in - cash_out;

        // IT specific breakdown
        let it_cash_in = (it_rev * rev_ratio).round_dp(2);
        let it_cash_out = (it_cost * cost_ratio).round_dp(2);

        // Discount factor for the current year: (1 + discount_rate)^year
        // In standard NPV formulas, Year 1 cash flow is discounted by (1+r)^1, Year 2 by (1+r)^2, etc.
        let mut pv_factor = Decimal::ONE;
        for _ in 0..year {
            pv_factor *= (Decimal::ONE + discount_rate);
        }

        let pv_in = (cash_in / pv_factor).round_dp(2);
        let pv_out = (cash_out / pv_factor).round_dp(2);
        let pv_net = pv_in - pv_out;

        let it_pv_in = (it_cash_in / pv_factor).round_dp(2);
        let it_pv_out = (it_cash_out / pv_factor).round_dp(2);

        total_pv_in += pv_in;
        total_pv_out += pv_out;
        
        total_it_pv_in += it_pv_in;
        total_it_pv_out += it_pv_out;

        cum_net_cash += net_cash;
        cum_pv += pv_net;

        // Determine dynamic payback period
        if !payback_found && cum_pv >= Decimal::ZERO {
            dynamic_payback_year = year;
            payback_found = true;
        }

        cashflow.push(IctCashflowRow {
            year: year as i32,
            cash_in: cash_in.to_string(),
            cash_out: cash_out.to_string(),
            net_cash: net_cash.to_string(),
            cum_net_cash: cum_net_cash.to_string(),
            pv: pv_net.to_string(),
            cum_pv: cum_pv.to_string(),
        });
    }

    let npv = total_pv_in - total_pv_out;
    let npv_rate = if total_pv_out.is_zero() { Decimal::ZERO } else { (npv / total_pv_out).round_dp(4) };

    let it_npv = total_it_pv_in - total_it_pv_out;
    let it_npv_rate = if total_it_pv_out.is_zero() { Decimal::ZERO } else { (it_npv / total_it_pv_out).round_dp(4) };

    let dynamic_payback_str = if payback_found {
        dynamic_payback_year.to_string()
    } else {
        ">10".to_string()
    };

    Ok(IctResult {
        npv: npv.to_string(),
        npv_rate: npv_rate.to_string(),
        margin_rate: margin_rate.to_string(),
        dynamic_payback: dynamic_payback_str,
        irr: "--".to_string(),
        it_npv: it_npv.to_string(),
        it_npv_rate: it_npv_rate.to_string(),
        it_margin_rate: it_margin_rate.to_string(),
        cashflow,
    })
}

#[tauri::command]
pub fn reverse_calc_ict_target(input: IctInput, target_type: String, target_value: String) -> Result<String, String> {
    let target = Decimal::from_str(&target_value).unwrap_or(Decimal::ZERO);
    let mut low = Decimal::ZERO;
    let mut high = Decimal::new(10_000_000_000, 0); // 10 billion limit
    let mut best_mid = Decimal::ZERO;

    for _ in 0..100 {
        let mid = (low + high) / Decimal::new(2, 0);
        best_mid = mid;

        let mut test_input = input.clone();
        test_input.cost_it_integration.incl_tax = mid.to_string();

        let res = calculate_ict_benefit(test_input)?;
        
        let current_val = if target_type == "margin" {
            Decimal::from_str(&res.margin_rate).unwrap_or(Decimal::ZERO)
        } else {
            Decimal::from_str(&res.npv_rate).unwrap_or(Decimal::ZERO)
        };

        // As cost increases, margin_rate and npv_rate both decrease.
        if current_val > target {
            // We have a higher metric than target -> need to decrease metric -> need to INCREASE cost
            low = mid;
        } else {
            // We have a lower metric than target -> need to increase metric -> need to DECREASE cost
            high = mid;
        }
    }

    Ok(best_mid.round_dp(2).to_string())
}

#[tauri::command]
pub fn reverse_calc_ict_revenue_target(input: IctInput, target_type: String, target_value: String) -> Result<String, String> {
    let target = Decimal::from_str(&target_value).unwrap_or(Decimal::ZERO);
    let mut low = Decimal::ZERO;
    let mut high = Decimal::new(10_000_000_000, 0);
    let mut best_mid = Decimal::ZERO;

    for _ in 0..100 {
        let mid = (low + high) / Decimal::new(2, 0);
        best_mid = mid;

        let mut test_input = input.clone();
        test_input.rev_it_integration.incl_tax = mid.to_string();

        let res = calculate_ict_benefit(test_input)?;

        let current_val = if target_type == "margin" {
            Decimal::from_str(&res.margin_rate).unwrap_or(Decimal::ZERO)
        } else {
            Decimal::from_str(&res.npv_rate).unwrap_or(Decimal::ZERO)
        };

        if current_val < target {
            low = mid;
        } else {
            high = mid;
        }
    }

    Ok(best_mid.round_dp(2).to_string())
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
    let mut rev_col = None;
    let mut exp_col = None;
    
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
                    "收款方式" => rev_col = Some(c_idx),
                    "付款方式" => exp_col = Some(c_idx),
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
    
    let headers = ["项目名称", "含税总收入", "含税总投入", "目标利润率", "目标净现值率", "项目周期", "CT产品名", "CT产品含税总额", "IT税率", "CT税率", "收款方式", "付款方式"];
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
    worksheet.write_string(1, 10, "合同签订后支付XX%").unwrap(); // 收款方式
    worksheet.write_string(1, 11, "背靠背支付").unwrap(); // 付款方式
    
    workbook.save(&path).map_err(|e| format!("写入模板失败: {}", e))?;
    Ok(())
}
