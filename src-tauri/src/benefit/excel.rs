use super::models::{IctInput, IctItem};
use calamine::{open_workbook, Reader, Xlsx};
use std::collections::HashMap;
use std::path::Path;

#[derive(serde::Serialize, Clone)]
pub struct ExcelParsedTaxItem {
    pub incl_tax: f64,
    pub excl_tax: f64,
    pub tax_rate: f64,
    pub custom_subject_name: Option<String>,
    pub billing_subject_name: Option<String>,
    pub standard_subject_name: Option<String>,
    pub display_name: Option<String>,
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
                .replace("（", "")
                .replace("）", "")
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

fn normalize_subject_text(value: &str) -> String {
    value.split_whitespace().collect::<String>()
}

fn extract_custom_subject_name(
    display_name: &str,
    standard_subject_name: &str,
    template_subject_name: &str,
) -> Option<String> {
    let display_trimmed = display_name.trim().replace(['\n', '\r'], "");
    let standard_trimmed = standard_subject_name.trim();
    let display = normalize_subject_text(display_name);
    let standard = normalize_subject_text(standard_subject_name);
    let template = normalize_subject_text(template_subject_name);

    if display.is_empty() || display == standard || display == template {
        return None;
    }

    let prefix = format!("{}（", standard_trimmed);
    if display_trimmed.starts_with(&prefix) && display_trimmed.ends_with('）') {
        let custom = display_trimmed[prefix.len()..display_trimmed.len() - '）'.len_utf8()]
            .trim()
            .to_string();
        if custom.is_empty() || custom.contains("具体测算规则") {
            return None;
        }
        return Some(custom);
    }

    None
}

fn item_from_cells(
    range: &calamine::Range<calamine::Data>,
    name_cell: &str,
    standard_subject_name: &str,
    template_subject_name: &str,
    incl_cell: &str,
    excl_cell: &str,
    tax_rate: f64,
) -> ExcelParsedTaxItem {
    let tax = normalize_tax_percent(tax_rate);
    let incl = get_cell_f64(range, incl_cell);
    let excl = get_cell_f64(range, excl_cell);
    let display_name = get_cell_string(range, name_cell);
    let custom_subject_name =
        extract_custom_subject_name(&display_name, standard_subject_name, template_subject_name);

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
        custom_subject_name,
        billing_subject_name: None,
        standard_subject_name: Some(standard_subject_name.to_string()),
        display_name: if display_name.trim().is_empty() {
            None
        } else {
            Some(display_name)
        },
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
        (
            "rev_it_integration",
            "D3",
            "系统集成服务收入",
            "系统集成服务收入",
            "Q3",
            "G3",
            6.0,
        ),
        (
            "rev_it_maintenance",
            "D4",
            "维保收入",
            "维保收入",
            "Q4",
            "G4",
            6.0,
        ),
        (
            "rev_it_device_sales",
            "D5",
            "设备销售收入",
            "设备销售收入",
            "Q5",
            "G5",
            13.0,
        ),
        (
            "rev_it_device_lease",
            "D6",
            "设备租赁收入",
            "设备租赁收入",
            "Q6",
            "G6",
            13.0,
        ),
        (
            "rev_it_other",
            "D7",
            "其他收入",
            "其他收入（代销设备、代理采购手续费等）",
            "Q7",
            "G7",
            6.0,
        ),
        (
            "rev_it_cloud",
            "D8",
            "移动云-定制化收入",
            "移动云-定制化收入",
            "Q8",
            "G8",
            6.0,
        ),
        ("rev_ct_line", "D9", "专线收入", "专线收入", "Q9", "G9", 9.0),
        (
            "rev_ct_product",
            "D10",
            "产品收入",
            "产品收入",
            "Q10",
            "G10",
            6.0,
        ),
        (
            "rev_non_it_ct",
            "D11",
            "工程施工收入等",
            "工程施工收入等",
            "Q11",
            "G11",
            9.0,
        ),
        (
            "cost_it_device",
            "E13",
            "主要设备/甲供材料",
            "主要设备/甲供材料",
            "Q13",
            "G13",
            13.0,
        ),
        (
            "cost_it_construction",
            "E14",
            "施工",
            "施工",
            "Q14",
            "G14",
            9.0,
        ),
        (
            "cost_it_survey",
            "E15",
            "勘察设计/预备费",
            "勘察设计/预备费",
            "Q15",
            "G15",
            6.0,
        ),
        (
            "cost_it_integration",
            "E16",
            "集成服务",
            "集成服务",
            "Q16",
            "G16",
            6.0,
        ),
        (
            "cost_it_other",
            "E17",
            "其他投入",
            "其他投入",
            "Q17",
            "G17",
            6.0,
        ),
        (
            "cost_it_maintenance",
            "E18",
            "维护费用",
            "维护费用",
            "Q18",
            "G18",
            6.0,
        ),
        (
            "cost_it_running",
            "E19",
            "其他运行支出（电费等）",
            "其他运行支出（电费等）",
            "Q19",
            "G19",
            13.0,
        ),
        (
            "cost_it_bidding",
            "E20",
            "中标服务费",
            "中标服务费",
            "Q20",
            "G20",
            6.0,
        ),
        (
            "cost_it_design_eval",
            "E21",
            "设计院成本评估费",
            "设计院成本评估费",
            "Q21",
            "G21",
            6.0,
        ),
        (
            "cost_it_audit",
            "E22",
            "第三方审计评估费",
            "第三方审计评估费",
            "Q22",
            "G22",
            6.0,
        ),
        (
            "cost_ct_construction",
            "E23",
            "专线建设",
            "专线建设",
            "Q23",
            "G23",
            6.0,
        ),
        (
            "cost_ct_maintenance",
            "E24",
            "专线维护",
            "专线维护",
            "Q24",
            "G24",
            9.0,
        ),
        (
            "cost_ct_other",
            "E25",
            "其他产品成本",
            "其他产品成本（具体测算规则详见Sheet5）",
            "Q25",
            "G25",
            6.0,
        ),
        (
            "cost_ct_bandwidth",
            "E26",
            "专线带宽成本",
            "专线带宽成本",
            "Q26",
            "G26",
            9.0,
        ),
        (
            "cost_ct_renewal",
            "E27",
            "专线/其他产品续签成本",
            "专线/其他产品续签成本",
            "Q27",
            "G27",
            9.0,
        ),
        (
            "cost_non_it_ct",
            "E28",
            "工程施工投入等",
            "工程施工投入等",
            "Q28",
            "G28",
            9.0,
        ),
        (
            "cost_mix_marketing",
            "D29",
            "融合营销成本",
            "融合营销成本",
            "Q29",
            "G29",
            6.0,
        ),
        (
            "cost_mix_channel",
            "D30",
            "渠道酬金",
            "渠道酬金",
            "Q30",
            "G30",
            6.0,
        ),
        (
            "cost_mix_other",
            "D31",
            "其他管理费用等",
            "其他管理费用等",
            "Q31",
            "G31",
            6.0,
        ),
    ];

    for (
        key,
        name_cell,
        standard_subject_name,
        template_subject_name,
        incl_cell,
        excl_cell,
        tax_rate,
    ) in mappings
    {
        items.insert(
            key.to_string(),
            item_from_cells(
                &econ_range,
                name_cell,
                standard_subject_name,
                template_subject_name,
                incl_cell,
                excl_cell,
                tax_rate,
            ),
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
pub async fn parse_benefit_excel(
    runtime: tauri::State<'_, std::sync::Arc<crate::workspace::WorkspaceRuntime>>,
    file_path: String,
) -> Result<ExcelParsedData, String> {
    let path_buf = std::path::PathBuf::from(&file_path);
    let resolved_path = if !path_buf.is_absolute() {
        if let Some(ws) = runtime.get_current_workspace() {
            let ws_path = Path::new(&ws.workspace_root);
            crate::workspace::resolve_workspace_path(ws_path, &file_path)
        } else {
            path_buf
        }
    } else {
        path_buf
    };

    tauri::async_runtime::spawn_blocking(move || {
        if !resolved_path.exists() {
            return Err(format!("文件不存在: {}", file_path));
        }
        parse_benefit_excel_internal(&resolved_path)
    })
    .await
    .map_err(|e| format!("异步执行异常: {}", e))?
}

pub fn parse_benefit_excel_internal(resolved_path: &Path) -> Result<ExcelParsedData, String> {
    let resolved_path_str = resolved_path.to_string_lossy().to_string();

    let mut workbook: Xlsx<_> =
        open_workbook(&resolved_path_str).map_err(|e| format!("打开 Excel 异常: {}", e))?;
    if let Some(parsed) = parse_lifecycle_benefit_excel(&mut workbook)? {
        return Ok(parsed);
    }

    let sheet_names = workbook.sheet_names().to_owned();
    let sheet_name = sheet_names.first().ok_or("找不到工作表")?.clone();
    let range = workbook
        .worksheet_range(&sheet_name)
        .map_err(|e| format!("读取工作表异常: {}", e))?;

    let mut rows = range.rows();
    let headers_row = rows.next().ok_or("Excel 表为空")?;

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
            "CT产品名称" | "CT产品" => ct_name_col = Some(c_idx),
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
}

pub fn auto_import_excel_calculation(
    project_id: &str,
    _file_name: &str,
    parsed_data: ExcelParsedData,
    service: &crate::benefit::service::ProjectService,
) -> Result<(), String> {
    let ct_income = parsed_data.ct_income_incl;
    let it_income = (parsed_data.total_income_incl - ct_income).max(0.0);

    let to_tax_percent = |value: f64, fallback: f64| -> f64 {
        if !value.is_finite() || value <= 0.0 {
            fallback
        } else if value > 0.0 && value < 1.0 {
            value * 100.0
        } else {
            value
        }
    };

    let it_tax = to_tax_percent(parsed_data.it_tax, 6.0);
    let ct_tax = to_tax_percent(parsed_data.ct_tax, 6.0);

    let has_detailed_items = parsed_data
        .items
        .values()
        .any(|item| item.incl_tax.abs() > 0.0 || item.excl_tax.abs() > 0.0);

    let make_item = |incl: f64,
                     tax: f64,
                     custom_subject_name: Option<String>,
                     billing_subject_name: Option<String>|
     -> IctItem {
        IctItem {
            incl_tax: format!("{:.2}", if incl.is_finite() { incl } else { 0.0 }),
            tax_rate: format!("{:.4}", if tax.is_finite() { tax } else { 0.0 }),
            custom_subject_name: custom_subject_name
                .map(|value| value.trim().to_string())
                .filter(|value| !value.is_empty()),
            billing_subject_name: billing_subject_name
                .map(|value| value.trim().to_string())
                .filter(|value| !value.is_empty()),
        }
    };

    let make_parsed_item = |key: &str, default_tax: f64, fallback_incl: f64| -> IctItem {
        if let Some(item) = parsed_data.items.get(key) {
            let tax = to_tax_percent(item.tax_rate, default_tax);
            if item.incl_tax.is_finite() && item.incl_tax.abs() > 0.0 {
                return make_item(
                    item.incl_tax,
                    tax,
                    item.custom_subject_name.clone(),
                    item.billing_subject_name.clone(),
                );
            }
            if item.excl_tax.is_finite() && item.excl_tax.abs() > 0.0 {
                return make_item(
                    item.excl_tax * (1.0 + tax / 100.0),
                    tax,
                    item.custom_subject_name.clone(),
                    item.billing_subject_name.clone(),
                );
            }
            return make_item(
                0.0,
                tax,
                item.custom_subject_name.clone(),
                item.billing_subject_name.clone(),
            );
        }
        make_item(
            if has_detailed_items {
                0.0
            } else {
                fallback_incl
            },
            default_tax,
            None,
            None,
        )
    };

    let payload = IctInput {
        project_name: if parsed_data.project_name.is_empty() {
            "未命名项目".to_string()
        } else {
            parsed_data.project_name.clone()
        },
        customer_name: Some("CMCC".to_string()),
        property_rights: "客户".to_string(),
        discount_rate: format!(
            "{:.4}",
            if parsed_data.discount_rate > 0.0 {
                parsed_data.discount_rate
            } else {
                0.055
            }
        ),
        project_years: Some(if parsed_data.project_years > 0 {
            parsed_data.project_years
        } else {
            1
        }),
        cashflow_model: Some("model_a".to_string()),
        cashflow_calculation_source: Some("legacy_model".to_string()),
        cashflow_segment_value_mode: Some("ratio".to_string()),
        cashflow_segments: Some(vec![]),
        project_background: None,
        revenue_balance_rule: None,
        investment_balance_rule: None,
        ignore_tail_difference: Some(false),
        tail_difference_value: Some("0".to_string()),
        rev_distribution: vec![1.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
        cost_distribution: vec![1.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
        rev_cashflow_excl: None,
        cost_cashflow_excl: None,
        it_rev_cashflow_excl: None,
        it_cost_cashflow_excl: None,

        rev_it_integration: make_parsed_item("rev_it_integration", it_tax, it_income),
        rev_it_maintenance: make_parsed_item("rev_it_maintenance", 6.0, 0.0),
        rev_it_device_sales: make_parsed_item("rev_it_device_sales", 13.0, 0.0),
        rev_it_device_lease: make_parsed_item("rev_it_device_lease", 13.0, 0.0),
        rev_it_other: make_parsed_item("rev_it_other", 6.0, 0.0),
        rev_it_cloud: make_parsed_item("rev_it_cloud", 6.0, 0.0),

        rev_ct_line: make_parsed_item("rev_ct_line", 9.0, 0.0),
        rev_ct_product: make_parsed_item("rev_ct_product", ct_tax, ct_income),

        rev_non_it_ct: make_parsed_item("rev_non_it_ct", 9.0, 0.0),

        cost_it_device: make_parsed_item("cost_it_device", 13.0, parsed_data.total_cost_incl),
        cost_it_construction: make_parsed_item("cost_it_construction", 9.0, 0.0),
        cost_it_survey: make_parsed_item("cost_it_survey", 6.0, 0.0),
        cost_it_integration: make_parsed_item("cost_it_integration", 6.0, 0.0),
        cost_it_other: make_parsed_item("cost_it_other", 6.0, 0.0),
        cost_it_maintenance: make_parsed_item("cost_it_maintenance", 6.0, 0.0),
        cost_it_running: make_parsed_item("cost_it_running", 13.0, 0.0),
        cost_it_bidding: make_parsed_item("cost_it_bidding", 6.0, 0.0),
        cost_it_design_eval: make_parsed_item("cost_it_design_eval", 6.0, 0.0),
        cost_it_audit: make_parsed_item("cost_it_audit", 6.0, 0.0),

        cost_ct_construction: make_parsed_item("cost_ct_construction", 6.0, 0.0),
        cost_ct_maintenance: make_parsed_item("cost_ct_maintenance", 9.0, 0.0),
        cost_ct_other: make_parsed_item("cost_ct_other", ct_tax, 0.0),
        cost_ct_bandwidth: make_parsed_item("cost_ct_bandwidth", 9.0, 0.0),
        cost_ct_renewal: make_parsed_item("cost_ct_renewal", 9.0, 0.0),

        cost_non_it_ct: make_parsed_item("cost_non_it_ct", 9.0, 0.0),
        cost_mix_marketing: make_parsed_item("cost_mix_marketing", 6.0, 0.0),
        cost_mix_channel: make_parsed_item("cost_mix_channel", 6.0, 0.0),
        cost_mix_other: make_parsed_item("cost_mix_other", 6.0, 0.0),
    };

    let result = crate::benefit::calculator::calculate_ict_benefit(payload.clone())?;
    let _ = service.save_benefit_scheme(
        project_id.to_string(),
        None,
        "Excel导入测算方案".to_string(),
        payload,
        result,
        false,
    )?;

    Ok(())
}
