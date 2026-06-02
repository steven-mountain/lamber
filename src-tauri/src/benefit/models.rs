use serde::{Deserialize, Serialize};

#[derive(Deserialize, Serialize, Clone)]
pub struct IctItem {
    pub incl_tax: String,
    pub tax_rate: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub custom_subject_name: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub billing_subject_name: Option<String>,
}

#[derive(Deserialize, Serialize, Clone, Default)]
pub struct BalanceSubjectRef {
    #[serde(default)]
    pub subject_code: String,
    #[serde(default)]
    pub group_id: String,
    #[serde(default)]
    pub key: String,
}

#[derive(Deserialize, Serialize, Clone, Default)]
pub struct BalanceAllocationRule {
    #[serde(default)]
    pub enabled: bool,
    #[serde(default)]
    pub total_incl_amount: Option<f64>,
    #[serde(default)]
    pub balancing_subject: Option<BalanceSubjectRef>,
}

#[derive(Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct CashflowSegment {
    pub id: String,
    pub name: String,
    pub value: f64,
    pub revenue_value: f64,
    pub revenue_tax: f64,
    pub revenue_scope: String,
    pub cost_value: f64,
    pub cost_tax: f64,
    pub cost_scope: String,
    pub start_year: i32,
    pub service_years: i32,
    pub revenue_mode: String,
    pub cost_mode: String,
    pub revenue_annual_values: Vec<f64>,
    pub cost_annual_values: Vec<f64>,
}

#[derive(Deserialize, Serialize, Clone)]
pub struct IctInput {
    pub project_name: String,
    pub customer_name: Option<String>,
    pub property_rights: String,
    pub discount_rate: String,
    pub project_years: Option<i32>,
    pub cashflow_model: Option<String>,
    pub cashflow_segment_value_mode: Option<String>,
    pub cashflow_segments: Option<Vec<CashflowSegment>>,
    pub project_background: Option<String>,

    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub revenue_balance_rule: Option<BalanceAllocationRule>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub investment_balance_rule: Option<BalanceAllocationRule>,

    // Ignore Tail Difference Payload
    pub ignore_tail_difference: Option<bool>,
    pub tail_difference_value: Option<String>,

    // The revenue and cost distributions over 10 years (e.g., [1.0, 0.0, ..., 0.0])
    pub rev_distribution: Vec<f64>,
    pub cost_distribution: Vec<f64>,

    // Optional yearly tax-exclusive cashflow overrides. Model E amount mode uses these
    // to preserve segment-specific tax rates and custom payment schedules.
    pub rev_cashflow_excl: Option<Vec<String>>,
    pub cost_cashflow_excl: Option<Vec<String>>,
    pub it_rev_cashflow_excl: Option<Vec<String>>,
    pub it_cost_cashflow_excl: Option<Vec<String>>,

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

#[derive(Serialize, Deserialize, Clone)]
pub struct IctCashflowRow {
    pub year: i32,
    pub cash_in: String,
    pub cash_out: String,
    pub net_cash: String,
    pub cum_net_cash: String,
    pub pv: String,
    pub cum_pv: String,
}

#[derive(Serialize, Deserialize, Clone)]
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

#[derive(Serialize, Deserialize, Clone)]
pub struct SelectionFeeResult {
    pub selection_fee: String,
    pub actual_cost: String,
    pub final_limit: String,
    pub quote: String,
}

#[derive(Serialize, Deserialize, Clone)]
pub struct ProjectLog {
    pub id: String,
    pub timestamp: String,
    pub description: String,
}

#[derive(Serialize, Deserialize, Clone)]
pub struct SummaryMetrics {
    pub margin_rate: String,
    pub npv: String,
    pub npv_rate: String,
    pub irr: String,
    pub dynamic_payback: String,
    pub risk_level: String,
}

#[derive(Serialize, Deserialize, Clone)]
pub struct Project {
    pub id: String,
    pub name: String,
    pub customer_name: String,
    pub status: String, // User-editable lifecycle tag, defaults to "需求导入"
    pub benefit_status: String, // "not_started", "normal", "outdated"
    pub default_scheme_id: Option<String>,
    pub created_at: String,
    pub updated_at: String,

    // Fingerprint parameters
    pub total_revenue_incl: f64,
    pub total_cost_incl: f64,
    pub project_years: i32,
    pub discount_rate: f64,
    pub cashflow_model: String,

    pub summary_metrics: Option<SummaryMetrics>,

    // Folders binding fields
    pub folder_path: Option<String>,
    pub main_document_path: Option<String>,
    pub main_budget_file_path: Option<String>,

    #[serde(default)]
    pub note: Option<String>,

    pub logs: Vec<ProjectLog>,

    #[serde(default)]
    pub folder_name: Option<String>,
    #[serde(default)]
    pub relative_path: Option<String>,
    #[serde(default)]
    pub progress: f64,
    #[serde(default)]
    pub deadline: Option<String>,
    #[serde(default)]
    pub linked_folder_type: Option<String>,
    #[serde(default)]
    pub linked_folder_relative_path: Option<String>,
    #[serde(default)]
    pub linked_folder_external_path: Option<String>,
}

#[derive(Serialize, Deserialize, Clone)]
pub struct BenefitAnalysisScheme {
    pub id: String,
    pub project_id: String,
    pub name: String,
    pub created_at: String,
    pub updated_at: String,
}

#[derive(Serialize, Deserialize, Clone)]
pub struct BenefitAnalysisSnapshot {
    pub id: String,
    pub scheme_id: String,
    pub project_id: String,
    pub version: i32,
    pub input_params: IctInput,
    pub output_metrics: IctResult,
    pub fingerprint: String,
    pub created_at: String,
}

use crate::project_files::models::ProjectFile;

// StoreData is a shared storage database model across the project, benefit analysis, and project_files modules.
// In the future, this should be migrated to a dedicated shared storage module.
#[derive(Serialize, Deserialize, Clone)]
pub struct StoreData {
    pub schema_version: i32,
    pub projects: Vec<Project>,
    pub schemes: Vec<BenefitAnalysisScheme>,
    pub snapshots: Vec<BenefitAnalysisSnapshot>,
    #[serde(default)]
    pub project_files: Vec<ProjectFile>,
}
