use serde::{Deserialize, Serialize};

fn default_zero_string() -> String {
    "0".to_string()
}

#[derive(Deserialize, Serialize, Clone)]
pub struct IctTaxSplitPart {
    pub incl_tax: String,
    pub excl_tax: String,
}

#[derive(Deserialize, Serialize, Clone)]
pub struct IctItem {
    pub incl_tax: String,
    pub tax_rate: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub custom_subject_name: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub billing_subject_name: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub split_parts: Option<Vec<IctTaxSplitPart>>,
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
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cashflow_calculation_source: Option<String>,
    pub cashflow_segment_value_mode: Option<String>,
    pub cashflow_segments: Option<Vec<CashflowSegment>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub subject_funding_plans: Option<serde_json::Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub subject_funding_plan_migration_version: Option<i32>,
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

    // Optional procurement selection-fee helper inputs. These are persisted for
    // state restoration and document back-filling, but are not part of benefit math.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub selection_fee_quote: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub selection_fee_markup: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub selection_fee_actual_cost: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub selection_fee_amount: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub selection_fee_limit: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub selection_fee_anchor: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub selection_fee_target_subject_code: Option<String>,

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

    // New fields for IT Cashflow display
    #[serde(default = "default_zero_string")]
    pub it_cash_in: String,
    #[serde(default = "default_zero_string")]
    pub it_cash_out: String,
    #[serde(default = "default_zero_string")]
    pub net_it_cash: String,
    #[serde(default = "default_zero_string")]
    pub it_pv: String,
}

#[derive(Serialize, Deserialize, Clone)]
pub struct IctResult {
    pub npv: String,
    pub npv_rate: String,
    pub margin_rate: String,
    pub dynamic_payback: String,
    pub irr: String,

    #[serde(default = "default_zero_string")]
    pub it_npv: String,
    #[serde(default = "default_zero_string")]
    pub it_npv_rate: String,
    #[serde(default = "default_zero_string")]
    pub it_margin_rate: String,

    #[serde(default, alias = "cashflows")]
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
    #[serde(default = "default_project_type")]
    pub project_type: String,
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

fn default_project_type() -> String {
    "ict".to_string()
}

#[derive(Serialize, Deserialize, Clone)]
pub struct BenefitAnalysisScheme {
    pub id: String,
    pub project_id: String,
    pub name: String,
    /// 甄选阶段标签："pre_selection"（甄选前）/ "post_selection"（甄选后）/ None（未标注）。
    #[serde(default)]
    pub stage: Option<String>,
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

#[cfg(test)]
mod tests {
    use super::{IctItem, IctResult};

    #[test]
    fn ict_item_preserves_split_parts_through_snapshot_serialization() {
        let value = serde_json::json!({
            "incl_tax": "826.00",
            "tax_rate": "6",
            "split_parts": [
                { "incl_tax": "413.00", "excl_tax": "389.62" },
                { "incl_tax": "413.00", "excl_tax": "389.62" }
            ]
        });
        let item: IctItem = serde_json::from_value(value).unwrap();
        let serialized = serde_json::to_value(item).unwrap();

        assert_eq!(serialized["split_parts"][0]["incl_tax"], "413.00");
        assert_eq!(serialized["split_parts"][1]["excl_tax"], "389.62");
    }

    #[test]
    fn ict_result_deserializes_legacy_cashflow_without_it_fields() {
        let legacy = serde_json::json!({
            "npv": "100",
            "npv_rate": "0.1",
            "margin_rate": "0.2",
            "dynamic_payback": "1",
            "irr": "--",
            "cashflow": [
                {
                    "year": 1,
                    "cash_in": "1000",
                    "cash_out": "800",
                    "net_cash": "200",
                    "cum_net_cash": "200",
                    "pv": "189.57",
                    "cum_pv": "189.57"
                }
            ]
        });

        let result: IctResult = serde_json::from_value(legacy).unwrap();

        assert_eq!(result.it_npv, "0");
        assert_eq!(result.it_npv_rate, "0");
        assert_eq!(result.it_margin_rate, "0");
        assert_eq!(result.cashflow[0].it_cash_in, "0");
        assert_eq!(result.cashflow[0].it_cash_out, "0");
        assert_eq!(result.cashflow[0].net_it_cash, "0");
        assert_eq!(result.cashflow[0].it_pv, "0");
    }
}
