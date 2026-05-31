use serde::{Deserialize, Serialize};
use serde_json::Value;

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct AiProjectContextRequest {
    pub project_id: String,
    pub requested_sources: Option<Vec<String>>,
    pub active_template_id: Option<String>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct AiProjectContextBundle {
    pub project_id: String,
    pub project_name: String,
    pub overview: AiProjectOverview,
    pub lifecycle: Option<AiLifecycleContext>,
    pub cashflow: Option<AiCashflowContext>,
    pub benefit: Option<AiBenefitContext>,
    pub templates: Option<Vec<AiTemplateContextSummary>>,
    pub template_detail: Option<AiTemplateDetailContext>,
    pub files: Option<AiFileContextSummary>,
    pub sources: Vec<AiContextSourceMeta>,
    pub warnings: Vec<String>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct AiWorkspaceProjectIndexItem {
    pub project_id: String,
    pub project_name: String,
    pub customer_name: Option<String>,
    pub status: Option<String>,
    pub updated_at: Option<String>,
    pub has_lifecycle_state: bool,
    pub has_cashflow_state: bool,
    pub has_template_state: bool,
    pub template_names: Vec<String>,
    pub has_benefit_schemes: bool,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct AiProjectOverview {
    pub name: String,
    pub customer_name: Option<String>,
    pub status: Option<String>,
    pub phase: Option<String>,
    pub deadline: Option<String>,
    pub description: Option<String>,
    pub progress: Option<f64>,
    pub benefit_status: Option<String>,
    pub folder_linked: bool,
    pub updated_at: Option<String>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct AiLifecycleContext {
    pub has_saved_state: bool,
    pub lifecycle_version: Option<i64>,
    pub updated_at: Option<String>,
    pub summary_json: Value,
    pub profile_json: Option<Value>,
    pub parameters_json: Option<Value>,
    pub background_json: Option<Value>,
    pub input_payload_json: Option<Value>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct AiCashflowContext {
    pub has_saved_state: bool,
    pub cashflow_version: Option<i64>,
    pub cashflow_model: Option<String>,
    pub has_yearly_cashflow: bool,
    pub year_count: Option<usize>,
    pub updated_at: Option<String>,
    pub summary_json: Value,
    pub payment_model_json: Option<Value>,
    pub yearly_cashflow_json: Option<Value>,
    pub sector_cashflow_json: Option<Value>,
    pub assumptions_json: Option<Value>,
    pub metrics_json: Option<Value>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct AiBenefitContext {
    pub scheme_count: usize,
    pub default_scheme: Option<AiBenefitSchemeSummary>,
    pub latest_scheme: Option<AiBenefitSchemeSummary>,
    pub latest_snapshot: Option<AiBenefitSnapshotSummary>,
    pub project_summary_metrics: Option<Value>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct AiBenefitSchemeSummary {
    pub id: String,
    pub name: String,
    pub updated_at: Option<String>,
    pub is_default: bool,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct AiBenefitSnapshotSummary {
    pub id: String,
    pub scheme_id: String,
    pub version: i64,
    pub created_at: Option<String>,
    pub output_metrics_summary: Option<Value>,
    pub input_params: Option<Value>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct AiTemplateContextSummary {
    pub template_id: String,
    pub template_name: Option<String>,
    pub has_saved_state: bool,
    pub updated_at: Option<String>,
    pub field_count: Option<usize>,
    pub asset_count: Option<usize>,
    pub source: String,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct AiTemplateDetailContext {
    pub project_id: String,
    pub template_id: String,
    pub template_name: Option<String>,
    pub source: String,
    pub has_saved_state: bool,
    pub updated_at: Option<String>,
    pub fields: Value,
    pub field_mapping: Option<Value>,
    pub output_config: Option<Value>,
    pub assets: Vec<AiTemplateAssetReference>,
    pub warnings: Vec<String>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct AiTemplateAssetReference {
    pub asset_id: String,
    pub field_key: Option<String>,
    pub file_name: Option<String>,
    pub mime_type: Option<String>,
    pub file_size: i64,
    pub width: Option<i32>,
    pub height: Option<i32>,
    pub exists: Option<bool>,
    pub updated_at: Option<String>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct AiTemplateAssetRequest {
    pub project_id: String,
    pub asset_id: String,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct AiTemplateAssetImageInput {
    pub id: String,
    pub project_id: String,
    pub name: String,
    pub mime_type: String,
    pub size: i64,
    pub width: Option<i32>,
    pub height: Option<i32>,
    pub data_url: String,
    pub source: String,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct AiFileContextSummary {
    pub total_files: usize,
    pub existing_files: usize,
    pub missing_files: usize,
    pub file_type_counts: Vec<AiNamedCount>,
    pub storage_mode_counts: Vec<AiNamedCount>,
    pub main_document_count: usize,
    pub main_budget_file_count: usize,
    pub files: Option<Vec<AiProjectFileSummary>>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct AiProjectFileSummary {
    pub id: String,
    pub file_name: String,
    pub file_type: String,
    pub extension: String,
    pub size: i64,
    pub exists: bool,
    pub storage_mode: String,
    pub is_main_document: bool,
    pub is_main_budget_file: bool,
    pub file_role: Option<String>,
    pub modified_at: Option<String>,
    pub updated_at: Option<String>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct AiNamedCount {
    pub name: String,
    pub count: usize,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct AiContextSourceMeta {
    pub source_type: String,
    pub source_id: Option<String>,
    pub updated_at: Option<String>,
}
