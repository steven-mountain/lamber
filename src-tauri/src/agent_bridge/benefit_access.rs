//! Read-only benefit projections. Subject names come from the desktop catalog.
use super::{
    bridge_server::BridgeReply,
    calculation::{self, CalculateRequest},
    project_bindings::ProjectBindings,
};
use crate::{
    benefit::{
        models::{BenefitAnalysisScheme, BenefitAnalysisSnapshot, Project},
        service::ProjectService,
    },
    workspace::WorkspaceRuntime,
};
use rust_decimal::Decimal;
use serde::Deserialize;
use serde_json::{json, Value};
use std::str::FromStr;

pub const READ_ROUTE: &str = "/lamber-bridge/read-benefit-inputs";
#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct ReadRequest {
    session_id: String,
    scenario: Option<String>,
}
#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct Subject {
    subject_code: String,
    group_id: String,
    side: String,
    standard_subject_name: String,
}

pub fn handle_read(
    runtime: &WorkspaceRuntime,
    bindings: &ProjectBindings,
    body: &str,
) -> BridgeReply {
    let result = (|| {
        let request: ReadRequest = serde_json::from_str(body)
            .map_err(|e| format!("读取参数错误（只接受 scenario，项目由会话绑定确定）: {e}"))?;
        let (workspace, conn) = runtime.require_context()?;
        let cwd = std::fs::canonicalize(&workspace.workspace_root).map_err(|_| "工作区不可用")?;
        let scope = bindings.authorize(
            &request.session_id,
            "read_benefit_inputs",
            None,
            &workspace.workspace_id,
            &cwd,
        )?;
        let service = ProjectService::new(Box::new(
            crate::benefit::repository::SqliteProjectRepository::new(conn),
        ));
        read_inputs(
            &service,
            &CalculateRequest {
                project_id: scope.bound_project_id().ok_or("要求绑定项目")?.into(),
                scenario: request.scenario,
            },
        )
    })();
    match result {
        Ok(value) => BridgeReply::ok(value.to_string()),
        Err(e) => BridgeReply::error(422, &e),
    }
}

pub(super) fn read_inputs(
    service: &ProjectService,
    request: &CalculateRequest,
) -> Result<Value, String> {
    let (project, scheme, snapshot) = calculation::resolve_snapshot(service, request)?;
    project_inputs(&project, &scheme, &snapshot)
}

pub(super) fn project_inputs(
    project: &Project,
    scheme: &BenefitAnalysisScheme,
    snapshot: &BenefitAnalysisSnapshot,
) -> Result<Value, String> {
    let input = serde_json::to_value(&snapshot.input_params).map_err(|e| e.to_string())?;
    let catalog: Vec<Subject> =
        serde_json::from_str(include_str!("../../../src-ui/src/lib/ictSubjects.json"))
            .map_err(|e| e.to_string())?;
    let mut groups: Vec<Value> = Vec::new();
    let mut group_totals: Vec<Decimal> = Vec::new();
    let (mut revenue, mut cost) = (Decimal::ZERO, Decimal::ZERO);
    let mut subjects = Vec::new();
    for subject in catalog {
        let item = &input[&subject.subject_code];
        let raw = item["incl_tax"].as_str().ok_or("快照科目金额缺失")?;
        let amount = Decimal::from_str(raw).map_err(|_| {
            format!(
                "快照科目 {} 金额无效，未静默归零",
                subject.standard_subject_name
            )
        })?;
        let tax = item["tax_rate"].as_str().ok_or("快照科目税率缺失")?;
        Decimal::from_str(tax)
            .map_err(|_| format!("快照科目 {} 税率无效", subject.standard_subject_name))?;
        if subject.side == "revenue" {
            revenue += amount;
        } else {
            cost += amount;
        }
        let group = match groups.iter().position(|g| g["groupId"] == subject.group_id) {
            Some(index) => index,
            None => {
                group_totals.push(Decimal::ZERO);
                groups
                    .push(json!({"groupId":subject.group_id,"side":subject.side,"inclTax":"0.00"}));
                groups.len() - 1
            }
        };
        group_totals[group] += amount;
        let total = group_totals[group];
        groups[group]["inclTax"] = json!(format!("{total:.2}"));
        let custom = item["custom_subject_name"].as_str().unwrap_or("");
        let billing = item["billing_subject_name"].as_str().unwrap_or("");
        let preferred = if billing.trim().is_empty() {
            custom.trim()
        } else {
            billing.trim()
        };
        let display = if preferred.is_empty() {
            subject.standard_subject_name.clone()
        } else {
            format!("{}（{}）", subject.standard_subject_name, preferred)
        };
        let parts = item["split_parts"].as_array().cloned().unwrap_or_default();
        subjects.push(json!({"key":subject.subject_code,"name":subject.standard_subject_name,"displayName":display,
            "groupId":subject.group_id,"side":subject.side,"inclTax":raw,"taxRate":tax,
            "custom_subject_name":custom,"billing_subject_name":billing,"split":!parts.is_empty(),"split_parts":parts}));
    }
    let mut assumptions = serde_json::Map::new();
    for key in [
        "discount_rate",
        "project_years",
        "rev_distribution",
        "cost_distribution",
        "cashflow_model",
        "cashflow_calculation_source",
        "cashflow_segment_value_mode",
        "cashflow_segments",
        "subject_funding_plans",
        "rev_cashflow_excl",
        "cost_cashflow_excl",
        "it_rev_cashflow_excl",
        "it_cost_cashflow_excl",
    ] {
        assumptions.insert(key.into(), input.get(key).cloned().unwrap_or(Value::Null));
    }
    Ok(
        json!({"basis":"saved_snapshot","projectId":project.id,"projectName":project.name,"customerName":project.customer_name,
        "schemeId":scheme.id,"schemeName":scheme.name,"stage":scheme.stage.as_deref().unwrap_or("unlabeled"),"snapshotVersion":snapshot.version,
        "subjects":subjects,"groups":groups,"totals":{"revenueInclTax":format!("{revenue:.2}"),"costInclTax":format!("{cost:.2}")},
        "assumptions":assumptions,"irrNotice":"本系统未计算 IRR；不能按 IRR 筛选或排序，也不能解释为该项目无法求解。"}),
    )
}
