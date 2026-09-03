//! `POST /lamber-bridge/calculate` — the `run_benefit_calculation` tool's backend.
//!
//! The agent never receives raw project state and never performs math. It names
//! a project (and optionally a scheme), and lamber replays that scheme's latest
//! saved input snapshot through `benefit::calculator::calculate_ict_benefit`,
//! the same engine the desktop 测算表 uses. That keeps the numbers the agent
//! quotes identical to the numbers the user sees, and keeps this route strictly
//! read-only: nothing here writes to the workspace database.

use crate::benefit::models::{BenefitAnalysisScheme, BenefitAnalysisSnapshot, IctResult, Project};
use crate::benefit::service::ProjectService;
use serde::{Deserialize, Serialize};

/// Route path served for benefit calculations.
pub const CALCULATE_ROUTE: &str = "/lamber-bridge/calculate";

/// Stage selector the agent may pass as `scenario` (matches `lib/schemeStage.ts`).
const STAGE_PRE_SELECTION: &str = "pre_selection";
const STAGE_POST_SELECTION: &str = "post_selection";
/// Reported stage for a scheme that carries no 甄选 label.
const STAGE_UNLABELED: &str = "unlabeled";

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CalculateRequest {
    pub project_id: String,
    /// `pre_selection` / `post_selection` / a scheme id / a scheme name. `None` uses the default scheme.
    #[serde(default)]
    pub scenario: Option<String>,
}

#[derive(Serialize, Debug)]
#[serde(rename_all = "camelCase")]
pub struct CalculateMetrics {
    pub npv: String,
    pub npv_rate: String,
    pub margin_rate: String,
    pub dynamic_payback: String,
    pub irr: String,
    pub it_npv: String,
    pub it_npv_rate: String,
    pub it_margin_rate: String,
}

#[derive(Serialize, Debug)]
#[serde(rename_all = "camelCase")]
pub struct CalculateCashflowRow {
    pub year: i32,
    pub cash_in: String,
    pub cash_out: String,
    pub net_cash: String,
    pub cum_net_cash: String,
    pub pv: String,
    pub cum_pv: String,
}

#[derive(Serialize, Debug)]
#[serde(rename_all = "camelCase")]
pub struct CalculateResponse {
    pub project_id: String,
    pub project_name: String,
    pub customer_name: String,
    pub scheme_id: String,
    pub scheme_name: String,
    /// Always a string; an unlabeled scheme reports `unlabeled` rather than null.
    pub stage: String,
    pub snapshot_version: i32,
    pub calculated_at: String,
    pub metrics: CalculateMetrics,
    pub cashflow: Vec<CalculateCashflowRow>,
}

/// Resolve a project + scheme and re-run its saved inputs through the ICT engine.
///
/// @param service - project service bound to the open workspace database.
/// @param request - project id and optional scheme selector.
/// @returns the freshly computed metrics, tagged with what was actually used.
pub fn run_calculation(
    service: &ProjectService,
    request: &CalculateRequest,
) -> Result<CalculateResponse, String> {
    let project = service
        .get_project(&request.project_id)?
        .ok_or_else(|| format!("未找到项目 {}", request.project_id))?;

    let schemes = service.get_schemes(&project.id)?;
    if schemes.is_empty() {
        return Err(format!("项目「{}」还没有任何测算方案", project.name));
    }
    let scheme = select_scheme(&project, &schemes, request.scenario.as_deref())?;

    let snapshot = latest_snapshot(service, &scheme.id)?;
    let result = crate::benefit::calculator::calculate_ict_benefit(snapshot.input_params.clone())?;

    Ok(build_response(&project, scheme, &snapshot, &result))
}

/// Pick the scheme the `scenario` selector names, falling back to the project default.
///
/// Accepted selectors, in priority order: a 甄选 stage tag, an exact scheme id,
/// then an exact scheme name. An unmatched selector is an error rather than a
/// silent fallback — quoting the wrong scheme's financials would be worse than
/// telling the agent its selector was wrong.
fn select_scheme<'a>(
    project: &Project,
    schemes: &'a [BenefitAnalysisScheme],
    scenario: Option<&str>,
) -> Result<&'a BenefitAnalysisScheme, String> {
    let Some(selector) = scenario.map(str::trim).filter(|s| !s.is_empty()) else {
        return default_scheme(project, schemes);
    };

    if selector == STAGE_PRE_SELECTION || selector == STAGE_POST_SELECTION {
        return newest(schemes.iter().filter(|s| s.stage.as_deref() == Some(selector)))
            .ok_or_else(|| {
                format!(
                    "项目「{}」没有标记为 {} 的测算方案",
                    project.name, selector
                )
            });
    }

    if let Some(found) = schemes.iter().find(|s| s.id == selector) {
        return Ok(found);
    }
    if let Some(found) = newest(schemes.iter().filter(|s| s.name == selector)) {
        return Ok(found);
    }

    Err(format!(
        "项目「{}」中没有匹配 `{}` 的测算方案；可选方案：{}",
        project.name,
        selector,
        schemes
            .iter()
            .map(|s| s.name.as_str())
            .collect::<Vec<_>>()
            .join("、")
    ))
}

fn default_scheme<'a>(
    project: &Project,
    schemes: &'a [BenefitAnalysisScheme],
) -> Result<&'a BenefitAnalysisScheme, String> {
    if let Some(default_id) = project.default_scheme_id.as_deref() {
        if let Some(found) = schemes.iter().find(|s| s.id == default_id) {
            return Ok(found);
        }
    }
    newest(schemes.iter())
        .ok_or_else(|| format!("项目「{}」还没有任何测算方案", project.name))
}

/// Most recently updated scheme of an iterator, matching the UI's ordering rule.
fn newest<'a>(
    schemes: impl Iterator<Item = &'a BenefitAnalysisScheme>,
) -> Option<&'a BenefitAnalysisScheme> {
    schemes.max_by(|a, b| {
        a.updated_at
            .cmp(&b.updated_at)
            .then_with(|| a.created_at.cmp(&b.created_at))
    })
}

/// Highest-version snapshot of a scheme; that is the input set the UI last saved.
fn latest_snapshot(
    service: &ProjectService,
    scheme_id: &str,
) -> Result<BenefitAnalysisSnapshot, String> {
    service
        .get_snapshots(scheme_id)?
        .into_iter()
        .max_by_key(|s| s.version)
        .ok_or_else(|| format!("测算方案 {scheme_id} 还没有保存过任何测算快照"))
}

fn build_response(
    project: &Project,
    scheme: &BenefitAnalysisScheme,
    snapshot: &BenefitAnalysisSnapshot,
    result: &IctResult,
) -> CalculateResponse {
    CalculateResponse {
        project_id: project.id.clone(),
        project_name: project.name.clone(),
        customer_name: project.customer_name.clone(),
        scheme_id: scheme.id.clone(),
        scheme_name: scheme.name.clone(),
        stage: scheme
            .stage
            .clone()
            .unwrap_or_else(|| STAGE_UNLABELED.to_string()),
        snapshot_version: snapshot.version,
        calculated_at: chrono::Utc::now().to_rfc3339(),
        metrics: CalculateMetrics {
            npv: result.npv.clone(),
            npv_rate: result.npv_rate.clone(),
            margin_rate: result.margin_rate.clone(),
            dynamic_payback: result.dynamic_payback.clone(),
            irr: result.irr.clone(),
            it_npv: result.it_npv.clone(),
            it_npv_rate: result.it_npv_rate.clone(),
            it_margin_rate: result.it_margin_rate.clone(),
        },
        cashflow: result
            .cashflow
            .iter()
            .map(|row| CalculateCashflowRow {
                year: row.year,
                cash_in: row.cash_in.clone(),
                cash_out: row.cash_out.clone(),
                net_cash: row.net_cash.clone(),
                cum_net_cash: row.cum_net_cash.clone(),
                pv: row.pv.clone(),
                cum_pv: row.cum_pv.clone(),
            })
            .collect(),
    }
}
