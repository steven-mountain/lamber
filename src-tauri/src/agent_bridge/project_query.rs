//! Bounded projection over the existing projects service. No SQL or detail-table access.
use crate::benefit::{
    models::{BenefitAnalysisScheme, Project},
    service::ProjectService,
};
use chrono::{DateTime, NaiveDate, NaiveDateTime, Utc};
use rust_decimal::Decimal;
use serde::{Deserialize, Serialize};
use std::str::FromStr;

pub const QUERY_ROUTE: &str = "/lamber-bridge/query-projects";
pub const QUERY_TOOL: &str = "query_projects";
pub const MAX_LIMIT: usize = 50;
pub const NOTICE: &str = "以下为跨项目检索结果，不得用作当前绑定项目的测算结论。isBoundProject 标记命中的绑定项目；金额为含税元，比例使用小数（0.2=20%）。指标是已保存的项目摘要，非本次重新测算；缺失指标不等于零。指标属于 defaultSchemeId 对应的最后保存方案；必须同时说明 stageLabel、方案名和保存时间，未标注不得猜测。合计按甄选阶段分组覆盖全部命中，混合口径不提供跨阶段合计。";

#[derive(Default, Deserialize, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct Range {
    pub gte: Option<f64>,
    pub gt: Option<f64>,
    pub lte: Option<f64>,
    pub lt: Option<f64>,
}
impl Range {
    fn validate(&self) -> Result<(), String> {
        if [self.gte, self.gt, self.lte, self.lt]
            .iter()
            .all(Option::is_none)
        {
            return Err("数值区间至少需要一个边界".into());
        }
        if [self.gte, self.gt, self.lte, self.lt]
            .into_iter()
            .flatten()
            .any(|v| !v.is_finite())
        {
            return Err("数值区间必须是有限数字".into());
        }
        let lows = [(self.gte, false), (self.gt, true)];
        let highs = [(self.lte, false), (self.lt, true)];
        for (low, exclusive_low) in lows {
            for (high, exclusive_high) in highs {
                if let (Some(low), Some(high)) = (low, high) {
                    if low > high || (low == high && (exclusive_low || exclusive_high)) {
                        return Err("数值区间为空或上下限颠倒".into());
                    }
                }
            }
        }
        Ok(())
    }
    fn matches(&self, value: Option<f64>) -> bool {
        value.filter(|v| v.is_finite()).is_some_and(|v| {
            self.gte.is_none_or(|x| v >= x)
                && self.gt.is_none_or(|x| v > x)
                && self.lte.is_none_or(|x| v <= x)
                && self.lt.is_none_or(|x| v < x)
        })
    }
}
#[derive(Default, Deserialize, Serialize, Clone, Copy)]
#[serde(rename_all = "camelCase")]
pub enum SortField {
    Name,
    CreatedAt,
    #[default]
    UpdatedAt,
    TotalRevenueIncl,
    TotalCostIncl,
    MarginRate,
    Npv,
    NpvRate,
    Irr,
    DynamicPayback,
}
#[derive(Default, Deserialize, Serialize, Clone, Copy)]
#[serde(rename_all = "camelCase")]
pub enum SortOrder {
    Asc,
    #[default]
    Desc,
}
#[derive(Default, Deserialize, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct QueryRequest {
    pub customer_name: Option<String>,
    pub status: Option<String>,
    pub benefit_status: Option<String>,
    pub created_from: Option<String>,
    pub created_before: Option<String>,
    pub updated_from: Option<String>,
    pub updated_before: Option<String>,
    pub total_revenue_incl: Option<Range>,
    pub total_cost_incl: Option<Range>,
    pub margin_rate: Option<Range>,
    pub npv: Option<Range>,
    pub npv_rate: Option<Range>,
    pub irr: Option<Range>,
    pub dynamic_payback: Option<Range>,
    #[serde(default)]
    pub sort_by: SortField,
    #[serde(default)]
    pub sort_order: SortOrder,
    pub limit: Option<usize>,
}
#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct QueryEnvelope {
    pub session_id: String,
    pub query: QueryRequest,
}
#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Metrics {
    margin_rate: Option<f64>,
    npv: Option<f64>,
    npv_rate: Option<f64>,
    irr: Option<f64>,
    dynamic_payback: Option<f64>,
    risk_level: String,
}
#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ProjectRow {
    id: String,
    name: String,
    customer_name: String,
    status: String,
    benefit_status: String,
    created_at: String,
    updated_at: String,
    progress: Option<f64>,
    project_years: i32,
    discount_rate: Option<f64>,
    total_revenue_incl: Option<f64>,
    total_cost_incl: Option<f64>,
    summary_metrics: Option<Metrics>,
    is_bound_project: bool,
    default_scheme_id: Option<String>,
    stage: String,
    stage_label: &'static str,
    scheme_name: Option<String>,
    scheme_updated_at: Option<String>,
}
#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Totals {
    total_revenue_incl: String,
    total_cost_incl: String,
    revenue_value_count: usize,
    cost_value_count: usize,
}
#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct StageTotals {
    stage: String,
    stage_label: &'static str,
    matched_count: usize,
    totals: Totals,
}
#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct QueryResponse {
    scope: &'static str,
    source: &'static str,
    notice: &'static str,
    bound_project_id: Option<String>,
    matched_count: usize,
    returned_count: usize,
    applied_limit: usize,
    truncated: bool,
    message: String,
    totals: Option<Totals>,
    mixed_stages: bool,
    comparison_notice: &'static str,
    stage_totals: Vec<StageTotals>,
    projects: Vec<ProjectRow>,
    queried_at: String,
}
fn clip(value: &str, max: usize) -> String {
    if value.chars().count() > max {
        format!("{}…", value.chars().take(max).collect::<String>())
    } else {
        value.into()
    }
}
fn metric(value: &str, rate: bool) -> Option<f64> {
    let trimmed = value.trim();
    let (value, scale) = if rate && trimmed.ends_with('%') {
        (trimmed.trim_end_matches('%'), 100.0)
    } else {
        (trimmed, 1.0)
    };
    value
        .parse::<f64>()
        .ok()
        .filter(|v| v.is_finite())
        .map(|v| v / scale)
}
fn timestamp(value: &str) -> Option<i64> {
    DateTime::parse_from_rfc3339(value)
        .ok()
        .map(|d| d.timestamp_millis())
        .or_else(|| {
            NaiveDateTime::parse_from_str(value, "%Y-%m-%d %H:%M:%S")
                .ok()
                .map(|d| d.and_utc().timestamp_millis())
        })
        .or_else(|| {
            NaiveDate::parse_from_str(value, "%Y-%m-%d")
                .ok()
                .and_then(|d| d.and_hms_opt(0, 0, 0))
                .map(|d| d.and_utc().timestamp_millis())
        })
}
fn project_row(
    p: Project,
    bound: Option<&str>,
    scheme: Option<BenefitAnalysisScheme>,
) -> ProjectRow {
    let stage = scheme
        .as_ref()
        .and_then(|s| s.stage.as_deref())
        .unwrap_or("unlabeled");
    let (stage, stage_label) = match stage {
        "pre_selection" => ("pre_selection", "甄选前（限价口径）"),
        "post_selection" => ("post_selection", "甄选后（中标口径）"),
        _ => ("unlabeled", "未标注"),
    };
    ProjectRow {
        default_scheme_id: p.default_scheme_id.clone(),
        stage: stage.into(),
        stage_label,
        scheme_name: scheme.as_ref().map(|s| clip(&s.name, 256)),
        scheme_updated_at: scheme.as_ref().map(|s| clip(&s.updated_at, 40)),
        is_bound_project: bound == Some(p.id.as_str()),
        id: p.id,
        name: clip(&p.name, 256),
        customer_name: clip(&p.customer_name, 128),
        status: clip(&p.status, 64),
        benefit_status: clip(&p.benefit_status, 64),
        created_at: clip(&p.created_at, 40),
        updated_at: clip(&p.updated_at, 40),
        progress: p.progress.is_finite().then_some(p.progress),
        project_years: p.project_years,
        discount_rate: p.discount_rate.is_finite().then_some(p.discount_rate),
        total_revenue_incl: p
            .total_revenue_incl
            .is_finite()
            .then_some(p.total_revenue_incl),
        total_cost_incl: p.total_cost_incl.is_finite().then_some(p.total_cost_incl),
        summary_metrics: p.summary_metrics.map(|m| Metrics {
            margin_rate: metric(&m.margin_rate, true),
            npv: metric(&m.npv, false),
            npv_rate: metric(&m.npv_rate, true),
            irr: metric(&m.irr, true),
            dynamic_payback: metric(&m.dynamic_payback, false),
            risk_level: clip(&m.risk_level, 64),
        }),
    }
}
fn compare<T: PartialOrd>(a: Option<T>, b: Option<T>, order: SortOrder) -> std::cmp::Ordering {
    use std::cmp::Ordering;
    match (a, b) {
        (Some(a), Some(b)) => {
            let cmp = a.partial_cmp(&b).unwrap_or(Ordering::Equal);
            if matches!(order, SortOrder::Desc) {
                cmp.reverse()
            } else {
                cmp
            }
        }
        (Some(_), None) => Ordering::Less,
        (None, Some(_)) => Ordering::Greater,
        _ => Ordering::Equal,
    }
}
impl ProjectRow {
    fn sort_number(&self, field: SortField) -> Option<f64> {
        match field {
            SortField::TotalRevenueIncl => self.total_revenue_incl,
            SortField::TotalCostIncl => self.total_cost_incl,
            SortField::MarginRate => self.summary_metrics.as_ref()?.margin_rate,
            SortField::Npv => self.summary_metrics.as_ref()?.npv,
            SortField::NpvRate => self.summary_metrics.as_ref()?.npv_rate,
            SortField::Irr => self.summary_metrics.as_ref()?.irr,
            SortField::DynamicPayback => self.summary_metrics.as_ref()?.dynamic_payback,
            _ => None,
        }
    }
}
/// The same service method used by the existing get_projects Tauri command.
pub fn query_projects(
    service: &ProjectService,
    request: &QueryRequest,
    bound: Option<&str>,
) -> Result<QueryResponse, String> {
    if request.irr.is_some() || matches!(request.sort_by, SortField::Irr) {
        return Err("本系统未计算 IRR，不能按 IRR 筛选或排序；这不表示没有符合条件的项目。".into());
    }
    for range in [
        &request.total_revenue_incl,
        &request.total_cost_incl,
        &request.margin_rate,
        &request.npv,
        &request.npv_rate,
        &request.irr,
        &request.dynamic_payback,
    ]
    .into_iter()
    .flatten()
    {
        range.validate()?;
    }
    if request.limit == Some(0) {
        return Err("limit 必须大于零".into());
    }
    for filter in [
        &request.customer_name,
        &request.status,
        &request.benefit_status,
    ]
    .into_iter()
    .flatten()
    {
        if filter.trim().is_empty() || filter.chars().count() > 128 {
            return Err("文本筛选需为 1–128 字符".into());
        }
    }
    let dates: Vec<Option<i64>> = [
        &request.created_from,
        &request.created_before,
        &request.updated_from,
        &request.updated_before,
    ]
    .into_iter()
    .map(|value| {
        value
            .as_ref()
            .map(|s| timestamp(s).ok_or("时间须为 RFC3339 或 YYYY-MM-DD；日期按 UTC 零点解释"))
            .transpose()
    })
    .collect::<Result<_, _>>()?;
    for pair in dates.chunks(2) {
        if pair[0].zip(pair[1]).is_some_and(|(a, b)| a >= b) {
            return Err("时间范围起点必须早于终点（终点不含）".into());
        }
    }
    let customer = request
        .customer_name
        .as_ref()
        .map(|s| s.trim().to_lowercase());
    // get_schemes selects only scheme identity/stage/name/timestamps; it never
    // reads snapshots or financial detail. Follow the captured default id only.
    let projects = service.get_projects()?;
    let projected = projects
        .into_iter()
        .map(|p| {
            let scheme = if let Some(id) = &p.default_scheme_id {
                service
                    .get_schemes(&p.id)?
                    .into_iter()
                    .find(|s| &s.id == id)
            } else {
                None
            };
            Ok(project_row(p, bound, scheme))
        })
        .collect::<Result<Vec<_>, String>>()?;
    let mut rows: Vec<ProjectRow> = projected
        .into_iter()
        .filter(|p| {
            customer
                .as_ref()
                .is_none_or(|s| p.customer_name.to_lowercase().contains(s))
                && request.status.as_ref().is_none_or(|s| p.status == s.trim())
                && request
                    .benefit_status
                    .as_ref()
                    .is_none_or(|s| p.benefit_status == s.trim())
        })
        .filter(|p| {
            let times = [timestamp(&p.created_at), timestamp(&p.updated_at)];
            let time_ok = times.iter().zip(dates.chunks(2)).all(|(value, pair)| {
                pair[0].is_none_or(|from| value.is_some_and(|v| v >= from))
                    && pair[1].is_none_or(|before| value.is_some_and(|v| v < before))
            });
            time_ok
                && [
                    (&request.total_revenue_incl, p.total_revenue_incl),
                    (&request.total_cost_incl, p.total_cost_incl),
                    (&request.margin_rate, p.sort_number(SortField::MarginRate)),
                    (&request.npv, p.sort_number(SortField::Npv)),
                    (&request.npv_rate, p.sort_number(SortField::NpvRate)),
                    (&request.irr, p.sort_number(SortField::Irr)),
                    (
                        &request.dynamic_payback,
                        p.sort_number(SortField::DynamicPayback),
                    ),
                ]
                .into_iter()
                .all(|(range, value)| range.as_ref().is_none_or(|r| r.matches(value)))
        })
        .collect();
    let stages: std::collections::BTreeSet<_> = rows.iter().map(|r| r.stage.clone()).collect();
    let mixed_stages = stages.len() > 1;
    rows.sort_by(|a, b| {
        // Financial rankings are meaningful within one selection-stage basis.
        if mixed_stages
            && !matches!(
                request.sort_by,
                SortField::Name | SortField::CreatedAt | SortField::UpdatedAt
            )
        {
            let stage_order = a.stage.cmp(&b.stage);
            if !stage_order.is_eq() {
                return stage_order;
            }
        }
        let cmp = match request.sort_by {
            SortField::Name => compare(Some(&a.name), Some(&b.name), request.sort_order),
            SortField::CreatedAt => compare(
                timestamp(&a.created_at),
                timestamp(&b.created_at),
                request.sort_order,
            ),
            SortField::UpdatedAt => compare(
                timestamp(&a.updated_at),
                timestamp(&b.updated_at),
                request.sort_order,
            ),
            field => compare(
                a.sort_number(field),
                b.sort_number(field),
                request.sort_order,
            ),
        };
        cmp.then_with(|| a.id.cmp(&b.id))
    });
    let stage_totals = stages
        .iter()
        .map(|stage| {
            let matching: Vec<_> = rows.iter().filter(|r| &r.stage == stage).collect();
            Ok(StageTotals {
                stage: stage.clone(),
                stage_label: matching[0].stage_label,
                matched_count: matching.len(),
                totals: totals_for(matching.into_iter())?,
            })
        })
        .collect::<Result<Vec<_>, String>>()?;
    let totals = if mixed_stages {
        None
    } else {
        Some(totals_for(rows.iter())?)
    };
    let matched = rows.len();
    let limit = request.limit.unwrap_or(20).min(MAX_LIMIT);
    rows.truncate(limit);
    Ok(QueryResponse {
        scope: "workspace_project_search",
        source: "projects + summary_metrics (saved)",
        notice: NOTICE,
        bound_project_id: bound.map(str::to_string),
        matched_count: matched,
        returned_count: rows.len(),
        applied_limit: limit,
        truncated: matched > rows.len(),
        message: format!(
            "命中 {matched} 条，{}返回 {} 条（服务端上限 {MAX_LIMIT}）",
            if matched > rows.len() { "仅" } else { "已" },
            rows.len()
        ),
        totals,
        mixed_stages,
        comparison_notice: if mixed_stages {
            "本次结果包含不同甄选阶段的口径，不可直接横向比较。金额仅按阶段合计；财务指标在阶段内排序。"
        } else {
            "本次命中为同一阶段标签；未标注不代表已确认相同口径。请结合方案名和保存时间使用指标。"
        },
        stage_totals,
        projects: rows,
        queried_at: Utc::now().to_rfc3339(),
    })
}

fn totals_for<'a>(rows: impl Iterator<Item = &'a ProjectRow>) -> Result<Totals, String> {
    let mut revenue = Decimal::ZERO;
    let mut cost = Decimal::ZERO;
    let mut revenue_count = 0;
    let mut cost_count = 0;
    for row in rows {
        for (value, sum, count) in [
            (row.total_revenue_incl, &mut revenue, &mut revenue_count),
            (row.total_cost_incl, &mut cost, &mut cost_count),
        ] {
            if let Some(value) = value {
                let number = Decimal::from_str(&value.to_string())
                    .map_err(|_| "项目汇总金额超出可表示范围")?;
                *sum = sum.checked_add(number).ok_or("项目汇总金额溢出")?;
                *count += 1;
            }
        }
    }
    Ok(Totals {
        total_revenue_incl: revenue.normalize().to_string(),
        total_cost_incl: cost.normalize().to_string(),
        revenue_value_count: revenue_count,
        cost_value_count: cost_count,
    })
}
