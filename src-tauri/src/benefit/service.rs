use super::models::{
    BenefitAnalysisScheme, BenefitAnalysisSnapshot, IctInput, IctResult, Project, ProjectLog,
    SummaryMetrics,
};
use super::repository::ProjectRepository;
use std::time::{SystemTime, UNIX_EPOCH};

pub struct RiskConfig {
    pub margin_high_threshold: f64,
    pub margin_medium_threshold: f64,
    pub npv_high_threshold: f64,
    pub npvr_medium_threshold: f64,
}

impl Default for RiskConfig {
    fn default() -> Self {
        Self {
            margin_high_threshold: 0.0,
            margin_medium_threshold: 0.08,
            npv_high_threshold: 0.0,
            npvr_medium_threshold: 0.04,
        }
    }
}

pub struct ProjectService {
    repository: Box<dyn ProjectRepository + Send + Sync>,
}

fn generate_id() -> String {
    let start = SystemTime::now();
    let since_the_epoch = start
        .duration_since(UNIX_EPOCH)
        .unwrap_or_else(|_| std::time::Duration::from_secs(0));
    format!(
        "id_{}_{}",
        since_the_epoch.as_millis(),
        since_the_epoch.subsec_nanos()
    )
}

fn normalize_project_name(name: &str) -> String {
    name.split_whitespace()
        .collect::<Vec<_>>()
        .join(" ")
        .to_lowercase()
}

fn get_incl(item: &super::models::IctItem) -> f64 {
    item.incl_tax.parse::<f64>().unwrap_or(0.0)
}

fn get_total_revenue_incl(input: &IctInput) -> f64 {
    get_incl(&input.rev_it_integration)
        + get_incl(&input.rev_it_maintenance)
        + get_incl(&input.rev_it_device_sales)
        + get_incl(&input.rev_it_device_lease)
        + get_incl(&input.rev_it_other)
        + get_incl(&input.rev_it_cloud)
        + get_incl(&input.rev_ct_line)
        + get_incl(&input.rev_ct_product)
        + get_incl(&input.rev_non_it_ct)
}

fn get_total_cost_incl(input: &IctInput) -> f64 {
    get_incl(&input.cost_it_device)
        + get_incl(&input.cost_it_construction)
        + get_incl(&input.cost_it_survey)
        + get_incl(&input.cost_it_integration)
        + get_incl(&input.cost_it_other)
        + get_incl(&input.cost_it_maintenance)
        + get_incl(&input.cost_it_running)
        + get_incl(&input.cost_it_bidding)
        + get_incl(&input.cost_it_design_eval)
        + get_incl(&input.cost_it_audit)
        + get_incl(&input.cost_ct_construction)
        + get_incl(&input.cost_ct_maintenance)
        + get_incl(&input.cost_ct_other)
        + get_incl(&input.cost_ct_bandwidth)
        + get_incl(&input.cost_ct_renewal)
        + get_incl(&input.cost_non_it_ct)
        + get_incl(&input.cost_mix_marketing)
        + get_incl(&input.cost_mix_channel)
        + get_incl(&input.cost_mix_other)
}

pub fn compute_fingerprint(input: &IctInput) -> String {
    let total_revenue_incl = get_total_revenue_incl(input);
    let total_cost_incl = get_total_cost_incl(input);
    let project_years = input.project_years.unwrap_or(1);
    let discount_rate = input.discount_rate.parse::<f64>().unwrap_or(0.0);
    let cashflow_model = input
        .cashflow_model
        .clone()
        .unwrap_or_else(|| "model_a".to_string());
    format!(
        "{:.2}:{:.2}:{}:{:.4}:{}",
        total_revenue_incl, total_cost_incl, project_years, discount_rate, cashflow_model
    )
}

impl ProjectService {
    pub fn new(repository: Box<dyn ProjectRepository + Send + Sync>) -> Self {
        Self { repository }
    }

    pub fn get_projects(&self) -> Result<Vec<Project>, String> {
        self.repository.get_projects()
    }

    pub fn get_project(&self, id: &str) -> Result<Option<Project>, String> {
        self.repository.get_project(id)
    }

    pub fn create_project(&self, name: String, customer_name: String) -> Result<Project, String> {
        let name = name.trim().to_string();
        if name.is_empty() {
            return Err("项目名称不能为空".to_string());
        }
        let normalized_name = normalize_project_name(&name);
        let existing_projects = self.repository.get_projects()?;
        if existing_projects
            .iter()
            .any(|project| normalize_project_name(&project.name) == normalized_name)
        {
            return Err(format!("项目名称已存在：{}", name));
        }

        let customer_name = {
            let trimmed = customer_name.trim();
            if trimmed.is_empty() {
                "未知客户".to_string()
            } else {
                trimmed.to_string()
            }
        };
        let timestamp = chrono::Utc::now().to_rfc3339();
        let project = Project {
            id: generate_id(),
            name,
            customer_name,
            status: "需求导入".to_string(),
            benefit_status: "not_started".to_string(),
            default_scheme_id: None,
            created_at: timestamp.clone(),
            updated_at: timestamp,
            total_revenue_incl: 0.0,
            total_cost_incl: 0.0,
            project_years: 1,
            discount_rate: 0.055,
            cashflow_model: "model_a".to_string(),
            summary_metrics: None,
            folder_path: None,
            main_document_path: None,
            main_budget_file_path: None,
            note: None,
            logs: vec![],
            folder_name: None,
            relative_path: None,
            progress: 0.0,
            deadline: None,
            linked_folder_type: Some("none".to_string()),
            linked_folder_relative_path: None,
            linked_folder_external_path: None,
        };
        self.repository.save_project(&project)?;
        Ok(project)
    }

    pub fn update_project(&self, mut project: Project) -> Result<Project, String> {
        project.name = project.name.trim().to_string();
        if project.name.is_empty() {
            return Err("项目名称不能为空".to_string());
        }
        project.customer_name = {
            let trimmed = project.customer_name.trim();
            if trimmed.is_empty() {
                "未知客户".to_string()
            } else {
                trimmed.to_string()
            }
        };
        let normalized_name = normalize_project_name(&project.name);
        let existing_projects = self.repository.get_projects()?;
        if existing_projects.iter().any(|existing| {
            existing.id != project.id && normalize_project_name(&existing.name) == normalized_name
        }) {
            return Err(format!("项目名称已存在：{}", project.name));
        }

        let original_project_opt = self.repository.get_project(&project.id)?;
        if original_project_opt.is_none() {
            return Err(format!("ProjectNotFoundInCurrentWorkspace::{}", project.id));
        }

        // Calculate/verify status based on fingerprint mismatch
        if let Some(default_scheme_id) = &project.default_scheme_id {
            let snapshots = self.repository.get_snapshots(default_scheme_id)?;
            if let Some(latest) = snapshots.first() {
                // Compute updated project details fingerprint
                let project_fp = format!(
                    "{:.2}:{:.2}:{}:{:.4}:{}",
                    project.total_revenue_incl,
                    project.total_cost_incl,
                    project.project_years,
                    project.discount_rate,
                    project.cashflow_model
                );
                if project_fp == latest.fingerprint {
                    project.benefit_status = "normal".to_string();
                } else {
                    project.benefit_status = "outdated".to_string();
                }
            } else {
                project.benefit_status = "not_started".to_string();
            }
        } else {
            project.benefit_status = "not_started".to_string();
        }

        // Add a log entry if project metadata changed (e.g. name, customer_name, status, or outdated status)
        let mut logs = project.logs.clone();
        if let Some(orig) = original_project_opt {
            let mut changes = Vec::new();
            if orig.name != project.name {
                changes.push(format!("名称变更为 '{}'", project.name));
            }
            if orig.customer_name != project.customer_name {
                changes.push(format!("客户名称变更为 '{}'", project.customer_name));
            }
            if orig.status != project.status {
                changes.push(format!("状态变更为 '{}'", project.status));
            }
            if orig.note != project.note {
                changes.push("备注更新".to_string());
            }
            if orig.benefit_status != project.benefit_status {
                changes.push(format!("效益分析状态变更为 '{}'", project.benefit_status));
            }
            if !changes.is_empty() {
                logs.push(ProjectLog {
                    id: generate_id(),
                    timestamp: chrono::Utc::now().to_rfc3339(),
                    description: format!("更新项目信息: {}", changes.join(", ")),
                });
            }
        }
        project.logs = logs;
        project.updated_at = chrono::Utc::now().to_rfc3339();

        self.repository.save_project(&project)?;
        Ok(project)
    }

    pub fn delete_project(&self, id: &str) -> Result<(), String> {
        self.repository.delete_project(id)
    }

    pub fn calculate_risk_level(&self, result: &IctResult, config: &RiskConfig) -> String {
        let npv_val = result.npv.parse::<f64>().unwrap_or(0.0);

        // margin_rate is e.g. "20.5%" or "0.205" or "20.50%". Let's handle percentage signs
        let margin_str = result.margin_rate.trim_end_matches('%');
        let margin_val = margin_str.parse::<f64>().unwrap_or(0.0)
            / if result.margin_rate.contains('%') {
                100.0
            } else {
                1.0
            };

        let npv_rate_str = result.npv_rate.trim_end_matches('%');
        let npv_rate_val = npv_rate_str.parse::<f64>().unwrap_or(0.0)
            / if result.npv_rate.contains('%') {
                100.0
            } else {
                1.0
            };

        if npv_val < config.npv_high_threshold || margin_val < config.margin_high_threshold {
            "高风险".to_string()
        } else if margin_val < config.margin_medium_threshold
            || npv_rate_val < config.npvr_medium_threshold
        {
            "中风险".to_string()
        } else {
            "低风险".to_string()
        }
    }

    pub fn save_benefit_scheme(
        &self,
        project_id: String,
        scheme_id_opt: Option<String>,
        scheme_name: String,
        input_params: IctInput,
        output_metrics: IctResult,
        is_save_as_new: bool,
    ) -> Result<Project, String> {
        let mut project = self
            .repository
            .get_project(&project_id)?
            .ok_or_else(|| format!("Project {} not found", project_id))?;

        let timestamp = chrono::Utc::now().to_rfc3339();

        let existing_schemes = self.repository.get_schemes(&project_id)?;
        let (scheme_id, is_new_scheme) = if is_save_as_new {
            (generate_id(), true)
        } else if let Some(scheme_id) = scheme_id_opt {
            (scheme_id, false)
        } else if let Some(existing) = existing_schemes.iter().find(|s| s.name == scheme_name) {
            (existing.id.clone(), false)
        } else {
            (generate_id(), true)
        };

        if is_new_scheme {
            let scheme = BenefitAnalysisScheme {
                id: scheme_id.clone(),
                project_id: project_id.clone(),
                name: scheme_name,
                created_at: timestamp.clone(),
                updated_at: timestamp.clone(),
            };
            self.repository.save_scheme(&scheme)?;
        } else {
            if let Some(mut existing) = existing_schemes.into_iter().find(|s| s.id == scheme_id) {
                existing.name = scheme_name;
                existing.updated_at = timestamp.clone();
                self.repository.save_scheme(&existing)?;
            } else {
                let scheme = BenefitAnalysisScheme {
                    id: scheme_id.clone(),
                    project_id: project_id.clone(),
                    name: scheme_name,
                    created_at: timestamp.clone(),
                    updated_at: timestamp.clone(),
                };
                self.repository.save_scheme(&scheme)?;
            }
        }

        let mut version = 1;
        if !is_new_scheme {
            let snapshots = self.repository.get_snapshots(&scheme_id)?;
            if let Some(latest) = snapshots.first() {
                version = latest.version + 1;
            }
        }

        let fp = compute_fingerprint(&input_params);

        let snapshot = BenefitAnalysisSnapshot {
            id: generate_id(),
            scheme_id: scheme_id.clone(),
            project_id: project_id.clone(),
            version,
            input_params: input_params.clone(),
            output_metrics: output_metrics.clone(),
            fingerprint: fp,
            created_at: timestamp.clone(),
        };
        self.repository.save_snapshot(&snapshot)?;

        // Update Project context:
        project.default_scheme_id = Some(scheme_id);
        project.benefit_status = "normal".to_string();

        // Sync fingerprint parameters from calculations to project
        project.total_revenue_incl = get_total_revenue_incl(&input_params);
        project.total_cost_incl = get_total_cost_incl(&input_params);
        project.project_years = input_params.project_years.unwrap_or(1);
        project.discount_rate = input_params.discount_rate.parse::<f64>().unwrap_or(0.055);
        project.cashflow_model = input_params
            .cashflow_model
            .clone()
            .unwrap_or_else(|| "model_a".to_string());

        // Calculate risk level via service logic
        let risk_level = self.calculate_risk_level(&output_metrics, &RiskConfig::default());

        let summary = SummaryMetrics {
            margin_rate: output_metrics.margin_rate.clone(),
            npv: output_metrics.npv.clone(),
            npv_rate: output_metrics.npv_rate.clone(),
            irr: output_metrics.irr.clone(),
            dynamic_payback: output_metrics.dynamic_payback.clone(),
            risk_level,
        };
        project.summary_metrics = Some(summary);

        // Add to log
        project.logs.push(ProjectLog {
            id: generate_id(),
            timestamp,
            description: format!(
                "更新效益分析方案：毛利率 {}, NPV {}, NPVR {}, IRR {}, 风险等级 {}",
                output_metrics.margin_rate,
                output_metrics.npv,
                output_metrics.npv_rate,
                output_metrics.irr,
                project.summary_metrics.as_ref().unwrap().risk_level
            ),
        });

        self.repository.save_project(&project)?;

        Ok(project)
    }

    pub fn delete_benefit_scheme(
        &self,
        project_id: String,
        scheme_id: String,
    ) -> Result<Project, String> {
        let mut project = self
            .repository
            .get_project(&project_id)?
            .ok_or_else(|| format!("Project {} not found", project_id))?;

        let schemes = self.repository.get_schemes(&project_id)?;
        let deleted_scheme = schemes
            .iter()
            .find(|s| s.id == scheme_id)
            .cloned()
            .ok_or_else(|| "未找到指定测算方案".to_string())?;

        self.repository.delete_scheme(&project_id, &scheme_id)?;

        let remaining_schemes = self.repository.get_schemes(&project_id)?;
        let next_default = remaining_schemes
            .iter()
            .max_by(|a, b| a.updated_at.cmp(&b.updated_at))
            .cloned();

        if let Some(scheme) = next_default {
            project.default_scheme_id = Some(scheme.id.clone());
            let snapshots = self.repository.get_snapshots(&scheme.id)?;
            if let Some(latest) = snapshots.first() {
                project.benefit_status = "normal".to_string();
                project.total_revenue_incl = get_total_revenue_incl(&latest.input_params);
                project.total_cost_incl = get_total_cost_incl(&latest.input_params);
                project.project_years = latest.input_params.project_years.unwrap_or(1);
                project.discount_rate = latest
                    .input_params
                    .discount_rate
                    .parse::<f64>()
                    .unwrap_or(0.055);
                project.cashflow_model = latest
                    .input_params
                    .cashflow_model
                    .clone()
                    .unwrap_or_else(|| "model_a".to_string());

                let risk_level =
                    self.calculate_risk_level(&latest.output_metrics, &RiskConfig::default());
                project.summary_metrics = Some(SummaryMetrics {
                    margin_rate: latest.output_metrics.margin_rate.clone(),
                    npv: latest.output_metrics.npv.clone(),
                    npv_rate: latest.output_metrics.npv_rate.clone(),
                    irr: latest.output_metrics.irr.clone(),
                    dynamic_payback: latest.output_metrics.dynamic_payback.clone(),
                    risk_level,
                });
            } else {
                project.benefit_status = "not_started".to_string();
                project.summary_metrics = None;
            }
        } else {
            project.default_scheme_id = None;
            project.benefit_status = "not_started".to_string();
            project.total_revenue_incl = 0.0;
            project.total_cost_incl = 0.0;
            project.summary_metrics = None;
        }

        let timestamp = chrono::Utc::now().to_rfc3339();
        project.updated_at = timestamp.clone();
        project.logs.push(ProjectLog {
            id: generate_id(),
            timestamp,
            description: format!("删除效益分析方案：{}", deleted_scheme.name),
        });

        self.repository.save_project(&project)?;
        Ok(project)
    }

    pub fn get_schemes(&self, project_id: &str) -> Result<Vec<BenefitAnalysisScheme>, String> {
        self.repository.get_schemes(project_id)
    }

    pub fn get_snapshots(&self, scheme_id: &str) -> Result<Vec<BenefitAnalysisSnapshot>, String> {
        self.repository.get_snapshots(scheme_id)
    }
}
