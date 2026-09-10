//! Session-long permissions. Deliberately separate from per-turn stream routing.
use serde::{Deserialize, Serialize};
use std::{collections::HashMap, path::PathBuf, sync::Mutex};

pub const AUTHORIZE_ROUTE: &str = "/lamber-bridge/authorize";
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProjectBinding {
    pub workspace_id: String,
    pub cwd: PathBuf,
    // Explicit None is general chat; an absent registry entry is legacy/unbound.
    pub project_id: Option<String>,
}
/// Explicit classification: absence of projectId is never an authorization rule.
const AGGREGATE_READ_TOOLS: &[&str] = &[super::project_query::QUERY_TOOL];
pub enum AuthorizedScope {
    Project(String),
    PureCalculation,
    AggregateRead { bound_project_id: Option<String> },
}
impl AuthorizedScope {
    pub fn bound_project_id(&self) -> Option<&str> {
        match self {
            Self::PureCalculation => None,
            Self::Project(id) => Some(id),
            Self::AggregateRead { bound_project_id } => bound_project_id.as_deref(),
        }
    }
}
#[derive(Default)]
pub struct ProjectBindings(Mutex<HashMap<String, ProjectBinding>>);
impl ProjectBindings {
    pub fn register(&self, session: &str, binding: ProjectBinding) -> Result<(), String> {
        if session.trim().is_empty() {
            return Err("缺少可信 AI 会话身份".into());
        }
        let mut entries = self.0.lock().map_err(|_| "AI 项目绑定锁不可用")?;
        if entries
            .get(session)
            .is_some_and(|existing| existing != &binding)
        {
            return Err("AI 会话不能改绑，请新建会话".into());
        }
        entries.insert(session.into(), binding);
        Ok(())
    }
    pub fn forget(&self, session: &str) -> Result<(), String> {
        self.0
            .lock()
            .map_err(|_| "AI 项目绑定锁不可用")?
            .remove(session);
        Ok(())
    }
    pub fn authorize(
        &self,
        session: &str,
        tool: &str,
        project: Option<&str>,
        workspace_id: &str,
        cwd: &std::path::Path,
    ) -> Result<AuthorizedScope, String> {
        let entries = self.0.lock().map_err(|_| "AI 项目绑定锁不可用")?;
        let binding = entries
            .get(session)
            .ok_or("此 AI 会话尚未绑定，请新建会话并选择项目")?;
        if binding.workspace_id != workspace_id || binding.cwd != cwd {
            return Err("此 AI 会话属于另一个工作区".into());
        }
        if binding
            .project_id
            .as_ref()
            .is_some_and(|id| id.trim().is_empty())
        {
            return Err("项目绑定身份无效".into());
        }
        if ["calculate_selection_fee", "reverse_calculate_selection_fee"].contains(&tool) {
            if project.is_some() { return Err("纯计算工具不接受 projectId 参数".into()); }
            return Ok(AuthorizedScope::PureCalculation);
        }
        if AGGREGATE_READ_TOOLS.contains(&tool) {
            if project.is_some() {
                return Err("聚合工具不接受 projectId 参数".into());
            }
            return Ok(AuthorizedScope::AggregateRead {
                bound_project_id: binding.project_id.clone(),
            });
        }
        let bound = binding
            .project_id
            .as_deref()
            .filter(|id| !id.trim().is_empty())
            .ok_or("通用聊天仅能调用聚合只读及甄选费纯计算工具，请新建项目会话")?;
        match tool {
            "run_benefit_calculation" | "fill_template_fields" if project == Some(bound) => {
                Ok(AuthorizedScope::Project(bound.into()))
            }
            // These tools derive their project from the binding; neither accepts projectId.
            "write_test_marker" | "read_template_fields" | "read_benefit_inputs" | "simulate_benefit_calculation" if project.is_none() => Ok(AuthorizedScope::Project(bound.into())),
            "run_benefit_calculation" | "fill_template_fields" => {
                Err("工具请求被拒绝：只能访问会话绑定的项目，且必须提供 projectId".into())
            }
            _ => Err("该工具尚未定义项目权限，已拒绝调用".into()),
        }
    }
}
