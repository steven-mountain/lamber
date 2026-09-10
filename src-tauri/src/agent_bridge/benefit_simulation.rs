//! Main-window pure preparation: no editor navigation, writes, or model-supplied input set.
use super::{
    benefit_access,
    bridge_server::{BridgeHandler, BridgeReply},
    calculation,
    project_bindings::ProjectBindings,
};
use crate::{
    benefit::{models::IctInput, repository::SqliteProjectRepository, service::ProjectService},
    workspace::WorkspaceRuntime,
};
use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use std::{
    collections::HashMap,
    sync::{mpsc, Arc, Mutex},
    time::Duration,
};
use tauri::Emitter;
pub const ROUTE: &str = "/lamber-bridge/simulate-benefit-calculation";
const EVENT: &str = "lamber-prepare-benefit-simulation";
#[derive(Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct Override {
    pub subject: String,
    pub incl_tax: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tax_rate: Option<String>,
}
#[derive(Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Job {
    pub input: IctInput,
    pub overrides: Vec<Override>,
}
#[derive(Deserialize, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct Change {
    kind: String,
    subject: String,
    before: String,
    after: String,
}
#[derive(Deserialize, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct Prepared {
    pub input: IctInput,
    pub explicit_changes: Vec<Change>,
    pub linked_changes: Vec<Change>,
    pub tax_incl_auto_fix: bool,
}
struct Pending {
    job: Option<Job>,
    sender: mpsc::Sender<Result<Prepared, String>>,
}
#[derive(Default)]
pub struct SimulationJobs(Mutex<HashMap<String, Pending>>);
pub type Preparer = Arc<dyn Fn(Job) -> Result<Prepared, String> + Send + Sync>;
impl SimulationJobs {
    pub fn prepare(
        &self,
        job: Job,
        announce: impl FnOnce(&str) -> Result<(), String>,
    ) -> Result<Prepared, String> {
        let id = uuid::Uuid::new_v4().to_string();
        let (sender, receiver) = mpsc::channel();
        self.0.lock().map_err(|_| "试算服务锁不可用")?.insert(
            id.clone(),
            Pending {
                job: Some(job),
                sender,
            },
        );
        let result = announce(&id).and_then(|_| {
            receiver
                .recv_timeout(Duration::from_secs(15))
                .map_err(|_| {
                    "桌面试算准备服务未响应；未计算或保存任何假设值，请重新打开主窗口后重试。"
                        .to_string()
                })?
        });
        self.0.lock().map_err(|_| "试算服务锁不可用")?.remove(&id);
        result
    }
    fn claim(&self, id: &str) -> Result<Job, String> {
        self.0
            .lock()
            .map_err(|_| "试算服务锁不可用")?
            .get_mut(id)
            .and_then(|p| p.job.take())
            .ok_or("试算请求已领取或失效".into())
    }
    fn finish(&self, id: &str, result: Result<Prepared, String>) -> Result<(), String> {
        let mut jobs = self.0.lock().map_err(|_| "试算服务锁不可用")?;
        if jobs.get(id).is_some_and(|p| p.job.is_some()) {
            return Err("试算请求尚未领取".into());
        }
        jobs.remove(id)
            .ok_or("试算请求已失效")?
            .sender
            .send(result)
            .map_err(|_| "试算请求已结束".into())
    }
}
#[tauri::command]
pub fn ai_claim_benefit_simulation(
    window: tauri::WebviewWindow,
    state: tauri::State<'_, Arc<super::AgentRuntime>>,
    id: String,
) -> Result<Job, String> {
    if window.label() != "main" {
        return Err("只有主窗口可准备试算".into());
    }
    state.simulations.claim(&id)
}
#[tauri::command]
pub fn ai_finish_benefit_simulation(
    window: tauri::WebviewWindow,
    state: tauri::State<'_, Arc<super::AgentRuntime>>,
    id: String,
    prepared: Option<Prepared>,
    error: Option<String>,
) -> Result<(), String> {
    if window.label() != "main" {
        return Err("只有主窗口可完成试算".into());
    }
    let result = match (prepared, error) {
        (Some(value), None) => Ok(value),
        (None, Some(error)) => Err(error),
        _ => return Err("试算回执必须是结果或错误".into()),
    };
    state.simulations.finish(&id, result)
}
pub fn app_preparer(app: tauri::AppHandle, jobs: Arc<SimulationJobs>) -> Preparer {
    Arc::new(move |job| {
        jobs.prepare(job, |id| {
            app.emit_to("main", EVENT, id).map_err(|e| e.to_string())
        })
    })
}
#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct Request {
    session_id: String,
    scenario: Option<String>,
    overrides: Vec<Override>,
}
pub fn handler(
    runtime: Arc<WorkspaceRuntime>,
    bindings: Arc<ProjectBindings>,
    prepare: Preparer,
    next: BridgeHandler,
) -> BridgeHandler {
    Arc::new(move |path, body| {
        if path != ROUTE {
            return next(path, body);
        }
        let result = (|| -> Result<Value, String> {
            let request: Request = serde_json::from_str(body).map_err(|e| {
                format!("试算只接受 scenario 和具名 overrides，不能提供完整输入或 projectId：{e}")
            })?;
            if request.overrides.len() > 28 {
                return Err("最多接受28个具名覆盖项".into());
            }
            let (workspace, conn) = runtime.require_context()?;
            let cwd =
                std::fs::canonicalize(&workspace.workspace_root).map_err(|_| "工作区不可用")?;
            let scope = bindings.authorize(
                &request.session_id,
                "simulate_benefit_calculation",
                None,
                &workspace.workspace_id,
                &cwd,
            )?;
            let service = ProjectService::new(Box::new(SqliteProjectRepository::new(conn)));
            let calc_request = calculation::CalculateRequest {
                project_id: scope.bound_project_id().ok_or("要求绑定项目")?.into(),
                scenario: request.scenario,
            };
            let (project, scheme, snapshot) =
                calculation::resolve_snapshot(&service, &calc_request)?;
            // Validate the exact resolved snapshot; never resolve twice across a concurrent save.
            benefit_access::project_inputs(&project, &scheme, &snapshot)?;
            drop(service);
            let has_overrides = !request.overrides.is_empty();
            let prepared = if !has_overrides {
                Prepared {
                    input: snapshot.input_params.clone(),
                    explicit_changes: vec![],
                    linked_changes: vec![],
                    tax_incl_auto_fix: false,
                }
            } else {
                prepare(Job {
                    input: snapshot.input_params.clone(),
                    overrides: request.overrides,
                })?
            };
            let (current, conn) = runtime.require_context()?;
            if current.workspace_id != workspace.workspace_id
                || current.workspace_root != workspace.workspace_root
            {
                return Err("试算期间工作区发生变化，请重新试算".into());
            }
            bindings.authorize(
                &request.session_id,
                "simulate_benefit_calculation",
                None,
                &current.workspace_id,
                &cwd,
            )?;
            let service = ProjectService::new(Box::new(SqliteProjectRepository::new(conn)));
            let (_, current_scheme, latest) =
                calculation::resolve_snapshot(&service, &calc_request)?;
            if current_scheme.id != scheme.id
                || latest.id != snapshot.id
                || latest.version != snapshot.version
                || serde_json::to_value(&latest.input_params).map_err(|e| e.to_string())?
                    != serde_json::to_value(&snapshot.input_params).map_err(|e| e.to_string())?
            {
                return Err("试算期间已保存方案发生变化，请重新读取后试算".into());
            }
            let metrics = crate::benefit::calculator::calculate_ict_benefit(prepared.input)?;
            let mut result = calculation::build_response(&project, &scheme, &snapshot, &metrics);
            result.basis = "hypothetical".into();
            Ok(
                json!({"basis":"hypothetical","notice":"仅为假设试算；未保存、未生成方案、未回填任何科目。", "result":result,
                "explicitChanges":prepared.explicit_changes,"linkedChanges":prepared.linked_changes,"taxInclAutoFix":has_overrides.then_some(prepared.tax_incl_auto_fix)}),
            )
        })();
        match result {
            Ok(v) => BridgeReply::ok(v.to_string()),
            Err(e) => BridgeReply::error(422, &e),
        }
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn preparation_is_claimed_once_and_late_results_cannot_resolve_another_job() {
        let witness: Value =
            serde_json::from_str(include_str!("fixtures/benefit-simulation-desktop.json")).unwrap();
        let job = Job {
            input: serde_json::from_value(witness["cases"][0]["saved"].clone()).unwrap(),
            overrides: vec![],
        };
        let jobs = Arc::new(SimulationJobs::default());
        let worker_jobs = jobs.clone();
        let (sender, receiver) = mpsc::channel();
        let worker = std::thread::spawn(move || {
            worker_jobs.prepare(job, |id| {
                sender.send(id.to_string()).unwrap();
                Ok(())
            })
        });
        let id = receiver.recv().unwrap();
        assert!(jobs.claim(&id).is_ok());
        assert!(jobs.claim(&id).is_err());
        jobs.finish(&id, Err("明确失败，不计算".into())).unwrap();
        assert!(matches!(worker.join().unwrap(),Err(e) if e=="明确失败，不计算"));
        assert!(jobs.finish(&id, Err("迟到结果".into())).is_err());
        assert!(jobs.0.lock().unwrap().is_empty());
    }
}
