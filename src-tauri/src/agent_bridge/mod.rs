//! Agent bridge — runs lamber's business capabilities as deepseek-harness tools.
//!
//! ```text
//! React (AgentLabView)  --tauri invoke-->  agent_bridge commands
//!                                             |
//!                              ACP over stdio (dsh_session)
//!                                             v
//!                            dsh child process (--profile acp)
//!                                             |
//!                            dsh-tool-lamber tool body (HTTP)
//!                                             v
//!                       bridge_server --> benefit::calculator
//! ```
//!
//! The split is deliberate: dsh owns the agent loop and tool catalog; lamber
//! keeps every line of business math, and now owns the approval decision too.
//!
//! Two different seams, easily confused:
//!
//! * **Tool calls** run *into* lamber over the loopback bridge: the plugin's
//!   tool body POSTs to `bridge_server`, which dispatches to `benefit`.
//! * **Approval** runs *out of* lamber over ACP: dsh asks
//!   `session/requestPermission` on the same connection it streams output over.
//!
//! Approval used to travel over the bridge too, as a route the plugin posted
//! to. Under `--profile acp` that route is unreachable — `dsh-acp` answers the
//! `approval/request` Cordis event itself and forwards it to the client — so it
//! was removed rather than left as dead code. The bridge is back to carrying
//! read-only calculation and project permission checks.

pub mod approval;
pub mod approval_log;
pub mod bridge_server;
pub mod calculation;
mod benefit_access;
mod selection_fee;
pub(crate) mod benefit_simulation;
pub mod contract;
pub mod distribution;
pub mod dsh_session;
mod project_bindings;
mod project_query;
mod reviewed_arguments;
mod approval_rehearsal;
pub(crate) mod template_catalog;
mod template_write;
mod template_read;
pub(crate) mod template_images;
mod prompt;
mod session_store;
mod streaming;
pub mod tool_calls;
mod turns;

#[cfg(test)]
mod tests;

use approval::{ApprovalGate, ApprovalPrompt, APPROVAL_EVENT};
use bridge_server::{BridgeHandler, BridgeReply, BridgeServer};
use calculation::{CalculateRequest, CALCULATE_ROUTE};
use distribution::{locate as locate_distribution, prepare_user_home};
use dsh_session::{AcpRuntime, DshLaunchConfig};
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::sync::{Arc, Mutex};
use tauri::{Emitter, Manager};

use crate::config_manager::{
    AiAgentSettings, ConfigManager, DEFAULT_AI_AGENT_BASE_URL, DEFAULT_AI_AGENT_MODEL,
};

/// Frontend event carrying every ACP notification and turn outcome.
pub const SESSION_EVENT: &str = "ai://session-event";

fn open_session_store(app: &tauri::AppHandle) -> Result<session_store::SessionStore, String> {
    let directory = app.path().app_data_dir().map_err(|e| e.to_string())?;
    std::fs::create_dir_all(&directory).map_err(|e| format!("无法创建 AI 会话目录: {e}"))?;
    session_store::SessionStore::open(&directory.join("ai-sessions.sqlite"))
}

fn require_agent_workspace(
    runtime: &crate::workspace::WorkspaceRuntime,
) -> Result<crate::workspace::CurrentWorkspace, String> {
    runtime.require_workspace()
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct AiAgentModelOption {
    id: &'static str,
    name: &'static str,
    supports_images: bool,
}

const AI_AGENT_MODELS: [AiAgentModelOption; 3] = [
    AiAgentModelOption {
        id: "deepseek-v4-flash",
        name: "DeepSeek-V4-Flash",
        supports_images: false,
    },
    AiAgentModelOption {
        id: "deepseek-v4-pro",
        name: "DeepSeek-V4-Pro",
        supports_images: false,
    },
    AiAgentModelOption {
        id: "deepseek-v4-flash-vision-exp",
        name: "DeepSeek-V4-Flash-Vision-Exp",
        supports_images: true,
    },
];

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct AiAgentSettingsView {
    model: String,
    base_url: String,
    has_api_key: bool,
    models: Vec<AiAgentModelOption>,
}

impl From<AiAgentSettings> for AiAgentSettingsView {
    fn from(settings: AiAgentSettings) -> Self {
        Self {
            has_api_key: settings.api_key.as_ref().is_some_and(|key| !key.is_empty()),
            model: settings.model,
            base_url: settings.base_url,
            models: AI_AGENT_MODELS.to_vec(),
        }
    }
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct AiAgentSettingsUpdate {
    model: String,
    base_url: String,
    api_key: Option<String>,
    #[serde(default)]
    clear_api_key: bool,
}

/// Called with each approval question before lamber parks waiting for an answer.
///
/// Kept as a parameter rather than reaching for the `AppHandle` directly so
/// tests can drive the gate with a plain closure instead of a running app.
pub type ApprovalAnnouncer = Arc<dyn Fn(&ApprovalPrompt) + Send + Sync>;

/// Build the bridge route dispatcher for an open workspace.
///
/// Kept separate from the Tauri layer so tests can host the same routes over a
/// database they build themselves.
///
/// @param runtime - the workspace runtime holding the open database.
/// @returns a handler suitable for `BridgeServer::start`.
pub fn workspace_handler(
    runtime: Arc<crate::workspace::WorkspaceRuntime>,
    bindings: Arc<project_bindings::ProjectBindings>,
) -> BridgeHandler {
    Arc::new(move |path, body| {
        if path == selection_fee::FORWARD_ROUTE || path == selection_fee::REVERSE_ROUTE { return selection_fee::handle(&runtime, &bindings, path, body); }
        if path == benefit_access::READ_ROUTE { return benefit_access::handle_read(&runtime, &bindings, body); }
        if path == template_read::ROUTE { return template_read::handle(&runtime, &bindings, body); }
        if path != CALCULATE_ROUTE
            && path != project_bindings::AUTHORIZE_ROUTE
            && path != project_query::QUERY_ROUTE
        {
            return BridgeReply::error(404, "未知的 AI 桥接路由");
        }
        #[derive(Deserialize)]
        #[serde(rename_all = "camelCase")]
        struct Request {
            session_id: Option<String>,
            project_id: Option<String>,
            tool: Option<String>,
            scenario: Option<String>,
        }
        let request: Request = match serde_json::from_str(body) {
            Ok(value) => value,
            Err(_) => return BridgeReply::error(400, "工具请求格式错误"),
        };
        let session = match request
            .session_id
            .as_deref()
            .filter(|s| !s.trim().is_empty())
        {
            Some(id) => id,
            None => return BridgeReply::error(403, "缺少可信 AI 会话身份"),
        };
        let (workspace, conn) = match runtime.require_context() {
            Ok(value) => value,
            Err(e) => return BridgeReply::error(403, &e),
        };
        let cwd = match std::fs::canonicalize(&workspace.workspace_root) {
            Ok(value) => value,
            Err(_) => return BridgeReply::error(403, "工作区不可用"),
        };
        let tool = if path == CALCULATE_ROUTE {
            "run_benefit_calculation"
        } else if path == project_query::QUERY_ROUTE {
            project_query::QUERY_TOOL
        } else {
            request.tool.as_deref().unwrap_or("")
        };
        let scope = match bindings.authorize(
            session,
            tool,
            request.project_id.as_deref(),
            &workspace.workspace_id,
            &cwd,
        ) {
            Ok(id) => id,
            Err(e) => return BridgeReply::error(403, &e),
        };
        let service = crate::benefit::service::ProjectService::new(Box::new(
            crate::benefit::repository::SqliteProjectRepository::new(conn),
        ));
        if let Some(project) = scope.bound_project_id() {
            match service.get_project(project) {
                Ok(Some(_)) => {}
                _ => return BridgeReply::error(403, "绑定项目已不存在或不可访问，请新建会话"),
            }
        }
        if path == project_query::QUERY_ROUTE {
            let envelope: project_query::QueryEnvelope = match serde_json::from_str(body) {
                Ok(value) => value,
                Err(e) => return BridgeReply::error(400, &format!("查询参数错误: {e}")),
            };
            if envelope.session_id != session {
                return BridgeReply::error(403, "会话身份不一致");
            }
            return match project_query::query_projects(
                &service,
                &envelope.query,
                scope.bound_project_id(),
            ) {
                Ok(response) => match serde_json::to_string(&response) {
                    Ok(json) => BridgeReply::ok(json),
                    Err(e) => BridgeReply::error(500, &e.to_string()),
                },
                Err(e) => BridgeReply::error(422, &e),
            };
        }
        if path == project_bindings::AUTHORIZE_ROUTE {
            return BridgeReply::ok("{}".into());
        }
        let project = match scope {
            project_bindings::AuthorizedScope::Project(id) => id,
            _ => return BridgeReply::error(403, "该业务入口要求项目权限"),
        };
        match calculation::run_calculation(
            &service,
            &CalculateRequest {
                project_id: project,
                scenario: request.scenario,
            },
        ) {
            Ok(response) => match serde_json::to_string(&response) {
                Ok(json) => BridgeReply::ok(json),
                Err(e) => BridgeReply::error(500, &e.to_string()),
            },
            Err(e) => BridgeReply::error(422, &e),
        }
    })
}

/// The bridge server plus the dsh runtime it feeds, held for the app's lifetime.
///
/// Both are started lazily on the first prompt: launching a Node child process
/// at app boot would cost startup time for users who never open the AI panel.
#[derive(Default)]
pub struct AgentRuntime {
    inner: Mutex<Option<RunningAgent>>,
    /// Shared with the bridge handler; outlives individual dsh launches so a
    /// decision arriving during a restart cannot resolve into a dropped gate.
    gate: Arc<ApprovalGate>,
    simulations: Arc<benefit_simulation::SimulationJobs>,
}

struct RunningAgent {
    acp: AcpRuntime,
    config: DshLaunchConfig,
    /// lamber's own session ids mapped onto the ones the agent issued.
    ///
    /// ACP inverts session ownership: `session/new` returns an id the *agent*
    /// chose, where the SDK protocol accepted whatever id lamber invented. The
    /// frontend still names its own conversations, so each new name opens an ACP
    /// session once and reuses it for every later turn.
    sessions: HashMap<String, String>,
    turns: Arc<turns::Turns>,
    bindings: Arc<project_bindings::ProjectBindings>,
    /// Fields drop in declaration order: keep the bridge alive through ACP teardown.
    _bridge: BridgeServer,
}

impl AgentRuntime {
    /// Start the bridge and dsh if they are not already running, then queue one turn.
    ///
    /// Blocking: call it off the async executor (see `ai_send_prompt`). The turn
    /// itself is not awaited — output streams back as `SESSION_EVENT`s and the
    /// outcome arrives as a `session/turn-ended` event.
    ///
    /// @param app - handle used to emit session events to the frontend.
    /// @param runtime - workspace runtime backing the bridge routes.
    /// @param session_id - lamber's own conversation id; mapped to an ACP session.
    /// @param text - the user's prompt.
    /// @returns the ACP session id the turn was queued on.
    pub fn send_prompt(
        &self,
        app: &tauri::AppHandle,
        runtime: Arc<crate::workspace::WorkspaceRuntime>,
        session_id: &str,
        text: &str,
        request_id: Option<&str>,
        harness_session_id: Option<&str>,
        images: Vec<prompt::PromptImage>,
    ) -> Result<String, String> {
        let content = prompt::blocks(text, images)?;
        let mut guard = self
            .inner
            .lock()
            .map_err(|_| "AI 运行时锁已中毒".to_string())?;
        let workspace = require_agent_workspace(&runtime)?;
        let cwd = std::fs::canonicalize(&workspace.workspace_root)
            .map_err(|e| format!("无法访问工作区: {e}"))?;
        let store = open_session_store(app)?;
        let binding = store
            .binding(session_id)?
            .ok_or("请先选择项目或通用聊天；历史会话请新建后继续")?;
        if binding.cwd != cwd || binding.workspace_id != workspace.workspace_id {
            return Err("此 AI 会话属于另一个工作区，请打开原工作区或新建会话".into());
        }
        let saved = store.get(session_id, &cwd)?;
        if let Some(expected) = harness_session_id {
            if saved.as_deref() != Some(expected) {
                return Err("AI 会话映射与本地记录不一致，请保留记录并新建会话".into());
            }
        }
        if guard
            .as_ref()
            .is_some_and(|agent| !agent.acp.is_alive() || agent.config.cwd != cwd)
        {
            if guard
                .as_ref()
                .is_some_and(|agent| agent.acp.is_alive() && agent.turns.any_active())
            {
                return Err("原工作区仍有 AI 会话正在生成，请先停止生成".into());
            }
            *guard = None;
        }
        if guard.is_none() {
            *guard = Some(self.launch(app, runtime)?);
        }
        let agent = guard.as_mut().expect("just launched");
        let acp_session = match agent.sessions.get(session_id) {
            Some(existing) => existing.clone(),
            None => {
                let opened = match saved {
                    Some(id) => {
                        agent.acp.resume_session(&id, &cwd)?;
                        id
                    }
                    None => {
                        let id = agent.acp.new_session(&cwd)?;
                        store.insert(session_id, &id, &cwd)?;
                        id
                    }
                };
                agent
                    .sessions
                    .insert(session_id.to_string(), opened.clone());
                opened
            }
        };
        if agent.turns.is_active(session_id) {
            return Err("此会话正在生成，请先停止或等待结束".into());
        }
        agent.bindings.register(&acp_session, binding)?;
        agent
            .acp
            .set_model(&acp_session, &agent.config.provider, &agent.config.model)?;
        agent.turns.begin(&acp_session, session_id, request_id)?;
        if let Err(error) = agent.acp.prompt_blocks(&acp_session, content) {
            agent.turns.forget(&acp_session);
            return Err(error);
        }
        Ok(acp_session)
    }

    /// Tear the runtime down; the next prompt relaunches it.
    ///
    /// Open approvals are denied first, so a parked approval task is released
    /// immediately and its `session/requestPermission` is answered, instead of
    /// leaving dsh waiting out the rest of the gate's timeout.
    pub fn stop(&self) -> Result<(), String> {
        let denied = self.gate.shutdown();
        if denied > 0 {
            eprintln!("[agent_bridge] 关闭时拒绝了 {denied} 个未完成的审批请求");
        }
        let mut guard = self
            .inner
            .lock()
            .map_err(|_| "AI 运行时锁已中毒".to_string())?;
        *guard = None;
        // A later prompt relaunches the runtime, so the gate must accept again.
        self.gate.reopen();
        Ok(())
    }

    /// Deny open approvals without relaunching. Used on application exit.
    ///
    /// @returns how many open questions were denied.
    pub fn shutdown_approvals(&self) -> usize {
        self.gate.shutdown()
    }

    /// Deliver a user's decision to the parked approval task.
    pub fn resolve_approval(&self, request_id: &str, approved: bool, modified_args: Option<serde_json::Value>) -> Result<(), String> {
        self.gate.resolve_with_args(request_id, approved, modified_args)
    }

    fn launch(
        &self,
        app: &tauri::AppHandle,
        runtime: Arc<crate::workspace::WorkspaceRuntime>,
    ) -> Result<RunningAgent, String> {
        // This stays ahead of resource lookup and process launch: without a
        // user workspace there is no safe cwd, so repository/install fallback
        // is deliberately impossible.
        let workspace = require_agent_workspace(&runtime)?;
        let settings = ConfigManager::new(app).load().ai_agent;
        let distribution = locate_distribution(app)?;
        let (dsh_home, settings_patch) = prepare_user_home(app, &distribution, &settings)?;

        let announcer_app = app.clone();
        let announce: ApprovalAnnouncer = Arc::new(move |prompt: &ApprovalPrompt| {
            // Exactly one visible host: broadcasting would leave a stale duplicate
            // dialog in the main window after the floating chat settles a question.
            if let Some(window) = announcer_app
                .get_webview_window("ai-assistant")
                .or_else(|| announcer_app.get_webview_window("main"))
            {
                let _ = window.show();
                let _ = window.set_focus();
                let _ = announcer_app.emit_to(window.label(), APPROVAL_EVENT, prompt);
            }
        });
        // Persist every settled question for after-the-fact audit. With no
        // workspace open the recorder spools to disk and the next workspace
        // activation backfills it, so a decision is never silently lost.
        let spool = approval_log::spool_path(app)?;
        self.gate.set_recorder(approval_log::workspace_recorder(
            Arc::clone(&runtime),
            spool,
        ));
        let turns = Arc::new(turns::Turns::default());
        let bindings = Arc::new(project_bindings::ProjectBindings::default());
        let stream_app = app.clone();
        let approval_runtime = runtime.clone();
        let approval_bindings = bindings.clone();
        let changed_app = app.clone();
        let simulation = benefit_simulation::handler(runtime.clone(), bindings.clone(), benefit_simulation::app_preparer(app.clone(), self.simulations.clone()), workspace_handler(runtime.clone(), bindings.clone()));
        let writing = template_write::handler(runtime.clone(), bindings.clone(), self.gate.clone(),
            Arc::new(move |receipt| { let _ = changed_app.emit(template_write::CHANGED_EVENT, receipt); }),
            self.gate.reviewed.handler(simulation));
        let bridge = BridgeServer::start(streaming::handler(
            writing,
            Arc::clone(&turns),
            Arc::new(move |event| {
                let _ = stream_app.emit(SESSION_EVENT, event);
            }),
        ))?;

        let mut config = DshLaunchConfig {
            dsh_bin: distribution.node_bin,
            dsh_entry: Some(distribution.dsh_entry),
            profile: "acp".to_string(),
            patch_path: distribution.base_patch,
            extra_patch_path: Some(settings_patch),
            dsh_home,
            cwd: std::fs::canonicalize(&workspace.workspace_root).map_err(|e| e.to_string())?,
            provider: "deepseek-official".to_string(),
            model: settings.model.clone(),
            api_key: settings.api_key.or_else(|| {
                cfg!(debug_assertions)
                    .then(|| std::env::var("DEEPSEEK_API_KEY").ok())
                    .flatten()
                    .filter(|key| !key.is_empty())
            }),
            stream_display: true,
            bridge_url: String::new(),
            bridge_token: String::new(),
        };
        config.bridge_url = bridge.origin();
        config.bridge_token = bridge.token().to_string();

        let emitter = app.clone();
        let gate = Arc::clone(&self.gate);
        let event_turns = Arc::clone(&turns);
        let acp = AcpRuntime::start(
            &config,
            Arc::new(move |method, params| {
                event_turns.emit(method, params, |event| {
                    let _ = emitter.emit(SESSION_EVENT, event);
                });
            }),
            Arc::new(move |question| {
                template_write::request_approval(&gate, &approval_runtime, &approval_bindings, question, |prompt| announce(prompt))
            }),
        )?;

        Ok(RunningAgent {
            _bridge: bridge,
            acp,
            config,
            sessions: HashMap::new(),
            turns,
            bindings,
        })
    }
}

impl RunningAgent {
    /// What this runtime is, for diagnostics.
    ///
    /// Reports the negotiated protocol version and the agent's self-description
    /// alongside the route: under ACP those are the facts that say *which* peer
    /// lamber is actually talking to.
    fn describe(&self) -> serde_json::Value {
        let handshake = self.acp.handshake();
        serde_json::json!({
            "profile": self.config.profile,
            "provider": self.config.provider,
            "model": self.config.model,
            "hasApiKey": self.config.api_key.is_some(),
            "bridgeUrl": self.config.bridge_url,
            "protocolVersion": format!("{:?}", handshake.protocol_version),
            "agentName": handshake.agent_name,
            "agentVersion": handshake.agent_version,
            "supportsImagePrompts": handshake.supports_image_prompts,
            "openSessions": self.sessions.len(),
        })
    }
}

/// Send one prompt to the agent, starting the runtime on first use.
///
/// The work is pushed to a blocking thread: launching dsh and opening an ACP
/// session both park the caller, and doing that on Tauri's async executor would
/// tie up a worker that the connection's own tasks need.
#[tauri::command]
pub async fn ai_send_prompt(
    app: tauri::AppHandle,
    session_id: String,
    text: String,
    request_id: Option<String>,
    harness_session_id: Option<String>,
    images: Option<Vec<prompt::PromptImage>>,
) -> Result<String, String> {
    let runtime = app
        .state::<Arc<crate::workspace::WorkspaceRuntime>>()
        .inner()
        .clone();
    let agent = app.state::<Arc<AgentRuntime>>().inner().clone();
    let handle = app.clone();
    tauri::async_runtime::spawn_blocking(move || {
        agent.send_prompt(
            &handle,
            runtime,
            &session_id,
            &text,
            request_id.as_deref(),
            harness_session_id.as_deref(),
            images.unwrap_or_default(),
        )
    })
    .await
    .map_err(|e| format!("AI 请求执行失败: {e}"))?
}

/// Cancel only the named turn; a late stop can never cancel its successor.
#[tauri::command]
pub async fn ai_cancel_prompt(
    app: tauri::AppHandle,
    session_id: String,
    request_id: String,
) -> Result<(), String> {
    let agent = app.state::<Arc<AgentRuntime>>().inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let guard = agent.inner.lock().map_err(|_| "AI 运行时锁已中毒")?;
        if let Some(running) = guard.as_ref() {
            if let Some(acp) = running.turns.matches(&session_id, &request_id) {
                running.acp.cancel(&acp)?;
            }
        }
        Ok(())
    })
    .await
    .map_err(|e| format!("停止 AI 请求失败: {e}"))?
}

/// Only explicit UI actions establish immutable permission; never called by a tool.
#[tauri::command]
pub async fn ai_bind_session_to_project(
    app: tauri::AppHandle,
    session_id: String,
    project_id: Option<String>,
) -> Result<project_bindings::ProjectBinding, String> {
    let agent = app.state::<Arc<AgentRuntime>>().inner().clone();
    let runtime = app
        .state::<Arc<crate::workspace::WorkspaceRuntime>>()
        .inner()
        .clone();
    tauri::async_runtime::spawn_blocking(move || {
        let _guard = agent.inner.lock().map_err(|_| "AI 运行时锁不可用")?;
        let (workspace, conn) = runtime.require_context()?;
        let binding = project_bindings::ProjectBinding {
            workspace_id: workspace.workspace_id,
            cwd: std::fs::canonicalize(workspace.workspace_root).map_err(|e| e.to_string())?,
            project_id,
        };
        if let Some(id) = &binding.project_id {
            let service = crate::benefit::service::ProjectService::new(Box::new(
                crate::benefit::repository::SqliteProjectRepository::new(conn),
            ));
            if service.get_project(id)?.is_none() {
                return Err("所选项目不存在".into());
            }
        }
        open_session_store(&app)?.bind(&session_id, &binding)?;
        Ok(binding)
    })
    .await
    .map_err(|e| e.to_string())?
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SessionBindingView {
    #[serde(flatten)]
    binding: project_bindings::ProjectBinding,
    project_name: Option<String>,
}

#[tauri::command]
pub async fn ai_get_session_binding(
    app: tauri::AppHandle,
    session_id: String,
) -> Result<Option<SessionBindingView>, String> {
    let Some(binding) = open_session_store(&app)?.binding(&session_id)? else {
        return Ok(None);
    };
    let (workspace, conn) = app
        .state::<Arc<crate::workspace::WorkspaceRuntime>>()
        .require_context()?;
    let cwd = std::fs::canonicalize(workspace.workspace_root).map_err(|e| e.to_string())?;
    if workspace.workspace_id != binding.workspace_id || cwd != binding.cwd {
        return Err("此会话属于另一个工作区".into());
    }
    let project_name = if let Some(id) = &binding.project_id {
        let service = crate::benefit::service::ProjectService::new(Box::new(
            crate::benefit::repository::SqliteProjectRepository::new(conn),
        ));
        Some(
            service
                .get_project(id)?
                .ok_or("绑定项目已不存在，请新建会话")?
                .name,
        )
    } else {
        None
    };
    Ok(Some(SessionBindingView {
        binding,
        project_name,
    }))
}

/// Forget a mapping only after a user's clear/delete action and a settled turn.
#[tauri::command]
pub async fn ai_reset_session(app: tauri::AppHandle, session_id: String) -> Result<(), String> {
    let agent = app.state::<Arc<AgentRuntime>>().inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut guard = agent.inner.lock().map_err(|_| "AI 运行时锁已中毒")?;
        if let Some(running) = guard.as_mut() {
            if running.turns.is_active(&session_id) {
                return Err("请先停止此会话的生成".into());
            }
            if let Some(id) = running.sessions.get(&session_id) {
                if running.acp.is_alive() {
                    running.acp.close_session(id)?;
                }
            }
            if let Some(id) = running.sessions.get(&session_id) {
                running.bindings.forget(id)?;
            }
            running.sessions.remove(&session_id);
        }
        open_session_store(&app)?.remove(&session_id)
    })
    .await
    .map_err(|e| format!("清除 AI 会话失败: {e}"))?
}

/// Report whether the agent runtime is up, and on what route.
#[tauri::command]
pub async fn ai_agent_status(app: tauri::AppHandle) -> Result<serde_json::Value, String> {
    let agent = app.state::<Arc<AgentRuntime>>().inner().clone();
    let guard = agent
        .inner
        .lock()
        .map_err(|_| "AI 运行时锁已中毒".to_string())?;
    Ok(match guard.as_ref() {
        Some(running) => {
            let mut value = running.describe();
            value["running"] = serde_json::Value::Bool(true);
            value
        }
        None => serde_json::json!({ "running": false }),
    })
}

/// Stop the dsh child and release the bridge port.
#[tauri::command]
pub async fn ai_agent_stop(app: tauri::AppHandle) -> Result<(), String> {
    app.state::<Arc<AgentRuntime>>().inner().clone().stop()
}

/// Read non-secret AI settings. The stored key never crosses into the webview;
/// callers only learn whether one exists and may replace or clear it explicitly.
#[tauri::command]
pub async fn ai_get_settings(app: tauri::AppHandle) -> Result<AiAgentSettingsView, String> {
    Ok(ConfigManager::new(&app).load().ai_agent.into())
}

/// Persist AI settings and stop the current child so every changed launch fact
/// takes effect on the next prompt.
#[tauri::command]
pub async fn ai_save_settings(
    app: tauri::AppHandle,
    update: AiAgentSettingsUpdate,
) -> Result<AiAgentSettingsView, String> {
    let model = update.model.trim();
    if !AI_AGENT_MODELS
        .iter()
        .any(|candidate| candidate.id == model)
    {
        return Err(format!("不支持的 AI 模型：{model}"));
    }
    let mut base_url = update.base_url.trim().trim_end_matches('/').to_string();
    if base_url.is_empty() {
        base_url = DEFAULT_AI_AGENT_BASE_URL.to_string();
    }
    if !(base_url.starts_with("https://") || base_url.starts_with("http://")) {
        return Err("AI 服务地址必须以 http:// 或 https:// 开头".to_string());
    }

    let manager = ConfigManager::new(&app);
    let mut config = manager.load();
    let api_key = if update.clear_api_key {
        None
    } else {
        update
            .api_key
            .map(|key| key.trim().to_string())
            .filter(|key| !key.is_empty())
            .or(config.ai_agent.api_key)
    };
    config.ai_agent = AiAgentSettings {
        api_key,
        model: if model.is_empty() {
            DEFAULT_AI_AGENT_MODEL.to_string()
        } else {
            model.to_string()
        },
        base_url,
    };
    manager.save(&config)?;

    app.state::<Arc<AgentRuntime>>().inner().clone().stop()?;
    Ok(config.ai_agent.into())
}

/// Read the most recent approval audit entries, newest first.
///
/// @param limit - maximum rows; defaults to 50 and is capped at 500.
#[tauri::command]
pub async fn ai_list_approval_log(
    app: tauri::AppHandle,
    limit: Option<u32>,
) -> Result<Vec<approval_log::ApprovalLogEntry>, String> {
    let runtime = app
        .state::<Arc<crate::workspace::WorkspaceRuntime>>()
        .inner()
        .clone();
    let conn = runtime.require_db()?;
    approval_log::recent(&conn, limit.unwrap_or(50).min(500))
}

/// Deliver the user's answer to a pending `ai://approval-request`.
///
/// @param request_id - the `requestId` from the emitted prompt.
/// @param approved - `true` grants this one call; anything else denies it.
#[tauri::command]
pub async fn ai_resolve_approval(
    app: tauri::AppHandle,
    request_id: String,
    approved: bool,
    modified_args: Option<serde_json::Value>,
) -> Result<(), String> {
    app.state::<Arc<AgentRuntime>>()
        .inner()
        .clone()
        .resolve_approval(&request_id, approved, modified_args)?;
    let _ = app.emit("ai://approval-settled", serde_json::json!({"requestId":request_id}));
    Ok(())
}

/// Explicit diagnostic action; writes only an app-owned rehearsal file, never project data.
#[tauri::command]
pub async fn ai_rehearse_text_approval(
    app: tauri::AppHandle,
    window: tauri::WebviewWindow,
    proposed_text: String,
) -> Result<approval_rehearsal::RehearsalReceipt, String> {
    let agent = app.state::<Arc<AgentRuntime>>().inner().clone();
    let runtime = app.state::<Arc<crate::workspace::WorkspaceRuntime>>().inner().clone();
    // Use the same durable recorder and gate as real ACP approvals.
    agent.gate.set_recorder(approval_log::workspace_recorder(runtime, approval_log::spool_path(&app)?));
    let path = app.path().app_data_dir().map_err(|e| e.to_string())?.join("approval-rehearsal.txt");
    tauri::async_runtime::spawn_blocking(move || {
        approval_rehearsal::run(&agent.gate, &path, proposed_text, |prompt| {
            let _ = window.emit(APPROVAL_EVENT, prompt);
        })
    }).await.map_err(|e| e.to_string())?
}
