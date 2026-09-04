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
//! exactly one read-only route.

pub mod approval;
pub mod approval_log;
pub mod bridge_server;
pub mod calculation;
pub mod dsh_session;
pub mod tool_calls;

#[cfg(test)]
mod tests;

use approval::{ApprovalGate, ApprovalPrompt, APPROVAL_EVENT};
use bridge_server::{BridgeHandler, BridgeReply, BridgeServer};
use calculation::{CalculateRequest, CALCULATE_ROUTE};
use dsh_session::{AcpRuntime, DshLaunchConfig};
use std::collections::HashMap;
use std::sync::{Arc, Mutex};
use tauri::{Emitter, Manager};

/// Frontend event carrying every ACP notification and turn outcome.
pub const SESSION_EVENT: &str = "ai://session-event";

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
pub fn workspace_handler(runtime: Arc<crate::workspace::WorkspaceRuntime>) -> BridgeHandler {
    Arc::new(move |path, body| match path {
        CALCULATE_ROUTE => {
            let request: CalculateRequest = match serde_json::from_str(body) {
                Ok(request) => request,
                Err(e) => return BridgeReply::error(400, &format!("请求体解析失败: {e}")),
            };
            match calculate_with_runtime(&runtime, &request) {
                Ok(response) => match serde_json::to_string(&response) {
                    Ok(json) => BridgeReply::ok(json),
                    Err(e) => BridgeReply::error(500, &format!("结果序列化失败: {e}")),
                },
                Err(message) => BridgeReply::error(422, &message),
            }
        }
        other => BridgeReply::error(404, &format!("未知的 AI 桥接路由: {other}")),
    })
}

fn calculate_with_runtime(
    runtime: &crate::workspace::WorkspaceRuntime,
    request: &CalculateRequest,
) -> Result<calculation::CalculateResponse, String> {
    let conn = runtime.require_db()?;
    let service = crate::benefit::service::ProjectService::new(Box::new(
        crate::benefit::repository::SqliteProjectRepository::new(conn),
    ));
    calculation::run_calculation(&service, request)
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
}

struct RunningAgent {
    /// Dropped last; keeps the loopback listener alive while dsh may call it.
    _bridge: BridgeServer,
    acp: AcpRuntime,
    config: DshLaunchConfig,
    /// lamber's own session ids mapped onto the ones the agent issued.
    ///
    /// ACP inverts session ownership: `session/new` returns an id the *agent*
    /// chose, where the SDK protocol accepted whatever id lamber invented. The
    /// frontend still names its own conversations, so each new name opens an ACP
    /// session once and reuses it for every later turn.
    sessions: HashMap<String, String>,
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
    ) -> Result<String, String> {
        let mut guard = self
            .inner
            .lock()
            .map_err(|_| "AI 运行时锁已中毒".to_string())?;
        if guard.is_none() {
            *guard = Some(self.launch(app, runtime)?);
        }
        let agent = guard.as_mut().expect("just launched");
        let acp_session = match agent.sessions.get(session_id) {
            Some(existing) => existing.clone(),
            None => {
                let opened = agent.acp.new_session(&agent.config.cwd)?;
                agent
                    .sessions
                    .insert(session_id.to_string(), opened.clone());
                opened
            }
        };
        agent.acp.prompt(&acp_session, text)?;
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
    pub fn resolve_approval(&self, request_id: &str, approved: bool) -> Result<(), String> {
        self.gate.resolve(request_id, approved)
    }

    fn launch(
        &self,
        app: &tauri::AppHandle,
        runtime: Arc<crate::workspace::WorkspaceRuntime>,
    ) -> Result<RunningAgent, String> {
        let announcer_app = app.clone();
        let announce: ApprovalAnnouncer = Arc::new(move |prompt: &ApprovalPrompt| {
            let _ = announcer_app.emit(APPROVAL_EVENT, prompt);
        });
        // Persist every settled question for after-the-fact audit. With no
        // workspace open the recorder spools to disk and the next workspace
        // activation backfills it, so a decision is never silently lost.
        let spool = approval_log::spool_path(app)?;
        self.gate.set_recorder(approval_log::workspace_recorder(
            Arc::clone(&runtime),
            spool,
        ));
        let bridge = BridgeServer::start(workspace_handler(runtime))?;

        let mut config = DshLaunchConfig::from_repo_root(&repo_root()?);
        config.bridge_url = bridge.origin();
        config.bridge_token = bridge.token().to_string();

        let emitter = app.clone();
        let gate = Arc::clone(&self.gate);
        let acp = AcpRuntime::start(
            &config,
            Arc::new(move |method, params| {
                // The frontend subscribes to one channel and switches on `method`,
                // so a new notification kind never needs a new Tauri event name.
                let _ = emitter.emit(
                    SESSION_EVENT,
                    serde_json::json!({ "method": method, "params": params }),
                );
            }),
            Arc::new(move |question| {
                approval::handle_request(&gate, question, |prompt| announce(prompt))
            }),
        )?;

        Ok(RunningAgent {
            _bridge: bridge,
            acp,
            config,
            sessions: HashMap::new(),
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
            "openSessions": self.sessions.len(),
        })
    }
}

/// Locate the lamber repository root that hosts `agent-bridge/`.
///
/// In a dev run the binary sits under `src-tauri/target/<profile>/`; walking up
/// to the directory containing `agent-bridge` keeps the lookup independent of
/// the build profile.
fn repo_root() -> Result<std::path::PathBuf, String> {
    if let Ok(explicit) = std::env::var("LAMBER_REPO_ROOT") {
        if !explicit.is_empty() {
            return Ok(std::path::PathBuf::from(explicit));
        }
    }
    let exe = std::env::current_exe().map_err(|e| format!("无法定位可执行文件: {e}"))?;
    for ancestor in exe.ancestors() {
        if ancestor.join("agent-bridge").join("patch.yml").is_file() {
            return Ok(ancestor.to_path_buf());
        }
    }
    Err("未找到 agent-bridge 目录；请设置 LAMBER_REPO_ROOT 环境变量".to_string())
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
) -> Result<String, String> {
    let runtime = app
        .state::<Arc<crate::workspace::WorkspaceRuntime>>()
        .inner()
        .clone();
    let agent = app.state::<Arc<AgentRuntime>>().inner().clone();
    let handle = app.clone();
    tauri::async_runtime::spawn_blocking(move || {
        agent.send_prompt(&handle, runtime, &session_id, &text)
    })
    .await
    .map_err(|e| format!("AI 请求执行失败: {e}"))?
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
) -> Result<(), String> {
    app.state::<Arc<AgentRuntime>>()
        .inner()
        .clone()
        .resolve_approval(&request_id, approved)
}
