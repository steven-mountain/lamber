//! Supervises the `dsh` agent runtime child process and speaks ACP to it.
//!
//! The wire protocol is the [Agent Client Protocol](https://agentclientprotocol.com)
//! over the child's stdio, served by `dsh --profile acp`. Unlike the SDK
//! protocol this replaced, ACP is **bidirectional**: the agent asks the client
//! questions too, and `session/requestPermission` is the one that matters here.
//! A request/response client keyed on our own ids cannot answer those, so the
//! transport is the official `agent-client-protocol` crate rather than the
//! hand-written JSON-RPC client it replaced.
//!
//! ```text
//!  sync lamber backend                       tokio, one thread, owned here
//!  ───────────────────                       ────────────────────────────
//!  AcpRuntime::prompt() ──Command──▶ command loop ──session/prompt──▶ dsh
//!            ▲                             │
//!            └───────std::mpsc reply───────┘
//!
//!  dsh ──session/update──────────▶ notification handler ──▶ SessionUpdateSink
//!  dsh ──session/requestPermission──▶ request handler ────▶ PermissionResponder
//! ```
//!
//! tokio lives here and nowhere else. The rest of lamber's backend stays
//! synchronous: callers block on a plain `std::sync::mpsc` reply while the
//! connection runs on this module's own thread.

use std::path::{Path, PathBuf};
use std::sync::mpsc as sync_mpsc;
use std::sync::Arc;
use std::time::Duration;

use agent_client_protocol::schema::v1::{
    CancelNotification, CloseSessionRequest, ContentBlock, InitializeRequest, InitializeResponse,
    NewSessionRequest, PermissionOptionKind, PromptRequest, RequestPermissionOutcome,
    RequestPermissionRequest, RequestPermissionResponse, ResumeSessionRequest,
    SelectedPermissionOutcome, SessionId, SessionNotification, SessionUpdate,
    SetSessionConfigOptionRequest, TextContent,
};
use agent_client_protocol::schema::ProtocolVersion;
use agent_client_protocol::{AcpAgent, AcpAgentConfig, Agent, ConnectionTo};

use super::approval::{ApprovalDecision, ApprovalQuestion};
use super::bridge_server::BRIDGE_TOKEN_HEADER;
use super::tool_calls::{ToolCallIndex, TrackedCall};

/// The only ACP wire version lamber speaks.
///
/// Asserted explicitly on every handshake, because dsh does not negotiate:
/// `initialize(_params)` ignores what the client asked for and answers a
/// hard-coded `1` (`dsh-acp/lib/index.js:1143-1146`). Nothing on the wire would
/// tell us the day dsh moves to v2 — the mismatch would instead surface much
/// later as an unparseable field, on whichever message happened to change. This
/// constant turns that into a refusal at startup.
pub const EXPECTED_PROTOCOL_VERSION: ProtocolVersion = ProtocolVersion::V1;

/// How long a caller waits for the connection thread to answer a command.
const COMMAND_TIMEOUT: Duration = Duration::from_secs(60);

/// Event method carrying one ACP `session/update` to the frontend.
pub const UPDATE_METHOD: &str = "session/update";

/// Event method marking the end of one turn, successful or not.
pub const TURN_ENDED_METHOD: &str = "session/turn-ended";

/// Everything needed to launch one dsh runtime bound to this lamber instance.
///
/// `Debug` is hand-written to redact the API key and the bridge token: a config
/// dump in a panic message, a log line, or a bug report must never carry a
/// live credential.
#[derive(Clone)]
pub struct DshLaunchConfig {
    /// Node executable used to run dsh. A bare `node` is allowed for development.
    pub dsh_bin: PathBuf,
    /// JavaScript CLI entry passed as Node's first argument.
    pub dsh_entry: Option<PathBuf>,
    /// Profile to boot; the ACP server lives in the `acp` profile.
    pub profile: String,
    /// Absolute path to `agent-bridge/patch.yml`, which mounts `dsh-tool-lamber`.
    pub patch_path: PathBuf,
    /// Per-user model/baseURL overlay, applied after the immutable base patch.
    pub extra_patch_path: Option<PathBuf>,
    /// Writable `$DSH_HOME` holding profiles, sessions, and credentials.
    pub dsh_home: PathBuf,
    /// Working directory sent as `session/new`'s `cwd`.
    pub cwd: PathBuf,
    /// Provider route registered in the profile, e.g. `deepseek-official`.
    pub provider: String,
    /// Model on that route, e.g. `deepseek-v4-flash`.
    pub model: String,
    /// DeepSeek API key. Absent means the runtime boots but every turn fails at the LLM call.
    pub api_key: Option<String>,
    /// Origin of this instance's bridge server.
    pub stream_display: bool,
    pub bridge_url: String,
    /// Per-launch bridge token the plugin must present.
    pub bridge_token: String,
}

impl std::fmt::Debug for DshLaunchConfig {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("DshLaunchConfig")
            .field("dsh_bin", &self.dsh_bin)
            .field("dsh_entry", &self.dsh_entry)
            .field("profile", &self.profile)
            .field("patch_path", &self.patch_path)
            .field("extra_patch_path", &self.extra_patch_path)
            .field("dsh_home", &self.dsh_home)
            .field("cwd", &self.cwd)
            .field("provider", &self.provider)
            .field("model", &self.model)
            .field("api_key", &redacted(self.api_key.as_deref()))
            .field("bridge_url", &self.bridge_url)
            .field("bridge_token", &redacted(Some(self.bridge_token.as_str())))
            .finish()
    }
}

/// Render a secret as its presence, never its value.
fn redacted(secret: Option<&str>) -> &'static str {
    match secret {
        Some(value) if !value.is_empty() => "<redacted>",
        _ => "<unset>",
    }
}

impl DshLaunchConfig {
    /// Default launch configuration derived from the repository layout.
    ///
    /// @param repo_root - the lamber repository root containing `agent-bridge/`.
    /// @returns a config pointing at the provisioned local dsh install.
    pub fn from_repo_root(repo_root: &Path) -> Self {
        let agent_bridge = repo_root.join("agent-bridge");
        Self {
            dsh_bin: std::env::var_os("LAMBER_NODE_BIN")
                .map(PathBuf::from)
                .unwrap_or_else(|| PathBuf::from("node")),
            dsh_entry: Some(agent_bridge.join("node_modules/@deepseek-ai/dsh/lib/bin.js")),
            profile: "acp".to_string(),
            patch_path: agent_bridge.join("patch.yml"),
            extra_patch_path: None,
            dsh_home: agent_bridge.join(".dsh-home"),
            cwd: repo_root.to_path_buf(),
            provider: "deepseek-official".to_string(),
            model: "deepseek-v4-flash".to_string(),
            api_key: std::env::var("DEEPSEEK_API_KEY")
                .ok()
                .filter(|k| !k.is_empty()),
            stream_display: false,
            bridge_url: String::new(),
            bridge_token: String::new(),
        }
    }

    /// Describe the child process to launch.
    fn to_agent_config(&self) -> AcpAgentConfig {
        let mut config = AcpAgentConfig::new(&self.dsh_bin);
        if let Some(entry) = &self.dsh_entry {
            config = config.arg(entry.to_string_lossy().to_string());
        }
        config = config
            .arg("--profile")
            .arg(&self.profile)
            .arg("--patch")
            .arg(self.patch_path.to_string_lossy().to_string())
            .env("DSH_HOME", self.dsh_home.to_string_lossy().to_string())
            // dsh ships telemetry to an external host by default; lamber handles
            // customer financial data, so it stays off.
            .env("DSH_TELEMETRY_MODE", "DISABLED")
            .env("LAMBER_STREAM_DISPLAY", if self.stream_display { "1" } else { "0" })
            .env("LAMBER_BRIDGE_URL", &self.bridge_url)
            .env("LAMBER_BRIDGE_TOKEN", &self.bridge_token)
            .env("LAMBER_BRIDGE_TOKEN_HEADER", BRIDGE_TOKEN_HEADER);
        if let Some(extra_patch) = &self.extra_patch_path {
            config = config
                .arg("--patch")
                .arg(extra_patch.to_string_lossy().to_string());
        }

        // The child inherits lamber's environment and `AcpAgentConfig` offers no
        // way to unset a variable, so an absent key is passed as an empty one.
        // That is not a workaround: dsh's credential layer treats an empty value
        // as unset (`dsh-credentials-local/lib/index.js:427`), so this is exactly
        // the old `env_remove`. Letting an ambient key leak in instead would
        // silently bill an account the caller did not choose.
        config.env("DEEPSEEK_API_KEY", self.api_key.clone().unwrap_or_default())
    }
}

/// Called on the connection thread for every agent-to-client notification.
///
/// Receives `(method, params)` so the frontend keeps subscribing to one channel
/// and switching on the method, exactly as it did under the SDK protocol.
pub type SessionUpdateSink = Arc<dyn Fn(&str, &serde_json::Value) + Send + Sync>;

/// Answers one `session/requestPermission`. Blocks until a human decides.
pub type PermissionResponder = Arc<dyn Fn(ApprovalQuestion) -> ApprovalDecision + Send + Sync>;

/// What the agent reported about itself during the handshake.
#[derive(Debug, Clone)]
pub struct AgentHandshake {
    pub protocol_version: ProtocolVersion,
    pub agent_name: Option<String>,
    pub agent_version: Option<String>,
    /// Whether the selected model route can accept ACP inline image blocks.
    pub supports_image_prompts: bool,
}

impl AgentHandshake {
    fn from_response(response: &InitializeResponse) -> Self {
        Self {
            protocol_version: response.protocol_version,
            agent_name: response.agent_info.as_ref().map(|info| info.name.clone()),
            agent_version: response
                .agent_info
                .as_ref()
                .map(|info| info.version.clone()),
            supports_image_prompts: response.agent_capabilities.prompt_capabilities.image,
        }
    }
}

/// One instruction for the connection thread.
enum Command {
    Close {
        session_id: String,
        reply: sync_mpsc::Sender<Result<(), String>>,
    },
    SetModel {
        session_id: String,
        value: String,
        reply: sync_mpsc::Sender<Result<(), String>>,
    },
    Resume {
        session_id: String,
        cwd: PathBuf,
        reply: sync_mpsc::Sender<Result<(), String>>,
    },
    Cancel {
        session_id: String,
        reply: sync_mpsc::Sender<Result<(), String>>,
    },
    NewSession {
        cwd: PathBuf,
        reply: sync_mpsc::Sender<Result<String, String>>,
    },
    Prompt {
        session_id: String,
        content: Vec<ContentBlock>,
        reply: sync_mpsc::Sender<Result<(), String>>,
    },
}

/// A live dsh runtime: the child process, its ACP connection, and the thread they run on.
pub struct AcpRuntime {
    commands: tokio::sync::mpsc::UnboundedSender<Command>,
    handshake: AgentHandshake,
    thread: Option<std::thread::JoinHandle<()>>,
}

impl AcpRuntime {
    /// Launch `dsh --profile acp`, complete the handshake, and start serving.
    ///
    /// Returns only once `initialize` has succeeded and its protocol version has
    /// been checked, so a caller that gets an `AcpRuntime` back has a connection
    /// known to speak a version this client understands.
    ///
    /// @param config - launch parameters, including the bridge coordinates handed to the plugin.
    /// @param on_update - called with `(method, params)` for every notification.
    /// @param on_permission - answers permission requests; blocks until a human decides.
    /// @returns the running runtime, past its handshake.
    pub fn start(
        config: &DshLaunchConfig,
        on_update: SessionUpdateSink,
        on_permission: PermissionResponder,
    ) -> Result<Self, String> {
        if config.dsh_bin.components().count() > 1 && !config.dsh_bin.is_file() {
            return Err(format!(
                "AI 运行组件缺少 Node 可执行文件：{}。请重新安装 Lamber。",
                config.dsh_bin.display()
            ));
        }
        if let Some(entry) = &config.dsh_entry {
            if !entry.is_file() {
                return Err(format!(
                    "AI 运行组件缺少 dsh 入口：{}。请重新安装 Lamber。",
                    entry.display()
                ));
            }
        }
        if !config.patch_path.is_file() {
            return Err(format!(
                "AI 运行组件缺少基础配置：{}。请重新安装 Lamber。",
                config.patch_path.display()
            ));
        }
        if let Some(extra_patch) = &config.extra_patch_path {
            if !extra_patch.is_file() {
                return Err(format!("AI 用户配置不存在：{}", extra_patch.display()));
            }
        }

        let agent_config = config.to_agent_config();
        let (commands, command_rx) = tokio::sync::mpsc::unbounded_channel::<Command>();
        // The handshake result travels back to this (synchronous) caller, so it
        // can fail `start` rather than letting a bad version surface later.
        let (ready_tx, ready_rx) = sync_mpsc::channel::<Result<AgentHandshake, String>>();

        let thread = std::thread::Builder::new()
            .name("lamber-dsh-acp".into())
            .spawn(move || {
                run_connection(agent_config, command_rx, ready_tx, on_update, on_permission)
            })
            .map_err(|e| format!("无法启动 dsh 连接线程: {e}"))?;

        // A dead thread drops `ready_tx`, so this errors instead of hanging.
        let handshake = match ready_rx.recv() {
            Ok(result) => result,
            Err(_) => Err("dsh 连接线程在握手完成前退出".to_string()),
        };

        match handshake {
            Ok(handshake) => Ok(Self {
                commands,
                handshake,
                thread: Some(thread),
            }),
            Err(message) => {
                drop(commands);
                let _ = thread.join();
                Err(message)
            }
        }
    }

    /// What the agent reported during `initialize`.
    pub fn is_alive(&self) -> bool {
        !self.commands.is_closed()
    }

    pub fn handshake(&self) -> &AgentHandshake {
        &self.handshake
    }

    /// Open one ACP session and return the agent-issued session id.
    ///
    /// ACP sessions are named by the agent, not the caller — the opposite of the
    /// SDK protocol, where lamber invented the id. Callers that track their own
    /// session identity must map it onto what this returns.
    ///
    /// @param cwd - working directory the session is rooted at.
    pub fn new_session(&self, cwd: &Path) -> Result<String, String> {
        let (reply, answer) = sync_mpsc::channel();
        self.dispatch(Command::NewSession {
            cwd: cwd.to_path_buf(),
            reply,
        })?;
        Self::collect(answer, "session/new")
    }

    /// Queue one user turn; output streams back through the update sink.
    ///
    /// Returns as soon as the request is on the wire. ACP's `session/prompt`
    /// only answers when the whole turn has finished, so waiting for it here
    /// would hold the caller for the length of a model turn and stall the UI
    /// mid-generation. The turn's outcome arrives as a `TURN_ENDED_METHOD`
    /// event instead.
    ///
    /// @param session_id - an agent-issued id from `new_session`.
    /// @param text - the user's prompt text.
    pub fn prompt(&self, session_id: &str, text: &str) -> Result<(), String> {
        self.prompt_blocks(session_id, vec![ContentBlock::Text(TextContent::new(text))])
    }

    pub fn prompt_blocks(
        &self,
        session_id: &str,
        content: Vec<ContentBlock>,
    ) -> Result<(), String> {
        if content
            .iter()
            .any(|block| matches!(block, ContentBlock::Image(_)))
            && !self.handshake.supports_image_prompts
        {
            return Err("当前模型不支持图片，请在 AI 模型与服务中选择视觉模型".into());
        }
        let (reply, answer) = sync_mpsc::channel();
        self.dispatch(Command::Prompt {
            session_id: session_id.into(),
            content,
            reply,
        })?;
        Self::collect(answer, "session/prompt")
    }

    pub fn close_session(&self, session_id: &str) -> Result<(), String> {
        let (reply, answer) = sync_mpsc::channel();
        self.dispatch(Command::Close {
            session_id: session_id.into(),
            reply,
        })?;
        Self::collect(answer, "session/close")
    }

    pub fn set_model(&self, session_id: &str, provider: &str, model: &str) -> Result<(), String> {
        let (reply, answer) = sync_mpsc::channel();
        self.dispatch(Command::SetModel {
            session_id: session_id.into(),
            value: serde_json::json!([provider, model]).to_string(),
            reply,
        })?;
        Self::collect(answer, "session/set_config_option")
    }

    pub fn resume_session(&self, session_id: &str, cwd: &Path) -> Result<(), String> {
        let (reply, answer) = sync_mpsc::channel();
        self.dispatch(Command::Resume {
            session_id: session_id.into(),
            cwd: cwd.into(),
            reply,
        })?;
        Self::collect(answer, "session/resume")
    }

    pub fn cancel(&self, session_id: &str) -> Result<(), String> {
        let (reply, answer) = sync_mpsc::channel();
        self.dispatch(Command::Cancel {
            session_id: session_id.into(),
            reply,
        })?;
        Self::collect(answer, "session/cancel")
    }

    fn dispatch(&self, command: Command) -> Result<(), String> {
        self.commands
            .send(command)
            .map_err(|_| "dsh 连接已关闭".to_string())
    }

    /// Block on one command's reply, turning a dead connection into an error.
    ///
    /// The command's own failure and the transport's are both errors to the
    /// caller, so they are flattened into one `Result` here rather than making
    /// every call site unwrap two layers.
    fn collect<T>(
        answer: sync_mpsc::Receiver<Result<T, String>>,
        label: &str,
    ) -> Result<T, String> {
        match answer.recv_timeout(COMMAND_TIMEOUT) {
            Ok(result) => result,
            Err(sync_mpsc::RecvTimeoutError::Timeout) => {
                Err(format!("等待 dsh 响应超时（{label}）"))
            }
            Err(sync_mpsc::RecvTimeoutError::Disconnected) => {
                Err(format!("dsh 连接已断开（{label}）"))
            }
        }
    }
}

impl Drop for AcpRuntime {
    fn drop(&mut self) {
        // Closing the command channel ends the connection loop, which drops the
        // transport and, with it, the child's process group.
        let (dead, _) = tokio::sync::mpsc::unbounded_channel();
        let commands = std::mem::replace(&mut self.commands, dead);
        drop(commands);
        if let Some(thread) = self.thread.take() {
            let _ = thread.join();
        }
    }
}

/// Own the tokio runtime and drive the ACP connection until commands stop.
fn run_connection(
    agent_config: AcpAgentConfig,
    command_rx: tokio::sync::mpsc::UnboundedReceiver<Command>,
    ready_tx: sync_mpsc::Sender<Result<AgentHandshake, String>>,
    on_update: SessionUpdateSink,
    on_permission: PermissionResponder,
) {
    // Multi-threaded on purpose: the permission handler parks a blocking task
    // for as long as the dialog is open, and the dispatch loop has to keep
    // delivering `session/update` notifications while it does.
    let runtime = match tokio::runtime::Builder::new_multi_thread()
        .worker_threads(2)
        .enable_all()
        .build()
    {
        Ok(runtime) => runtime,
        Err(e) => {
            let _ = ready_tx.send(Err(format!("无法创建 dsh 异步运行时: {e}")));
            return;
        }
    };

    let index = Arc::new(ToolCallIndex::default());
    // ACP may accept initialize while another Cordis plugin is still loading.
    // Wait for the plugin's explicit post-registration receipt, not a delay or
    // a successful ACP response, before exposing a usable runtime.
    let (startup_tx, startup_rx) = tokio::sync::oneshot::channel();
    let startup_tx = std::sync::Mutex::new(Some(startup_tx));
    let startup_error_tx = ready_tx.clone();
    let agent = AcpAgent::new(agent_config).with_debug(move |line: &str, direction| {
        // Only the child's own diagnostics; protocol frames are handled above.
        if matches!(direction, agent_client_protocol::LineDirection::Stderr) {
            if let Some(outcome) = line.trim().strip_prefix(super::contract::STARTUP_PREFIX) {
                let result = match outcome {
                    "ready" => Ok(()),
                    "mismatch" => Err(super::contract::MISMATCH_MESSAGE.to_string()),
                    "unreachable" => Err(super::contract::UNREACHABLE_MESSAGE.to_string()),
                    _ => return,
                };
                if let Err(message) = &result {
                    // Preserve the readable cause even if the child exits before
                    // connect_with gets scheduled to consume the receipt.
                    let _ = startup_error_tx.send(Err(message.clone()));
                }
                if let Some(sender) = startup_tx.lock().unwrap().take() {
                    let _ = sender.send(result);
                }
            } else {
                eprintln!("[dsh] {line}");
            }
        }
    });

    let ready_for_main = ready_tx.clone();
    let closed_sink = Arc::clone(&on_update);
    let result = runtime.block_on(async move {
        agent_client_protocol::Client
            .builder()
            .name("lamber")
            .on_receive_notification(
                {
                    let index = Arc::clone(&index);
                    let on_update = Arc::clone(&on_update);
                    async move |notification: SessionNotification, _cx| {
                        handle_notification(&index, &on_update, &notification);
                        Ok(())
                    }
                },
                agent_client_protocol::on_receive_notification!(),
            )
            .on_receive_request(
                {
                    let index = Arc::clone(&index);
                    let on_permission = Arc::clone(&on_permission);
                    async move |request: RequestPermissionRequest, responder, cx| {
                        handle_permission(&index, &on_permission, request, responder, &cx)
                    }
                },
                agent_client_protocol::on_receive_request!(),
            )
            .connect_with(agent, move |connection: ConnectionTo<Agent>| async move {
                let startup = match tokio::time::timeout(Duration::from_secs(15), startup_rx).await {
                    Ok(Ok(result)) => result,
                    // No receipt means this plugin does not implement our startup
                    // contract. Never permit a session with an unverified plugin.
                    _ => Err(super::contract::MISMATCH_MESSAGE.to_string()),
                };
                if let Err(message) = startup {
                    let _ = ready_for_main.send(Err(message));
                    return Ok(());
                }
                let handshake = match initialize(&connection).await {
                    Ok(handshake) => handshake,
                    Err(message) => {
                        let _ = ready_for_main.send(Err(message.clone()));
                        // Reported to the caller already; end the connection quietly.
                        return Ok(());
                    }
                };
                let _ = ready_for_main.send(Ok(handshake));
                serve_commands(&connection, command_rx, on_update).await;
                Ok(())
            })
            .await
    });

    closed_sink("connection/closed", &serde_json::json!({}));
    if let Err(e) = result {
        // `start` has usually already reported; a send to a hung-up channel is
        // the normal case and means the failure arrived after the handshake.
        let _ = ready_tx.send(Err(format!("dsh ACP 连接失败: {e}")));
        eprintln!("[agent_bridge] dsh ACP 连接结束: {e}");
    }
}

/// Handshake, and refuse any protocol version but the one this client implements.
async fn initialize(connection: &ConnectionTo<Agent>) -> Result<AgentHandshake, String> {
    let response = connection
        .send_request(InitializeRequest::new(EXPECTED_PROTOCOL_VERSION))
        .block_task()
        .await
        .map_err(|e| format!("ACP initialize 失败: {e}"))?;

    let handshake = AgentHandshake::from_response(&response);
    if handshake.protocol_version != EXPECTED_PROTOCOL_VERSION {
        return Err(format!(
            "ACP 协议版本不匹配：本客户端实现 {:?}，dsh 返回 {:?}。\
             dsh 不做版本协商（收到任何版本都回自己的版本号），\
             因此这里必须直接失败，而不是带着不兼容的假设继续跑。",
            EXPECTED_PROTOCOL_VERSION, handshake.protocol_version
        ));
    }
    Ok(handshake)
}

/// Translate commands from the synchronous side until the channel closes.
async fn serve_commands(
    connection: &ConnectionTo<Agent>,
    mut command_rx: tokio::sync::mpsc::UnboundedReceiver<Command>,
    on_update: SessionUpdateSink,
) {
    while let Some(command) = command_rx.recv().await {
        match command {
            Command::Close { session_id, reply } => {
                let outcome = connection
                    .send_request(CloseSessionRequest::new(SessionId::new(session_id)))
                    .block_task()
                    .await
                    .map(|_| ())
                    .map_err(|e| format!("关闭 AI 会话失败: {e}"));
                let _ = reply.send(outcome);
            }
            Command::SetModel {
                session_id,
                value,
                reply,
            } => {
                let outcome = connection
                    .send_request(SetSessionConfigOptionRequest::new(
                        SessionId::new(session_id),
                        "model",
                        value.as_str(),
                    ))
                    .block_task()
                    .await
                    .map(|_| ())
                    .map_err(|e| format!("同步会话模型失败: {e}"));
                let _ = reply.send(outcome);
            }
            Command::Resume {
                session_id,
                cwd,
                reply,
            } => {
                let outcome = connection
                    .send_request(ResumeSessionRequest::new(SessionId::new(session_id), cwd))
                    .block_task()
                    .await
                    .map(|_| ())
                    .map_err(|e| format!("无法恢复原 AI 会话，请保留记录并新建会话: {e}"));
                let _ = reply.send(outcome);
            }
            Command::Cancel { session_id, reply } => {
                let outcome = connection
                    .send_notification(CancelNotification::new(SessionId::new(session_id)))
                    .map_err(|e| format!("ACP session/cancel 失败: {e}"));
                let _ = reply.send(outcome);
            }
            Command::NewSession { cwd, reply } => {
                let outcome = connection
                    .send_request(NewSessionRequest::new(cwd))
                    .block_task()
                    .await
                    .map(|response| response.session_id.0.to_string())
                    .map_err(|e| format!("ACP session/new 失败: {e}"));
                let _ = reply.send(outcome);
            }
            Command::Prompt {
                session_id,
                content,
                reply,
            } => {
                let request = PromptRequest::new(SessionId::new(session_id.clone()), content);
                // Fire-and-forget: the turn's result is announced as an event so
                // the caller is not held for the length of a model turn.
                let sink = Arc::clone(&on_update);
                let sent = connection.send_request(request).on_receiving_result({
                    move |result| {
                        let payload = match &result {
                            Ok(response) => serde_json::json!({
                                "sessionId": session_id,
                                "stopReason": format!("{:?}", response.stop_reason),
                            }),
                            Err(e) => serde_json::json!({
                                "sessionId": session_id,
                                "error": e.to_string(),
                            }),
                        };
                        sink(TURN_ENDED_METHOD, &payload);
                        async move { Ok(()) }
                    }
                });
                let _ = reply.send(sent.map_err(|e| format!("ACP session/prompt 发送失败: {e}")));
            }
        }
    }
}

/// Forward one notification to the frontend, indexing tool calls on the way.
fn handle_notification(
    index: &ToolCallIndex,
    on_update: &SessionUpdateSink,
    notification: &SessionNotification,
) {
    // dsh announces a call before it asks permission for it (it awaits
    // `drainUpdates()` first), and the permission request itself carries only an
    // id — so this is the only place the tool's name and arguments are seen.
    if let SessionUpdate::ToolCall(call) = &notification.update {
        index.record(
            call.tool_call_id.0.as_ref(),
            TrackedCall {
                tool_name: call.title.clone(),
                args: call.raw_input.clone().unwrap_or(serde_json::Value::Null),
            },
        );
    }

    match serde_json::to_value(notification) {
        Ok(params) => on_update(UPDATE_METHOD, &params),
        Err(e) => eprintln!("[agent_bridge] session/update 序列化失败: {e}"),
    }
}

/// Ask a human about one tool call and answer `session/requestPermission`.
///
/// The decision is taken on a blocking task rather than inline: handler bodies
/// run on the connection's dispatch loop, and parking there for the length of a
/// dialog would stop every other message on the connection.
fn handle_permission(
    index: &Arc<ToolCallIndex>,
    on_permission: &PermissionResponder,
    request: RequestPermissionRequest,
    responder: agent_client_protocol::Responder<RequestPermissionResponse>,
    cx: &ConnectionTo<Agent>,
) -> Result<(), agent_client_protocol::Error> {
    let call_id = request.tool_call.tool_call_id.0.to_string();
    let tracked = index.take(&call_id);
    let options = request.options.clone();
    let on_permission = Arc::clone(on_permission);

    cx.spawn(async move {
        let question = ApprovalQuestion {
            session_id: Some(request.session_id.0.to_string()),
            intent: None,
            // An unannounced call is still a call: name it by its id and ask
            // anyway. Answering "no" without asking would be safe but wrong —
            // it would deny work the user never got to see.
            tool_name: tracked
                .as_ref()
                .map(|call| call.tool_name.clone())
                .unwrap_or_else(|| format!("未知工具（调用 {call_id}）")),
            call_id: Some(call_id),
            reason: Some(
                tracked
                    .as_ref()
                    .and_then(|call| super::approval::gated_tool_reason(&call.tool_name))
                    .unwrap_or("该工具调用需要你的确认后才会执行。")
                    .to_string(),
            ),
            args: tracked
                .map(|call| call.args)
                .unwrap_or(serde_json::Value::Null),
        };

        let decision = tokio::task::spawn_blocking(move || on_permission(question))
            .await
            .unwrap_or_else(|e| ApprovalDecision {
                modified_args: None,
                approved: false,
                reason: format!("审批任务异常终止，按拒绝处理: {e}"),
            });

        responder.respond(RequestPermissionResponse::new(select_outcome(
            &options,
            decision.approved,
        )))
    })
}

/// Pick the option matching the user's answer, by kind rather than by id.
///
/// dsh happens to send `allow-once` / `reject-once` today
/// (`dsh-acp/lib/index.js:1126-1133`), but option ids are the agent's to choose;
/// the `kind` is what the protocol defines. An agent that offers no option of
/// the needed kind gets `Cancelled`, which every ACP agent must treat as "did
/// not proceed" — never an accidental grant.
fn select_outcome(
    options: &[agent_client_protocol::schema::v1::PermissionOption],
    approved: bool,
) -> RequestPermissionOutcome {
    let wanted = if approved {
        [
            PermissionOptionKind::AllowOnce,
            PermissionOptionKind::AllowAlways,
        ]
    } else {
        [
            PermissionOptionKind::RejectOnce,
            PermissionOptionKind::RejectAlways,
        ]
    };
    for kind in wanted {
        if let Some(option) = options.iter().find(|option| option.kind == kind) {
            return RequestPermissionOutcome::Selected(SelectedPermissionOutcome::new(
                option.option_id.clone(),
            ));
        }
    }
    RequestPermissionOutcome::Cancelled
}
