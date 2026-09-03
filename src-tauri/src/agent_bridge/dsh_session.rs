//! Supervises the `dsh` agent runtime child process and speaks its SDK protocol.
//!
//! The wire format is newline-delimited JSON-RPC 2.0 over the child's stdio
//! (`@deepseek-ai/dsh-sdk-protocol`): one JSON object per line. A line carrying
//! `id` is a response to one of our requests; a line without `id` is a
//! server-to-client notification (`session.event`, `session.status`,
//! `subagent.started`, `subagent.finished`).
//!
//! That protocol is small enough that a hand-written client is the honest
//! choice — it avoids embedding a Node SDK just to concatenate JSON lines, and
//! it keeps the reader on a plain thread, consistent with lamber's otherwise
//! synchronous backend.

use serde_json::{json, Value};
use std::collections::HashMap;
use std::io::{BufRead, BufReader, Write};
use std::path::PathBuf;
use std::process::{Child, ChildStdin, Command, Stdio};
use std::sync::atomic::{AtomicI64, Ordering};
use std::sync::{Arc, Condvar, Mutex};
use std::time::{Duration, Instant};

use super::bridge_server::BRIDGE_TOKEN_HEADER;

/// How long a request waits for its matching response before giving up.
const REQUEST_TIMEOUT: Duration = Duration::from_secs(60);

/// Everything needed to launch one dsh runtime bound to this lamber instance.
///
/// `Debug` is hand-written to redact the API key and the bridge token: a config
/// dump in a panic message, a log line, or a bug report must never carry a
/// live credential.
#[derive(Clone)]
pub struct DshLaunchConfig {
    /// Path to the `dsh` executable (usually `agent-bridge/node_modules/.bin/dsh`).
    pub dsh_bin: PathBuf,
    /// Profile to boot; the SDK JSON-RPC server lives in the `sdk` profile.
    pub profile: String,
    /// Absolute path to `agent-bridge/patch.yml`, which mounts `dsh-tool-lamber`.
    pub patch_path: PathBuf,
    /// Writable `$DSH_HOME` holding profiles, sessions, and credentials.
    pub dsh_home: PathBuf,
    /// Working directory recorded on every session header.
    pub cwd: PathBuf,
    /// Provider route registered in the profile, e.g. `deepseek-official`.
    pub provider: String,
    /// Model on that route, e.g. `deepseek-v4-flash`.
    pub model: String,
    /// DeepSeek API key. Absent means the runtime boots but every turn fails at the LLM call.
    pub api_key: Option<String>,
    /// Origin of this instance's bridge server.
    pub bridge_url: String,
    /// Per-launch bridge token the plugin must present.
    pub bridge_token: String,
}

impl std::fmt::Debug for DshLaunchConfig {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("DshLaunchConfig")
            .field("dsh_bin", &self.dsh_bin)
            .field("profile", &self.profile)
            .field("patch_path", &self.patch_path)
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
    pub fn from_repo_root(repo_root: &std::path::Path) -> Self {
        let agent_bridge = repo_root.join("agent-bridge");
        Self {
            dsh_bin: agent_bridge.join("node_modules/.bin/dsh"),
            profile: "sdk".to_string(),
            patch_path: agent_bridge.join("patch.yml"),
            dsh_home: agent_bridge.join(".dsh-home"),
            cwd: repo_root.to_path_buf(),
            provider: "deepseek-official".to_string(),
            model: "deepseek-v4-flash".to_string(),
            api_key: std::env::var("DEEPSEEK_API_KEY").ok().filter(|k| !k.is_empty()),
            bridge_url: String::new(),
            bridge_token: String::new(),
        }
    }
}

/// Callback invoked on the reader thread for every server-to-client notification.
pub type NotificationSink = Arc<dyn Fn(&str, &Value) + Send + Sync>;

/// Pending-response table shared between callers and the reader thread.
#[derive(Default)]
struct Pending {
    responses: HashMap<i64, Result<Value, String>>,
    /// Set when the child's stdout closes, so blocked callers fail instead of hanging.
    closed: bool,
}

/// A live dsh runtime: the child process plus its JSON-RPC client.
pub struct DshSession {
    child: Child,
    stdin: Mutex<ChildStdin>,
    next_id: AtomicI64,
    pending: Arc<(Mutex<Pending>, Condvar)>,
    reader: Option<std::thread::JoinHandle<()>>,
    stderr_reader: Option<std::thread::JoinHandle<()>>,
}

impl DshSession {
    /// Spawn `dsh --profile <profile> --patch <patch>` and start reading its stdout.
    ///
    /// @param config - launch parameters, including the bridge coordinates handed to the plugin.
    /// @param on_notification - called with `(method, params)` for every notification line.
    /// @returns the supervised session, before `initialize` has been sent.
    pub fn spawn(
        config: &DshLaunchConfig,
        on_notification: NotificationSink,
    ) -> Result<Self, String> {
        if !config.dsh_bin.exists() {
            return Err(format!(
                "未找到 dsh 可执行文件：{}。请先在 agent-bridge/ 目录运行 `npm install` 与 `npm run provision`。",
                config.dsh_bin.display()
            ));
        }

        let mut command = Command::new(&config.dsh_bin);
        command
            .arg("--profile")
            .arg(&config.profile)
            .arg("--patch")
            .arg(&config.patch_path)
            .current_dir(&config.cwd)
            .env("DSH_HOME", &config.dsh_home)
            // The sdk profile ships telemetry to an external host by default; lamber
            // handles customer financial data, so it stays off.
            .env("DSH_TELEMETRY_MODE", "DISABLED")
            .env("LAMBER_BRIDGE_URL", &config.bridge_url)
            .env("LAMBER_BRIDGE_TOKEN", &config.bridge_token)
            .env("LAMBER_BRIDGE_TOKEN_HEADER", BRIDGE_TOKEN_HEADER)
            .stdin(Stdio::piped())
            .stdout(Stdio::piped())
            .stderr(Stdio::piped());

        match &config.api_key {
            Some(key) => {
                command.env("DEEPSEEK_API_KEY", key);
            }
            // Inheriting a stale key from lamber's own environment would silently
            // bill the wrong account; an absent key must stay absent.
            None => {
                command.env_remove("DEEPSEEK_API_KEY");
            }
        }

        let mut child = command
            .spawn()
            .map_err(|e| format!("启动 dsh 子进程失败: {e}"))?;

        let stdin = child
            .stdin
            .take()
            .ok_or_else(|| "dsh 子进程没有 stdin".to_string())?;
        let stdout = child
            .stdout
            .take()
            .ok_or_else(|| "dsh 子进程没有 stdout".to_string())?;
        let stderr = child
            .stderr
            .take()
            .ok_or_else(|| "dsh 子进程没有 stderr".to_string())?;

        let pending: Arc<(Mutex<Pending>, Condvar)> = Arc::default();
        let reader = {
            let pending = Arc::clone(&pending);
            std::thread::Builder::new()
                .name("lamber-dsh-stdout".into())
                .spawn(move || read_stdout(stdout, pending, on_notification))
                .map_err(|e| format!("无法启动 dsh 读取线程: {e}"))?
        };
        let stderr_reader = std::thread::Builder::new()
            .name("lamber-dsh-stderr".into())
            .spawn(move || {
                for line in BufReader::new(stderr).lines().map_while(Result::ok) {
                    eprintln!("[dsh] {line}");
                }
            })
            .map_err(|e| format!("无法启动 dsh 错误流线程: {e}"))?;

        Ok(Self {
            child,
            stdin: Mutex::new(stdin),
            next_id: AtomicI64::new(0),
            pending,
            reader: Some(reader),
            stderr_reader: Some(stderr_reader),
        })
    }

    /// Send one JSON-RPC request and block until its response arrives.
    ///
    /// @param method - protocol method name.
    /// @param params - method params, or `Value::Null` for none.
    /// @returns the `result` payload, or the server-reported error.
    pub fn request(&self, method: &str, params: Value) -> Result<Value, String> {
        let id = self.next_id.fetch_add(1, Ordering::SeqCst) + 1;
        let mut frame = json!({ "jsonrpc": "2.0", "id": id, "method": method });
        if !params.is_null() {
            frame["params"] = params;
        }
        self.write_line(&frame.to_string())?;
        self.await_response(id)
    }

    /// Perform the process-wide SDK handshake.
    pub fn initialize(&self, config: &DshLaunchConfig) -> Result<Value, String> {
        self.request(
            "initialize",
            json!({
                "cwd": config.cwd.to_string_lossy(),
                "provider": config.provider,
                "model": config.model,
            }),
        )
    }

    /// Queue one user turn on a session; events stream back as notifications.
    ///
    /// @param session_id - a fresh id per turn-series; reusing a persisted id collides.
    /// @param text - the user's prompt text.
    pub fn prompt(&self, session_id: &str, text: &str) -> Result<Value, String> {
        self.request(
            "session/prompt",
            json!({
                "sessionId": session_id,
                "contentBlocks": [{ "type": "text", "text": text }],
            }),
        )
    }

    /// Ask the runtime to shut down cleanly; ignores an already-dead child.
    pub fn shutdown(&self) -> Result<(), String> {
        self.request("shutdown", Value::Null).map(|_| ())
    }

    fn write_line(&self, line: &str) -> Result<(), String> {
        let mut stdin = self
            .stdin
            .lock()
            .map_err(|_| "dsh stdin 锁已中毒".to_string())?;
        stdin
            .write_all(line.as_bytes())
            .and_then(|_| stdin.write_all(b"\n"))
            .and_then(|_| stdin.flush())
            .map_err(|e| format!("写入 dsh stdin 失败: {e}"))
    }

    fn await_response(&self, id: i64) -> Result<Value, String> {
        let (lock, cvar) = &*self.pending;
        let deadline = Instant::now() + REQUEST_TIMEOUT;
        let mut guard = lock.lock().map_err(|_| "dsh 响应表锁已中毒".to_string())?;
        loop {
            if let Some(result) = guard.responses.remove(&id) {
                return result;
            }
            if guard.closed {
                return Err("dsh 子进程已退出，未收到响应".to_string());
            }
            let remaining = deadline.saturating_duration_since(Instant::now());
            if remaining.is_zero() {
                return Err(format!("等待 dsh 响应超时（请求 id={id}）"));
            }
            let (next, _) = cvar
                .wait_timeout(guard, remaining)
                .map_err(|_| "dsh 响应表锁已中毒".to_string())?;
            guard = next;
        }
    }
}

impl Drop for DshSession {
    fn drop(&mut self) {
        let _ = self.shutdown();
        // The runtime may ignore shutdown if it never initialized; do not leak a child.
        let _ = self.child.kill();
        let _ = self.child.wait();
        if let Some(reader) = self.reader.take() {
            let _ = reader.join();
        }
        if let Some(reader) = self.stderr_reader.take() {
            let _ = reader.join();
        }
    }
}

/// Read newline-delimited frames until the child's stdout closes.
fn read_stdout(
    stdout: std::process::ChildStdout,
    pending: Arc<(Mutex<Pending>, Condvar)>,
    on_notification: NotificationSink,
) {
    let reader = BufReader::new(stdout);
    for line in reader.lines().map_while(Result::ok) {
        let trimmed = line.trim();
        if trimmed.is_empty() {
            continue;
        }
        let Ok(frame) = serde_json::from_str::<Value>(trimmed) else {
            // dsh may print non-protocol diagnostics; they are not fatal.
            eprintln!("[dsh] {trimmed}");
            continue;
        };

        match frame.get("id").and_then(Value::as_i64) {
            // A frame with an id and a method is a server-to-client *request*; the
            // SDK protocol defines none today, so it is logged rather than answered.
            Some(id) if frame.get("method").is_none() => {
                let outcome = if let Some(error) = frame.get("error") {
                    Err(format_rpc_error(error))
                } else {
                    Ok(frame.get("result").cloned().unwrap_or(Value::Null))
                };
                let (lock, cvar) = &*pending;
                if let Ok(mut guard) = lock.lock() {
                    guard.responses.insert(id, outcome);
                    cvar.notify_all();
                }
            }
            _ => {
                if let Some(method) = frame.get("method").and_then(Value::as_str) {
                    let params = frame.get("params").cloned().unwrap_or(Value::Null);
                    on_notification(method, &params);
                }
            }
        }
    }

    let (lock, cvar) = &*pending;
    if let Ok(mut guard) = lock.lock() {
        guard.closed = true;
        cvar.notify_all();
    }
}

fn format_rpc_error(error: &Value) -> String {
    let message = error
        .get("message")
        .and_then(Value::as_str)
        .unwrap_or("未知错误");
    match error.get("code").and_then(Value::as_i64) {
        Some(code) => format!("dsh 返回错误 {code}: {message}"),
        None => format!("dsh 返回错误: {message}"),
    }
}
