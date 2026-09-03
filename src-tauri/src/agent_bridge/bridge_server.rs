//! Loopback HTTP bridge that lets the dsh agent runtime call back into lamber.
//!
//! The dsh child process is a separate OS process, so it cannot reach lamber's
//! Rust functions directly. This module hosts a minimal HTTP server bound to
//! `127.0.0.1` on an ephemeral port; the `dsh-tool-lamber` plugin posts JSON to
//! it from inside a tool body and lamber dispatches into the existing business
//! modules.
//!
//! Two properties matter for safety:
//!
//! * The listener binds loopback only, so nothing off-machine can reach it.
//! * Every request must carry a per-launch bearer token. Loopback alone is not
//!   an authorization boundary — any local process could otherwise read a
//!   customer's project financials — so the token is generated per server and
//!   handed to the child through its environment.
//!
//! Each request runs on its own worker thread. The approval route parks its
//! request for as long as the user takes to answer, so a single-threaded accept
//! loop would stall every other route behind one open dialog.

use std::io::Read;
use std::net::{Ipv4Addr, SocketAddr, SocketAddrV4};
use std::sync::atomic::{AtomicBool, AtomicUsize, Ordering};
use std::sync::Arc;
use std::thread::JoinHandle;

/// Header carrying the per-launch bridge token.
pub const BRIDGE_TOKEN_HEADER: &str = "x-lamber-bridge-token";

/// Largest request body the bridge accepts, guarding against a runaway client.
const MAX_BODY_BYTES: usize = 1024 * 1024;

/// Cap on concurrently served requests, so a misbehaving client cannot spawn
/// threads without bound. Well above what one agent generates.
const MAX_IN_FLIGHT: usize = 16;

/// Outcome of one bridge route: an HTTP status plus a JSON body.
pub struct BridgeReply {
    pub status: u16,
    pub body: String,
}

impl BridgeReply {
    pub fn ok(body: String) -> Self {
        Self { status: 200, body }
    }

    /// Render an error as the JSON envelope the plugin surfaces to the model.
    pub fn error(status: u16, message: &str) -> Self {
        Self {
            status,
            body: serde_json::json!({ "error": message }).to_string(),
        }
    }
}

/// Route dispatcher: receives the request path and raw JSON body.
pub type BridgeHandler = Arc<dyn Fn(&str, &str) -> BridgeReply + Send + Sync>;

/// A running loopback bridge. Dropping it stops the server thread.
pub struct BridgeServer {
    server: Arc<tiny_http::Server>,
    addr: SocketAddr,
    token: String,
    stopping: Arc<AtomicBool>,
    worker: Option<JoinHandle<()>>,
}

impl BridgeServer {
    /// Bind an ephemeral loopback port and start serving `handler` on a worker thread.
    ///
    /// @param handler - route dispatcher invoked for every authenticated request.
    /// @returns the running server; its `origin()` and `token()` are what the child process needs.
    pub fn start(handler: BridgeHandler) -> Result<Self, String> {
        let bind = SocketAddrV4::new(Ipv4Addr::LOCALHOST, 0);
        let server = tiny_http::Server::http(bind)
            .map_err(|e| format!("无法启动 AI 桥接服务: {e}"))?;
        let addr = server
            .server_addr()
            .to_ip()
            .ok_or_else(|| "AI 桥接服务未绑定到 IP 地址".to_string())?;
        let server = Arc::new(server);
        let token = uuid::Uuid::new_v4().simple().to_string();
        let stopping = Arc::new(AtomicBool::new(false));

        let worker = {
            let server = Arc::clone(&server);
            let stopping = Arc::clone(&stopping);
            let token = token.clone();
            std::thread::Builder::new()
                .name("lamber-agent-bridge".into())
                .spawn(move || serve_loop(server, stopping, token, handler))
                .map_err(|e| format!("无法启动 AI 桥接服务线程: {e}"))?
        };

        Ok(Self {
            server,
            addr,
            token,
            stopping,
            worker: Some(worker),
        })
    }

    /// Origin the child process should use, e.g. `http://127.0.0.1:53211`.
    pub fn origin(&self) -> String {
        format!("http://{}:{}", self.addr.ip(), self.addr.port())
    }

    /// Per-launch bearer token every bridge request must present.
    pub fn token(&self) -> &str {
        &self.token
    }
}

impl Drop for BridgeServer {
    fn drop(&mut self) {
        self.stopping.store(true, Ordering::SeqCst);
        self.server.unblock();
        if let Some(worker) = self.worker.take() {
            let _ = worker.join();
        }
    }
}

fn serve_loop(
    server: Arc<tiny_http::Server>,
    stopping: Arc<AtomicBool>,
    token: String,
    handler: BridgeHandler,
) {
    let in_flight = Arc::new(AtomicUsize::new(0));
    loop {
        let request = match server.recv() {
            Ok(request) => request,
            // `unblock()` makes `recv` return an error; a shutdown is not a failure.
            Err(_) => break,
        };
        if stopping.load(Ordering::SeqCst) {
            break;
        }

        if in_flight.load(Ordering::SeqCst) >= MAX_IN_FLIGHT {
            respond(request, BridgeReply::error(503, "AI 桥接服务并发请求过多"));
            continue;
        }

        // One thread per request: the approval route blocks until the user
        // answers, and must not hold up the calculation route behind it.
        in_flight.fetch_add(1, Ordering::SeqCst);
        let token = token.clone();
        let handler = Arc::clone(&handler);
        let worker_counter = Arc::clone(&in_flight);
        let spawned = std::thread::Builder::new()
            .name("lamber-agent-bridge-req".into())
            .spawn(move || {
                handle_request(request, &token, &handler);
                worker_counter.fetch_sub(1, Ordering::SeqCst);
            });
        if spawned.is_err() {
            in_flight.fetch_sub(1, Ordering::SeqCst);
            eprintln!("[agent_bridge] 无法为请求创建线程");
        }
    }
}

/// Send one reply, discarding a write failure on an already-closed connection.
fn respond(request: tiny_http::Request, reply: BridgeReply) {
    let response = tiny_http::Response::from_string(reply.body)
        .with_status_code(reply.status)
        .with_header(
            tiny_http::Header::from_bytes(&b"Content-Type"[..], &b"application/json"[..])
                .expect("static header is valid"),
        );
    let _ = request.respond(response);
}

fn handle_request(mut request: tiny_http::Request, token: &str, handler: &BridgeHandler) {
    let reply = match validate(&request, token) {
        Err(reply) => reply,
        Ok(()) => {
            let path = request.url().split('?').next().unwrap_or("").to_string();
            match read_body(&mut request) {
                Ok(body) => handler(&path, &body),
                Err(message) => BridgeReply::error(400, &message),
            }
        }
    };

    respond(request, reply);
}

/// Reject anything that is not an authenticated POST before touching the body.
fn validate(request: &tiny_http::Request, token: &str) -> Result<(), BridgeReply> {
    if request.method() != &tiny_http::Method::Post {
        return Err(BridgeReply::error(405, "AI 桥接服务只接受 POST 请求"));
    }
    let presented = request
        .headers()
        .iter()
        .find(|h| h.field.equiv(BRIDGE_TOKEN_HEADER))
        .map(|h| h.value.as_str().to_string())
        .unwrap_or_default();
    if !constant_time_eq(presented.as_bytes(), token.as_bytes()) {
        return Err(BridgeReply::error(401, "AI 桥接服务令牌无效"));
    }
    Ok(())
}

fn read_body(request: &mut tiny_http::Request) -> Result<String, String> {
    if let Some(len) = request.body_length() {
        if len > MAX_BODY_BYTES {
            return Err("请求体超出 AI 桥接服务允许的大小".to_string());
        }
    }
    let mut body = String::new();
    request
        .as_reader()
        .take(MAX_BODY_BYTES as u64 + 1)
        .read_to_string(&mut body)
        .map_err(|e| format!("读取请求体失败: {e}"))?;
    if body.len() > MAX_BODY_BYTES {
        return Err("请求体超出 AI 桥接服务允许的大小".to_string());
    }
    Ok(body)
}

/// Length-independent comparison so a wrong token leaks no timing information.
fn constant_time_eq(a: &[u8], b: &[u8]) -> bool {
    if a.len() != b.len() {
        return false;
    }
    let mut diff = 0u8;
    for (x, y) in a.iter().zip(b.iter()) {
        diff |= x ^ y;
    }
    diff == 0
}
