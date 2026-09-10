//! One source manifest, embedded in the binary (never loaded from mutable resources).
use super::bridge_server::{BridgeHandler, BridgeReply};
use serde::Deserialize;
use std::sync::Arc;

pub const HANDSHAKE_ROUTE: &str = "/lamber-bridge/handshake";
pub const STARTUP_PREFIX: &str = "LAMBER_BRIDGE_STARTUP:";
pub const MISMATCH_MESSAGE: &str =
    "AI 组件版本不匹配，无法启动。请完整重新构建 Lamber 和 AI 组件，或重新安装最新版本。";
pub const UNREACHABLE_MESSAGE: &str =
    "无法连接 AI 本地服务，无法启动。请关闭并重新打开 Lamber；若仍失败，请重新安装。";
// Also used by packaging to inspect the actual target binary without executing
// it (works for cross-platform builds and Windows GUI binaries).
pub const EMBEDDED: &str = concat!(
    "LAMBER_BRIDGE_CONTRACT:",
    include_str!("../../../agent-bridge/bridge-contract.json"),
    ":END_LAMBER_BRIDGE_CONTRACT"
);

#[derive(Deserialize)]
pub struct Contract {
    pub version: u32,
    pub routes: Vec<String>,
}

pub fn manifest() -> &'static str {
    // Retain the exact attestation bytes even in optimized target binaries.
    std::hint::black_box(EMBEDDED)
        .strip_prefix("LAMBER_BRIDGE_CONTRACT:")
        .unwrap()
        .strip_suffix(":END_LAMBER_BRIDGE_CONTRACT")
        .unwrap()
}

pub fn handler(fallback: BridgeHandler) -> BridgeHandler {
    Arc::new(move |path, body| {
        if path != HANDSHAKE_ROUTE {
            return fallback(path, body);
        }
        let expected: Contract =
            serde_json::from_str(manifest()).expect("compiled bridge contract");
        let Ok(request) = serde_json::from_str::<Contract>(body) else {
            return BridgeReply::error(409, MISMATCH_MESSAGE);
        };
        if request.version != expected.version
            || request.routes.iter().any(|r| !expected.routes.contains(r))
        {
            return BridgeReply::error(409, MISMATCH_MESSAGE);
        }
        BridgeReply::ok(manifest().to_string())
    })
}
