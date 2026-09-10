//! Authenticated display-only side channel for alpha.5's session firehose.
use super::bridge_server::{BridgeHandler, BridgeReply};
use super::turns::Turns;
use serde::Deserialize;
use serde_json::Value;
use std::sync::Arc;

pub const STREAM_ROUTE: &str = "/lamber-bridge/stream";

#[derive(Deserialize)]
#[serde(tag = "kind", rename_all = "camelCase")]
enum Frame {
    Bind,
    Delta {
        step: u64,
        seq: u64,
        turn: u64,
        chunk: Chunk,
    },
    Commit {
        step: u64,
        seq: u64,
        turn: u64,
        #[serde(rename = "messageId")]
        message_id: String,
    },
    Error {
        error: String,
    },
}
#[derive(Deserialize)]
#[serde(tag = "type", rename_all = "kebab-case")]
enum Chunk {
    TextDelta {
        index: u64,
        bytes: Vec<u8>,
    },
    ReasoningDelta {
        index: u64,
        bytes: Vec<u8>,
    },
    ToolCallDelta {
        index: u64,
        id: String,
        bytes: Vec<u8>,
    },
}
impl Frame {
    fn valid(&self) -> bool {
        const SAFE: u64 = 9_007_199_254_740_991;
        match self {
            Self::Bind => true,
            Self::Error { error } => !error.is_empty() && error.len() <= 1024,
            Self::Commit {
                step,
                seq,
                turn,
                message_id,
            } => {
                *step <= SAFE
                    && *seq <= SAFE
                    && *turn <= SAFE
                    && !message_id.is_empty()
                    && message_id.len() <= 256
            }
            Self::Delta {
                step,
                seq,
                turn,
                chunk,
            } => {
                *step <= SAFE
                    && *seq <= SAFE
                    && *turn <= SAFE
                    && match chunk {
                        Chunk::TextDelta { index, bytes }
                        | Chunk::ReasoningDelta { index, bytes } => {
                            *index <= SAFE && bytes.len() <= 65536
                        }
                        Chunk::ToolCallDelta { index, id, bytes } => {
                            *index <= SAFE
                                && !id.is_empty()
                                && id.len() <= 256
                                && bytes.len() <= 65536
                        }
                    }
            }
        }
    }
}

pub fn handler(
    fallback: BridgeHandler,
    turns: Arc<Turns>,
    sink: Arc<dyn Fn(Value) + Send + Sync>,
) -> BridgeHandler {
    Arc::new(move |path, body| {
        if path != STREAM_ROUTE {
            return fallback(path, body);
        }
        let Ok(value) = serde_json::from_str::<Value>(body) else {
            return BridgeReply::error(400, "实时显示事件不是 JSON");
        };
        let Ok(frame) = serde_json::from_value::<Frame>(value.clone()) else {
            return BridgeReply::error(400, "实时显示事件结构无效");
        };
        if !frame.valid() {
            return BridgeReply::error(400, "实时显示事件超出允许范围");
        }
        let Some(session) = value["sessionId"].as_str().filter(|s| !s.is_empty()) else {
            return BridgeReply::error(400, "实时显示缺少会话标识");
        };
        if matches!(frame, Frame::Bind) {
            return match turns.binding(session) {
                Some(binding) => BridgeReply::ok(binding.to_string()),
                None => BridgeReply::error(409, "没有可绑定的 AI 请求"),
            };
        }
        turns.emit_stream(value, |event| sink(event));
        BridgeReply::ok("{}".into())
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;
    use std::sync::Mutex;

    #[test]
    fn stream_binding_rejects_late_previous_turn_and_validates_bytes() {
        let turns = Arc::new(Turns::default());
        let events = Arc::new(Mutex::new(Vec::new()));
        let captured = Arc::clone(&events);
        let route = handler(
            Arc::new(|_, _| BridgeReply::error(404, "not found")),
            Arc::clone(&turns),
            Arc::new(move |event| captured.lock().unwrap().push(event)),
        );
        let bind = json!({"kind": "bind", "sessionId": "acp"}).to_string();
        assert_eq!(route(STREAM_ROUTE, &bind).status, 409);
        turns.begin("acp", "front", Some("old")).unwrap();
        let binding: Value = serde_json::from_str(&route(STREAM_ROUTE, &bind).body).unwrap();
        assert_eq!(binding["requestId"], "old");
        let mut frame = json!({"kind": "delta", "sessionId": "acp", "lamberSessionId": "front",
            "requestId": "old", "turn": 1, "step": 1, "seq": 12,
            "chunk": {"type": "text-delta", "index": 0, "bytes": [228, 184]}});
        assert_eq!(route(STREAM_ROUTE, &frame.to_string()).status, 200);
        assert_eq!(events.lock().unwrap().len(), 1);
        turns.emit("session/turn-ended", &json!({"sessionId": "acp"}), |_| {});
        turns.begin("acp", "front", Some("new")).unwrap();
        route(STREAM_ROUTE, &frame.to_string());
        assert_eq!(
            events.lock().unwrap().len(),
            1,
            "old HTTP event must not become new turn text"
        );
        frame["requestId"] = json!("new");
        frame["chunk"]["bytes"] = json!([256]);
        assert_eq!(route(STREAM_ROUTE, &frame.to_string()).status, 400);
        frame["chunk"]["bytes"] = json!([173]);
        route(STREAM_ROUTE, &frame.to_string());
        assert_eq!(events.lock().unwrap().len(), 2);
        assert_eq!(route("/unknown", "{}").status, 404);
    }
}
