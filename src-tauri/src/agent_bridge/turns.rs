//! Correlate events before queuing a prompt, including events arriving before invoke returns.
use serde_json::{json, Value};
use std::collections::HashMap;
use std::sync::Mutex;

#[derive(Default)]
pub struct Turns(Mutex<HashMap<String, (String, Option<String>)>>);

impl Turns {
    pub fn begin(&self, acp: &str, session: &str, request: Option<&str>) -> Result<(), String> {
        let mut turns = self.0.lock().map_err(|_| "AI 轮次锁已中毒")?;
        if turns.contains_key(acp) {
            return Err("此会话正在生成，请先停止或等待结束".into());
        }
        turns.insert(acp.into(), (session.into(), request.map(String::from)));
        Ok(())
    }

    pub fn binding(&self, acp: &str) -> Option<Value> {
        let turns = self.0.lock().ok()?;
        let (session, request) = turns.get(acp)?;
        Some(json!({"lamberSessionId": session, "requestId": request}))
    }

    /// HTTP events carry the binding acquired before model execution. Never
    /// relabel a delayed event with whichever request happens to be active now.
    pub fn emit_stream(&self, params: Value, sink: impl Fn(Value)) {
        let Ok(turns) = self.0.lock() else { return };
        let Some(acp) = params["sessionId"].as_str() else {
            return;
        };
        let Some((session, request)) = turns.get(acp) else {
            return;
        };
        if params["lamberSessionId"].as_str() != Some(session.as_str())
            || params["requestId"].as_str() != request.as_deref()
        {
            return;
        }
        sink(
            json!({"method": "session/stream", "lamberSessionId": session,
            "requestId": request, "params": params}),
        );
    }

    pub fn matches(&self, session: &str, request: &str) -> Option<String> {
        self.0.lock().ok()?.iter().find_map(|(acp, (id, turn))| {
            (id == session && turn.as_deref() == Some(request)).then(|| acp.clone())
        })
    }

    pub fn any_active(&self) -> bool {
        self.0.lock().map(|turns| !turns.is_empty()).unwrap_or(true)
    }

    pub fn is_active(&self, session: &str) -> bool {
        self.0
            .lock()
            .map(|turns| turns.values().any(|(id, _)| id == session))
            .unwrap_or(true)
    }

    pub fn forget(&self, acp: &str) {
        if let Ok(mut turns) = self.0.lock() {
            turns.remove(acp);
        }
    }

    pub fn emit(&self, method: &str, params: &Value, sink: impl Fn(Value)) {
        let Ok(mut turns) = self.0.lock() else { return };
        if method == "connection/closed" {
            for (acp, (session, request)) in turns.drain() {
                sink(
                    json!({"method": "session/turn-ended", "lamberSessionId": session, "requestId": request,
                    "params": {"sessionId": acp, "error": "AI 运行组件已停止，请重新发送"}}),
                );
            }
            return;
        }
        let acp = params
            .get("sessionId")
            .and_then(Value::as_str)
            .unwrap_or_default();
        let mut event = json!({"method": method, "params": params});
        if let Some((session, request)) = turns.get(acp) {
            event["lamberSessionId"] = json!(session);
            event["requestId"] = json!(request);
        }
        sink(event);
        if method == "session/turn-ended" {
            turns.remove(acp);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn routing_is_fixed_before_send_and_old_cancel_cannot_target_next_turn() {
        let turns = Turns::default();
        turns.begin("acp-a", "front-a", Some("turn-1")).unwrap();
        assert!(turns.begin("acp-a", "front-b", Some("turn-2")).is_err());
        assert_eq!(turns.matches("front-a", "turn-1").as_deref(), Some("acp-a"));
        turns.emit("session/update", &json!({"sessionId": "acp-a"}), |event| {
            assert_eq!(event["lamberSessionId"], "front-a");
            assert_eq!(event["requestId"], "turn-1");
        });
        turns.emit("session/turn-ended", &json!({"sessionId": "acp-a"}), |_| {});
        turns.begin("acp-a", "front-a", Some("turn-2")).unwrap();
        assert!(turns.matches("front-a", "turn-1").is_none());
        assert!(turns.matches("front-b", "turn-2").is_none());
        turns.emit("connection/closed", &json!({}), |event| {
            assert_eq!(event["requestId"], "turn-2");
            assert!(event["params"]["error"].is_string());
        });
        assert!(!turns.any_active());
    }
}
