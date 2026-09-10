//! One-use handoff of approved arguments. ACP permission itself remains bool.
use super::{
    approval::ApprovalPrompt,
    bridge_server::{BridgeHandler, BridgeReply},
};
use serde_json::Value;
use std::{
    collections::HashMap,
    sync::{Arc, Mutex},
    time::{Duration, Instant},
};
pub const ROUTE: &str = "/lamber-bridge/reviewed-arguments";
struct Entry {
    tool: String,
    original: Value,
    approved: Value,
    expires: Instant,
    intent: Option<WriteIntent>,
}
#[derive(Default)]
pub struct ReviewedArguments(Mutex<HashMap<(String, String), Entry>>);
impl ReviewedArguments {
    pub fn publish(&self, prompt: &ApprovalPrompt, approved: Value) -> Result<(), String> {
        let (Some(session), Some(call)) = (&prompt.session_id, &prompt.call_id) else {
            return Ok(());
        };
        let mut entries = self.0.lock().map_err(|_| "审批参数不可用")?;
        entries.retain(|_, entry| entry.expires > Instant::now());
        if entries.len() >= 128 {
            return Err("未使用的审批过多，请重新启动 AI 会话".into());
        }
        entries.insert(
            (session.clone(), call.clone()),
            Entry {
                intent: prompt.intent.clone(),
                tool: prompt.tool_name.clone(),
                original: prompt.args.clone(),
                approved,
                expires: Instant::now() + Duration::from_secs(600),
            },
        );
        Ok(())
    }
    pub fn clear(&self) {
        if let Ok(mut entries) = self.0.lock() {
            entries.clear();
        }
    }
    pub fn consume(
        &self,
        session: &str,
        call: &str,
        tool: &str,
        original: &Value,
    ) -> Result<Value, String> {
        self.take(session, call, tool, original)
            .map(|(args, _)| args)
    }
    pub fn take(
        &self,
        session: &str,
        call: &str,
        tool: &str,
        original: &Value,
    ) -> Result<(Value, Option<WriteIntent>), String> {
        let mut entries = self.0.lock().map_err(|_| "审批参数不可用")?;
        let key = (session.to_string(), call.to_string());
        let entry = entries.get(&key).ok_or("没有可用的人工审批，拒绝执行")?;
        if entry.expires <= Instant::now() || entry.tool != tool || &entry.original != original {
            return Err("审批与本次调用不一致或已过期，拒绝执行".into());
        }
        let entry = entries.remove(&key).unwrap();
        Ok((entry.approved, entry.intent))
    }
    pub fn handler(self: &Arc<Self>, fallback: BridgeHandler) -> BridgeHandler {
        let reviewed = self.clone();
        Arc::new(move |path, body| {
            if path != ROUTE {
                return fallback(path, body);
            }
            #[derive(serde::Deserialize)]
            #[serde(rename_all = "camelCase", deny_unknown_fields)]
            struct Request {
                session_id: String,
                call_id: String,
                tool: String,
                original_args: Value,
            }
            let Ok(req) = serde_json::from_str::<Request>(body) else {
                return BridgeReply::error(400, "审批参数请求格式错误");
            };
            if req.tool != "write_test_marker" {
                return BridgeReply::error(403, "业务写入必须在服务端消费审批");
            }
            // Recheck workspace/project permission immediately before consuming.
            let scope = serde_json::json!({"sessionId":req.session_id,"tool":req.tool,"projectId":req.original_args.get("projectId")});
            let permission = fallback(super::project_bindings::AUTHORIZE_ROUTE, &scope.to_string());
            if permission.status != 200 {
                return permission;
            }
            match reviewed.consume(&req.session_id, &req.call_id, &req.tool, &req.original_args) {
                Ok(args) => BridgeReply::ok(serde_json::json!({"args":args}).to_string()),
                Err(e) => BridgeReply::error(403, &e),
            }
        })
    }
}

#[derive(serde::Serialize, Clone, Debug)]
#[serde(rename_all = "camelCase")]
pub struct ReviewField {
    pub key: String,
    pub label: String,
    pub previous_value: Option<String>,
    pub proposed_value: String,
}
#[derive(serde::Serialize, Clone, Debug)]
#[serde(rename_all = "camelCase")]
pub struct WriteIntent {
    pub workspace_id: Option<String>,
    pub state_version: Option<i64>,
    pub project_name: String,
    pub template_name: String,
    pub target_description: String,
    pub fields: Vec<ReviewField>,
}

pub fn marker_intent(args: &Value) -> WriteIntent {
    WriteIntent {
        workspace_id: None,
        state_version: None,
        project_name: "测试操作（不写入项目）".into(),
        template_name: "文本审批演练".into(),
        target_description: "系统临时目录中的新测试标记文件".into(),
        fields: vec![ReviewField {
            key: "note".into(),
            label: "需求分析 / 备注".into(),
            previous_value: None,
            proposed_value: args
                .get("note")
                .and_then(Value::as_str)
                .unwrap_or("")
                .into(),
        }],
    }
}
pub fn validate_edit(tool: &str, original: &Value, amended: &Value) -> Result<(), String> {
    if tool == super::template_write::TOOL {
        return super::template_write::validate_edit(original, amended);
    }
    if tool != "write_test_marker" {
        return Err("此操作尚不支持修改参数，请拒绝后重新发起".into());
    }
    let object = amended.as_object().ok_or("修改后的参数格式错误")?;
    if object.keys().any(|k| k != "note")
        || original
            .as_object()
            .is_none_or(|o| o.keys().any(|k| k != "note"))
    {
        return Err("只能修改本次审批展示的文本字段".into());
    }
    if !object
        .get("note")
        .is_some_and(|v| v.as_str().is_some_and(|s| s.chars().count() <= 20000))
    {
        return Err("备注须为不超过20000字的文本".into());
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn expired_grants_and_revoked_scope_cannot_execute() {
        let ledger = Arc::new(ReviewedArguments::default());
        let key = ("s".into(), "c".into());
        let entry = || Entry {
            intent: None,
            tool: "write_test_marker".into(),
            original: serde_json::json!({}),
            approved: serde_json::json!({"note":"edit"}),
            expires: Instant::now() + Duration::from_secs(60),
        };
        ledger.0.lock().unwrap().insert(
            key.clone(),
            Entry {
                expires: Instant::now() - Duration::from_secs(1),
                ..entry()
            },
        );
        assert!(ledger
            .consume("s", "c", "write_test_marker", &serde_json::json!({}))
            .is_err());
        ledger.0.lock().unwrap().insert(key, entry());
        let handler = ledger.handler(Arc::new(|_, _| BridgeReply::error(403, "项目绑定已撤销")));
        let reply = handler(ROUTE,&serde_json::json!({"sessionId":"s","callId":"c","tool":"write_test_marker","originalArgs":{}}).to_string());
        assert_eq!(reply.status, 403);
        ledger.clear();
        assert!(ledger
            .consume("s", "c", "write_test_marker", &serde_json::json!({}))
            .is_err());
    }
}
