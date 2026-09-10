//! Thin pure-calculator routes: validation and rates belong only to benefit::calculator.
use super::{bridge_server::BridgeReply, project_bindings::ProjectBindings};
use crate::workspace::WorkspaceRuntime;
use serde::Deserialize;
pub const FORWARD_ROUTE: &str = "/lamber-bridge/calculate-selection-fee";
pub const REVERSE_ROUTE: &str = "/lamber-bridge/reverse-calculate-selection-fee";
#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct Forward {
    session_id: String,
    quote: String,
    markup: String,
}
#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct Reverse {
    session_id: String,
    limit: String,
    markup: String,
}
pub fn handle(
    runtime: &WorkspaceRuntime,
    bindings: &ProjectBindings,
    path: &str,
    body: &str,
) -> BridgeReply {
    let result = (|| {
        let (session, value, markup, tool) = if path == FORWARD_ROUTE {
            let r: Forward = serde_json::from_str(body).map_err(|e| format!("参数错误: {e}"))?;
            (r.session_id, r.quote, r.markup, "calculate_selection_fee")
        } else {
            let r: Reverse = serde_json::from_str(body).map_err(|e| format!("参数错误: {e}"))?;
            (
                r.session_id,
                r.limit,
                r.markup,
                "reverse_calculate_selection_fee",
            )
        };
        let (workspace, _) = runtime.require_context()?;
        let cwd = std::fs::canonicalize(&workspace.workspace_root).map_err(|_| "工作区不可用")?;
        bindings.authorize(&session, tool, None, &workspace.workspace_id, &cwd)?;
        if path == FORWARD_ROUTE {
            crate::benefit::calculator::calculate_selection_fee(value, markup)
        } else {
            crate::benefit::calculator::reverse_calculate_selection_fee(value, markup)
        }
    })();
    match result {
        Ok(v) => match serde_json::to_string(&v) {
            Ok(s) => BridgeReply::ok(s),
            Err(e) => BridgeReply::error(500, &e.to_string()),
        },
        Err(e) => BridgeReply::error(422, &e),
    }
}
