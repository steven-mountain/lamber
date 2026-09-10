//! Bound-project saved template projection. No writes and no filesystem paths in the DTO.
use super::{bridge_server::BridgeReply, project_bindings::ProjectBindings};
use crate::{project_state, workspace::WorkspaceRuntime};
use serde::Deserialize;
use serde_json::{json, Map, Value};
pub const TOOL: &str = "read_template_fields";
pub const ROUTE: &str = "/lamber-bridge/read-template-fields";
pub const TEXT_LIMIT: usize = 24000;
const ROW_LIMIT: usize = 100;
#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct Request {
    session_id: String,
    template_id: String,
}
struct Budget {
    original: usize,
    returned: usize,
}
impl Budget {
    fn text(&mut self, value: &str) -> String {
        self.original += value.chars().count();
        let result: String = value
            .chars()
            .take(TEXT_LIMIT.saturating_sub(self.returned))
            .collect();
        self.returned += result.chars().count();
        result
    }
}
pub fn handle(runtime: &WorkspaceRuntime, bindings: &ProjectBindings, body: &str) -> BridgeReply {
    let request: Request = match serde_json::from_str(body) {
        Ok(value) => value,
        Err(_) => {
            return BridgeReply::error(
                400,
                "只接受 templateId；项目取自可信会话绑定，不接受 projectId 或额外参数",
            )
        }
    };
    let result = runtime.with_locked_context(|workspace, conn| {
        let cwd = std::fs::canonicalize(&workspace.workspace_root).map_err(|_| "工作区不可用")?;
        let scope = bindings.authorize(&request.session_id, TOOL, None, &workspace.workspace_id, &cwd)?;
        let project_id = scope.bound_project_id().ok_or("要求绑定项目")?;
        let project = project_state::get_project_locked(conn, project_id)?.ok_or("绑定项目已不存在或不可访问")?;
        let template = super::template_catalog::resolve(&request.template_id, true)?;
        let template_id = if request.template_id == template.name {
            let names: Vec<_> = project_state::list_template_states_locked(conn, project_id)?.into_iter()
                .map(|s| s.template_id).filter(|s| super::template_catalog::resolve(s, false).map(|t| t.id == template.id).unwrap_or(false)).collect();
            if names.len() != 1 { return Err(format!("请指定完整模板名称；本项目已保存的候选：{}", names.join("、"))); }
            names[0].clone()
        } else { request.template_id.clone() };
        let saved = project_state::get_template_state_locked(conn, project_id, &template_id)?;
        let empty = json!({});
        let state = saved.as_ref().map(|s| &s.filled_data_json).unwrap_or(&empty);
        let catalog = &template.fields;
        let mut budget = Budget { original: 0, returned: 0 };
        let mut completion_state = json!({"formData":{}});
        let mut fields = Vec::new();
        for rule in catalog.iter().filter(|r| r.kind == "text") {
            let raw = rule.raw(state);
            // Only presence reaches the pure completion function, independently of truncation.
            if let Some(value) = raw { rule.write(&mut completion_state, json!(if value.trim().is_empty() { "" } else { "filled" }))?; }
            let value = raw.or(rule.default_value.as_deref()).unwrap_or("");
            let returned = budget.text(value);
            fields.push(json!({"key":rule.key,"label":rule.label,"value":returned,
                "valueSource":if raw.is_some() {"saved"} else if rule.default_value.is_some() {"default"} else if rule.dynamic_default == Some(true) {"dynamic_default_unavailable"} else {"empty"},
                "originalCharacters":value.chars().count(),"truncated":returned.chars().count() < value.chars().count()}));
        }
        for rule in catalog {
            if let Some(condition) = &rule.required_when {
                let key = condition.field();
                if let Some(value) = state.get(key).filter(|v| v.is_string() || v.is_boolean() || v.is_number()) {
                    completion_state[key] = value.clone();
                }
            }
            for source in rule.completion_sources.iter().flatten() {
                let input = if source.source.as_deref() == Some("formData") { state.get("formData").unwrap_or(&empty) } else { state };
                if let Some(value) = input.get(&source.field).and_then(Value::as_str) {
                    let output = if source.source.as_deref() == Some("formData") { &mut completion_state["formData"] } else { &mut completion_state };
                    let used_in_condition = source.source.is_none() && catalog.iter().any(|field| field.required_when.as_ref().map(|c| c.field() == source.field).unwrap_or(false));
                    output[&source.field] = json!(if used_in_condition { value } else if value.trim().is_empty() { "" } else { "filled" });
                }
            }
        }
        let mut lists = Map::new();
        let mut row_counts = Map::new();
        let mut returned_counts = Map::new();
        let mut rows_truncated = false;
        for rule in catalog.iter().filter(|r| r.kind == "list") {
            let rows = state.get(&rule.key).and_then(Value::as_array).cloned().unwrap_or_default();
            let mut items = Vec::new();
            for (index, row) in rows.iter().enumerate() {
                let mut item = Map::new();
                for key in rule.columns.iter().flatten() {
                    let raw = match row.get(key) { Some(Value::String(s)) => s.clone(), Some(Value::Number(n)) => n.to_string(), _ => String::new() };
                    if index < ROW_LIMIT { item.insert(key.clone(), json!(budget.text(&raw))); }
                    else { budget.original += raw.chars().count(); }
                }
                if index < ROW_LIMIT { items.push(Value::Object(item)); }
            }
            completion_state[&rule.key] = if let Some(valid) = &rule.valid_row {
                // Lossless predicate-input projection. Deduplicate presence/sign patterns across ALL
                // saved rows; visible row truncation must never change completion. No names/images leak.
                let mut signatures = std::collections::BTreeSet::new();
                let mut evidence = Vec::new();
                for row in &rows {
                    let mut projected = Map::new();
                    for key in &valid.non_empty {
                        let value = row.get(key).and_then(Value::as_str).unwrap_or("");
                        projected.insert(key.clone(), json!(if value.trim().is_empty() { "" } else { "filled" }));
                    }
                    for key in &valid.positive {
                        let value = row.get(key).and_then(|v| v.as_f64().or_else(|| v.as_str().and_then(|s| s.trim().parse::<f64>().ok()))).unwrap_or(0.0);
                        projected.insert(key.clone(), json!(if value > 0.0 {1} else {0}));
                    }
                    let signature = serde_json::to_string(&projected).map_err(|e| e.to_string())?;
                    if signatures.insert(signature) { evidence.push(Value::Object(projected)); }
                }
                json!(evidence)
            } else if rows.is_empty() {json!([])} else {json!([{}])};
            rows_truncated |= rows.len() > ROW_LIMIT;
            row_counts.insert(rule.key.clone(), json!(rows.len()));
            returned_counts.insert(rule.key.clone(), json!(items.len()));
            lists.insert(rule.key.clone(), json!(items));
        }
        let assets = project_state::list_template_assets_locked(conn, project_id, Some(&template_id))?;
        let attachment_status: Vec<_> = catalog.iter().filter(|r| r.kind == "image").map(|rule| {
            let present = assets.iter().any(|a| a.usage.as_deref() == Some(&rule.key) &&
                crate::ai_context::service::safe_workspace_asset_exists(std::path::Path::new(&workspace.workspace_root), &a.relative_path) == Some(true));
            json!({"fieldKey":rule.key,"exists":present})
        }).collect();
        let has_public_url = state.get("hasPublicUrl").and_then(Value::as_bool).unwrap_or(false);
        let has_security = state.get("hasSecurity").and_then(Value::as_bool).unwrap_or(false);
        let truncated = budget.original > budget.returned || rows_truncated;
        Ok(json!({"projectId":project_id,"projectName":project.name,"templateId":template_id,
            "source":saved.as_ref().map(|s| s.source.as_str()).unwrap_or("no_saved_state"),
            "hasSavedState":saved.is_some(),"templateVersion":saved.as_ref().map(project_state::template_revision).unwrap_or(0),
            "fields":fields,"lists":lists,"listCounts":row_counts,"returnedListCounts":returned_counts,"techItems":lists.get("techItems").cloned().unwrap_or(json!([])),"hasPublicUrl":has_public_url,"hasSecurity":has_security,"attachments":attachment_status,
            "truncated":truncated,"textLimit":TEXT_LIMIT,"originalCharacters":budget.original,"returnedCharacters":budget.returned,
            "techItemCount":row_counts.get("techItems").unwrap_or(&json!(0)),"returnedTechItemCount":returned_counts.get("techItems").unwrap_or(&json!(0)),
            "notice":if truncated {"内容超过总文本24000字或清单100行上限，已明确截断；缺项按完整保存态计算，不能将未返回内容判为空。"} else {"读取绑定项目已保存状态；未保存草稿不在结果内。valueSource=default 表示界面默认值，并非用户已保存填写。"},
            "catalogId":template.id,
            "readOnlyFields":catalog.iter().filter(|r| r.kind != "text").map(|r|json!({"key":r.key,"label":r.label,"kind":r.kind,"reason":r.reason,"completionAlways":r.completion_always,"listType":r.list_type})).collect::<Vec<_>>(),
            "completionState":completion_state
        }))
    });
    match result {
        Ok(value) => BridgeReply::ok(value.to_string()),
        Err(error) => BridgeReply::error(403, &error),
    }
}
