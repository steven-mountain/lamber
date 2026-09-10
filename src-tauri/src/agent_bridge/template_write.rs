//! Approved text patches through the existing template state save transaction.
use super::{
    approval::{self, ApprovalDecision, ApprovalGate, ApprovalPrompt, ApprovalQuestion},
    bridge_server::{BridgeHandler, BridgeReply},
    project_bindings::ProjectBindings,
    reviewed_arguments::{ReviewField, WriteIntent},
};
use crate::{
    project_state::{self, TemplateStatePayload},
    workspace::{CurrentWorkspace, WorkspaceRuntime},
};
use serde::Deserialize;
use serde_json::{json, Value};
use std::{collections::BTreeMap, sync::Arc};
pub const TOOL: &str = "fill_template_fields";
pub const ROUTE: &str = "/lamber-bridge/fill-template-fields";
pub const CHANGED_EVENT: &str = "lamber-template-text-changed";
#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct Arguments {
    pub project_id: String,
    pub template_id: String,
    pub fields: BTreeMap<String, String>,
}
pub fn arguments(value: &Value) -> Result<Arguments, String> {
    let args: Arguments = serde_json::from_value(value.clone())
        .map_err(|_| "仅接受projectId、templateId及文本fields；不接受额外参数或非文本字段")?;
    if args.project_id.trim().is_empty() {
        return Err("必须指定绑定项目".into());
    }
    let catalog = &super::template_catalog::resolve(&args.template_id, false)?.fields;
    let max_fields = catalog.iter().filter(|r| r.kind == "text").count();
    if args.fields.is_empty() || args.fields.len() > max_fields {
        return Err(format!("请提供1至{max_fields}个目录文本字段"));
    }
    for (key, value) in &args.fields {
        if !catalog.iter().any(|r| r.key == *key && r.kind == "text") {
            return Err(format!("字段 {key} 不在文本白名单中；金额、税率、年限、折现率、测算结果、清单及图片均禁止写入"));
        }
        if value.trim().is_empty() || value.chars().count() > 20000 {
            return Err(format!("字段 {key} 必须为1至20000字文本"));
        }
    }
    Ok(args)
}
pub fn validate_edit(original: &Value, amended: &Value) -> Result<(), String> {
    let old = arguments(original)?;
    let new = arguments(amended)?;
    if old.project_id != new.project_id
        || old.template_id != new.template_id
        || old.fields.keys().ne(new.fields.keys())
    {
        return Err("只能修改本次审批展示的字段值，不能变更项目、模板或字段集合".into());
    }
    Ok(())
}
fn authorize(
    bindings: &ProjectBindings,
    workspace: &CurrentWorkspace,
    session: &str,
    project: &str,
) -> Result<(), String> {
    let cwd = std::fs::canonicalize(&workspace.workspace_root).map_err(|_| "工作区不可用")?;
    bindings.authorize(session, TOOL, Some(project), &workspace.workspace_id, &cwd)?;
    Ok(())
}
pub fn prepare(
    runtime: &WorkspaceRuntime,
    bindings: &ProjectBindings,
    session: &str,
    value: &Value,
) -> Result<WriteIntent, String> {
    let args = arguments(value)?;
    runtime.with_locked_context(|workspace, conn| {
        authorize(bindings, workspace, session, &args.project_id)?;
        let project =
            project_state::get_project_locked(conn, &args.project_id)?.ok_or("绑定项目不存在")?;
        let saved =
            project_state::get_template_state_locked(conn, &args.project_id, &args.template_id)?;
        let version = saved
            .as_ref()
            .map(project_state::template_revision)
            .unwrap_or(0);
        let catalog = &super::template_catalog::resolve(&args.template_id, false)?.fields;
        let fields = args
            .fields
            .iter()
            .map(|(key, proposed)| {
                let rule = catalog.iter().find(|r| r.key == *key).unwrap();
                let previous = rule.value(saved.as_ref().map(|s| &s.filled_data_json).unwrap_or(&Value::Null));
                ReviewField {
                    key: key.clone(),
                    label: rule.label.clone(),
                    previous_value: previous,
                    proposed_value: proposed.clone(),
                }
            })
            .collect();
        Ok(WriteIntent {
            workspace_id: Some(workspace.workspace_id.clone()),
            state_version: Some(version),
            project_name: project.name,
            template_name: args.template_id.clone(),
            target_description: format!("模板已保存正文 · {}", args.template_id),
            fields,
        })
    })
}
pub fn request_approval(
    gate: &Arc<ApprovalGate>,
    runtime: &WorkspaceRuntime,
    bindings: &ProjectBindings,
    mut question: ApprovalQuestion,
    announce: impl FnOnce(&ApprovalPrompt),
) -> ApprovalDecision {
    if question.tool_name == TOOL {
        let result = question
            .session_id
            .as_deref()
            .ok_or_else(|| "缺少可信会话身份".to_string())
            .and_then(|session| prepare(runtime, bindings, session, &question.args));
        match result {
            Ok(intent) => question.intent = Some(intent),
            Err(reason) => {
                return ApprovalDecision {
                    approved: false,
                    modified_args: None,
                    reason,
                }
            }
        }
    }
    approval::handle_request(gate, question, announce)
}
/// Consume permission inside the server, so posting arbitrary edited args cannot bypass approval.
pub fn handler(
    runtime: Arc<WorkspaceRuntime>,
    bindings: Arc<ProjectBindings>,
    gate: Arc<ApprovalGate>,
    changed: Arc<dyn Fn(Value) + Send + Sync>,
    fallback: BridgeHandler,
) -> BridgeHandler {
    Arc::new(move |path, body| {
        if path != ROUTE {
            return fallback(path, body);
        }
        #[derive(Deserialize)]
        #[serde(rename_all = "camelCase", deny_unknown_fields)]
        struct Request {
            session_id: String,
            call_id: String,
            original_args: Value,
        }
        let result = (|| {
            let req: Request = serde_json::from_str(body).map_err(|_| "模板写入请求格式错误")?;
            let original = arguments(&req.original_args)?;
            runtime.with_locked_context(|workspace,conn| {
                authorize(&bindings,workspace,&req.session_id,&original.project_id)?;
                let (approved,intent)=gate.reviewed.take(&req.session_id,&req.call_id,TOOL,&req.original_args)?;
                validate_edit(&req.original_args,&approved)?;
                let args=arguments(&approved)?;
                let intent=intent.ok_or("缺少可核验的写入意图")?;
                if intent.workspace_id.as_deref()!=Some(&workspace.workspace_id) {return Err("审批所属工作区已改变，请重新发起".into());}
                let version=intent.state_version.ok_or("缺少审批版本")?;
                let saved=project_state::get_template_state_locked(conn,&args.project_id,&args.template_id)?;
                if saved.as_ref().map(project_state::template_revision).unwrap_or(0)!=version {return Err("TemplateStateConflict::审批期间模板已改变，请基于最新内容重新审批".into());}
                // Legacy-only rows have no revision counter: still reject changes to reviewed fields.
                let catalog = &super::template_catalog::resolve(&args.template_id, false)?.fields;
                for reviewed in &intent.fields {
                    let rule = catalog.iter().find(|rule| rule.key == reviewed.key).ok_or("审批字段已不受支持")?;
                    let current = rule.value(saved.as_ref().map(|s| &s.filled_data_json).unwrap_or(&Value::Null));
                    if current != reviewed.previous_value { return Err("审批字段原值已改变，请重新审批".into()); }
                }
                let mut payload=match saved {
                    Some(s)=>TemplateStatePayload{template_name:s.template_name,template_type:s.template_type,template_path:s.template_path,template_path_type:s.template_path_type,filled_data_json:s.filled_data_json,field_mapping_json:s.field_mapping_json,output_config_json:s.output_config_json},
                    None=>TemplateStatePayload{template_name:Some(args.template_id.clone()),template_type:Some("word".into()),template_path:Some(args.template_id.clone()),template_path_type:Some("module".into()),filled_data_json:json!({"formData":{}}),field_mapping_json:json!({}),output_config_json:json!({})}
                };
                for (key,value) in &args.fields {
                    catalog.iter().find(|r| r.key == *key && r.kind == "text").ok_or("字段不在文本目录")?.write(&mut payload.filled_data_json,json!(value))?;
                }
                let saved=project_state::save_template_state_locked(conn,args.project_id.clone(),args.template_id.clone(),payload,Some(version))?;
                Ok(json!({"workspaceId":workspace.workspace_id,"projectId":args.project_id,"templateId":args.template_id,"fields":args.fields,"templateVersion":saved.template_version,"updatedAt":saved.updated_at,"message":"已按人工批准内容保存模板文本"}))
            })
        })();
        match result {
            Ok(receipt) => {
                changed(receipt.clone());
                BridgeReply::ok(receipt.to_string())
            }
            Err(e) => BridgeReply::error(403, &e),
        }
    })
}
