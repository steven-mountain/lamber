//! User-triggered long-text review rehearsal. No model tool or project write is added.
use super::approval::{self, ApprovalGate, ApprovalPrompt, ApprovalQuestion};
use super::reviewed_arguments::{marker_intent, validate_edit};
use std::{path::Path, sync::Mutex};
static REHEARSAL: Mutex<()> = Mutex::new(());
#[derive(serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub struct RehearsalReceipt {
    pub approved: bool,
    pub path: String,
    pub saved_text: Option<String>,
    pub reason: String,
}
pub fn run(
    gate: &std::sync::Arc<ApprovalGate>,
    path: &Path,
    proposed: String,
    announce: impl FnOnce(&ApprovalPrompt),
) -> Result<RehearsalReceipt, String> {
    let _exclusive = REHEARSAL
        .try_lock()
        .map_err(|_| "已有审批演练进行中，请先处理")?;
    let args = serde_json::json!({"note": proposed});
    validate_edit("write_test_marker", &args, &args)?;
    let previous = match std::fs::read_to_string(path) {
        Ok(text) => Some(text),
        Err(e) if e.kind() == std::io::ErrorKind::NotFound => None,
        Err(e) => return Err(format!("无法读取演练原文：{e}")),
    };
    let mut intent = marker_intent(&args);
    intent.target_description = path.display().to_string();
    intent.fields[0].previous_value = previous.clone();
    let decision = approval::handle_request(
        gate,
        ApprovalQuestion {
            session_id: None,
            call_id: None,
            tool_name: "write_test_marker".into(),
            reason: Some("长文本审批演练：仅保存到演练文本文件，不写入任何项目或模板。".into()),
            intent: Some(intent),
            args: args.clone(),
        },
        announce,
    );
    let saved_text = if decision.approved {
        let approved = decision.modified_args.as_ref().unwrap_or(&args);
        let text = approved["note"].as_str().ok_or("审批文本格式错误")?;
        if let Some(parent) = path.parent() {
            std::fs::create_dir_all(parent).map_err(|e| e.to_string())?;
        }
        // An interrupted write cannot replace the previously reviewed file with partial content.
        let pending = path.with_extension("pending");
        std::fs::write(&pending, text).map_err(|e| e.to_string())?;
        std::fs::rename(&pending, path).map_err(|e| e.to_string())?;
        Some(std::fs::read_to_string(path).map_err(|e| e.to_string())?)
    } else {
        previous
    };
    Ok(RehearsalReceipt {
        approved: decision.approved,
        path: path.display().to_string(),
        saved_text,
        reason: decision.reason,
    })
}
