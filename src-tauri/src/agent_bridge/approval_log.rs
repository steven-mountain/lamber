//! Durable audit trail for agent tool approvals.
//!
//! Every settled question — granted, refused, timed out, or cut short by a
//! shutdown — lands in the workspace database's `agent_approval_log` table, so
//! "who approved which tool call, and when" survives the process that decided
//! it. The table is append-only and independent of any project, because an
//! approval can happen with no project open.
//!
//! Recording must never change a decision: a closed workspace or a failing
//! write is logged to stderr and swallowed. Failing an approval because the
//! audit write failed would turn a storage hiccup into a blocked agent, and
//! failing *open* would be worse still.

use super::approval::ApprovalRecord;
use std::sync::Arc;

/// One row of the audit trail, as read back for inspection.
#[derive(serde::Serialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ApprovalLogEntry {
    pub request_id: String,
    pub tool_name: String,
    pub call_id: Option<String>,
    pub reason: Option<String>,
    pub args_json: String,
    pub approved: bool,
    pub decided_by: String,
    pub decision_reason: String,
    pub requested_at: String,
    pub decided_at: String,
}

/// Append one settled question to an open workspace database.
///
/// @param conn - the workspace connection.
/// @param record - the settled question.
pub fn insert(
    conn: &std::sync::Mutex<rusqlite::Connection>,
    record: &ApprovalRecord,
) -> Result<(), String> {
    let guard = conn.lock().map_err(|_| "数据库锁已中毒".to_string())?;
    guard
        .execute(
            "INSERT OR REPLACE INTO agent_approval_log
                (request_id, tool_name, call_id, reason, args_json, approved,
                 decided_by, decision_reason, requested_at, decided_at)
             VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10)",
            rusqlite::params![
                record.request_id,
                record.tool_name,
                record.call_id,
                record.reason,
                record.args_json,
                record.approved as i32,
                record.decided_by.as_str(),
                record.decision_reason,
                record.requested_at,
                record.decided_at,
            ],
        )
        .map_err(|e| format!("写入审批日志失败: {e}"))?;
    Ok(())
}

/// Read the most recent audit entries, newest first.
///
/// @param conn - the workspace connection.
/// @param limit - maximum rows to return.
pub fn recent(
    conn: &std::sync::Mutex<rusqlite::Connection>,
    limit: u32,
) -> Result<Vec<ApprovalLogEntry>, String> {
    let guard = conn.lock().map_err(|_| "数据库锁已中毒".to_string())?;
    let mut stmt = guard
        .prepare(
            "SELECT request_id, tool_name, call_id, reason, args_json, approved,
                    decided_by, decision_reason, requested_at, decided_at
             FROM agent_approval_log
             ORDER BY decided_at DESC, rowid DESC
             LIMIT ?1",
        )
        .map_err(|e| e.to_string())?;
    let rows = stmt
        .query_map([limit], |row| {
            Ok(ApprovalLogEntry {
                request_id: row.get(0)?,
                tool_name: row.get(1)?,
                call_id: row.get(2)?,
                reason: row.get(3)?,
                args_json: row.get(4)?,
                approved: row.get::<_, i32>(5)? != 0,
                decided_by: row.get(6)?,
                decision_reason: row.get(7)?,
                requested_at: row.get(8)?,
                decided_at: row.get(9)?,
            })
        })
        .map_err(|e| e.to_string())?;
    rows.collect::<Result<Vec<_>, _>>()
        .map_err(|e| e.to_string())
}

/// Build a recorder that appends to whatever workspace is open at decision time.
///
/// Resolving the connection lazily (rather than capturing one) means a workspace
/// opened after the agent started still gets its approvals recorded, and a
/// closed workspace degrades to a warning instead of a panic.
///
/// @param runtime - the workspace runtime holding the current database.
pub fn workspace_recorder(
    runtime: Arc<crate::workspace::WorkspaceRuntime>,
) -> super::approval::ApprovalRecorder {
    Arc::new(move |record| {
        let outcome = runtime
            .require_db()
            .and_then(|conn| insert(&conn, record));
        if let Err(e) = outcome {
            eprintln!(
                "[agent_bridge] 审批日志未能落库（决定本身不受影响）: {e} · request_id={}",
                record.request_id
            );
        }
    })
}
