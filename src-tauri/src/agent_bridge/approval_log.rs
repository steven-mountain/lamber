//! Durable audit trail for agent tool approvals.
//!
//! Every settled question — granted, refused, timed out, or cut short by a
//! shutdown — lands in the workspace database's `agent_approval_log` table, so
//! "who approved which tool call, and when" survives the process that decided
//! it. The table is append-only and independent of any project, because an
//! approval can happen with no project open.
//!
//! Recording must never change a decision, and a decision must never be lost.
//! Those two rules pull in opposite directions when no workspace is open, so
//! there is a spool: a decision that cannot reach the database is appended to a
//! JSONL file in the app data directory, and drained into `agent_approval_log`
//! the next time a workspace is activated.
//!
//! Blocking the approval until a workspace exists was the alternative, and it is
//! wrong here: the dialog cannot open a workspace, so the agent's turn would
//! hang on a condition the user cannot resolve from where they are standing —
//! against the same never-hang rule the timeout path already enforces. An
//! approval also need not concern a project at all (`write_test_marker` does
//! not), so demanding a workspace would gate decisions on something unrelated
//! to them.
//!
//! A genuinely unwritable spool (a full or read-only disk) still degrades to a
//! stderr warning rather than failing the decision. That is the one remaining
//! loss path, and it is loud.

use super::approval::ApprovalRecord;
use std::io::{BufRead, BufReader, Write};
use std::path::{Path, PathBuf};
use std::sync::Arc;

/// File holding decisions taken while no workspace was open.
pub const SPOOL_FILE: &str = "agent-approval-spool.jsonl";

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
    insert_with(&guard, record)
}

/// Append one settled question through any connection-like handle.
fn insert_with(conn: &rusqlite::Connection, record: &ApprovalRecord) -> Result<(), String> {
    conn
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

// --------------------------------------------------------------- the spool --

/// Path of the spool inside an app data directory.
pub fn spool_path_in(app_data_dir: &Path) -> PathBuf {
    app_data_dir.join(SPOOL_FILE)
}

/// Resolve the spool path for the running application.
pub fn spool_path(app: &tauri::AppHandle) -> Result<PathBuf, String> {
    use tauri::Manager;
    let dir = app
        .path()
        .app_data_dir()
        .map_err(|e| format!("无法定位应用数据目录: {e}"))?;
    std::fs::create_dir_all(&dir).map_err(|e| format!("无法创建应用数据目录: {e}"))?;
    Ok(spool_path_in(&dir))
}

/// Append one decision to the spool, to be drained when a workspace opens.
///
/// One JSON object per line, appended under `O_APPEND`, so a crash mid-write
/// can cost at most the trailing line — which `drain_spool` then skips.
pub fn append_to_spool(path: &Path, record: &ApprovalRecord) -> Result<(), String> {
    if let Some(parent) = path.parent() {
        std::fs::create_dir_all(parent).map_err(|e| format!("无法创建缓冲目录: {e}"))?;
    }
    let line = serde_json::to_string(record).map_err(|e| e.to_string())?;
    let mut file = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(path)
        .map_err(|e| format!("无法打开审批缓冲文件: {e}"))?;
    writeln!(file, "{line}").map_err(|e| format!("写入审批缓冲失败: {e}"))?;
    Ok(())
}

/// How many decisions are waiting in the spool. Diagnostics and tests.
pub fn spool_len(path: &Path) -> usize {
    let Ok(file) = std::fs::File::open(path) else {
        return 0;
    };
    BufReader::new(file)
        .lines()
        .map_while(Result::ok)
        .filter(|l| !l.trim().is_empty())
        .count()
}

/// Move every spooled decision into the workspace database.
///
/// Applied in one transaction and only then is the spool removed, so an
/// interrupted drain replays rather than loses. Rows are keyed by `request_id`
/// and written with `INSERT OR REPLACE`, so replaying is idempotent.
///
/// @param path - the spool file; a missing file drains zero.
/// @param conn - the freshly opened workspace connection.
/// @returns how many decisions were moved.
pub fn drain_spool(
    path: &Path,
    conn: &std::sync::Mutex<rusqlite::Connection>,
) -> Result<usize, String> {
    if !path.exists() {
        return Ok(0);
    }
    let file = std::fs::File::open(path).map_err(|e| format!("读取审批缓冲失败: {e}"))?;
    let mut records = Vec::new();
    for line in BufReader::new(file).lines().map_while(Result::ok) {
        let trimmed = line.trim();
        if trimmed.is_empty() {
            continue;
        }
        match serde_json::from_str::<ApprovalRecord>(trimmed) {
            Ok(record) => records.push(record),
            // A torn trailing line from a crash mid-append; the rest still drain.
            Err(e) => eprintln!("[agent_bridge] 跳过损坏的审批缓冲行: {e}"),
        }
    }

    if records.is_empty() {
        let _ = std::fs::remove_file(path);
        return Ok(0);
    }

    {
        let mut guard = conn.lock().map_err(|_| "数据库锁已中毒".to_string())?;
        let tx = guard
            .transaction()
            .map_err(|e| format!("开启审批回填事务失败: {e}"))?;
        for record in &records {
            insert_with(&tx, record)?;
        }
        tx.commit()
            .map_err(|e| format!("提交审批回填事务失败: {e}"))?;
    }

    std::fs::remove_file(path)
        .map_err(|e| format!("回填后清理审批缓冲失败: {e}"))?;
    Ok(records.len())
}

/// Drain the spool into a newly activated workspace, reporting but not failing.
///
/// Called from `workspace::open_workspace_internal`, the single point where a
/// database becomes available.
pub fn drain_spool_on_workspace_open(
    app: &tauri::AppHandle,
    conn: &std::sync::Mutex<rusqlite::Connection>,
) {
    let outcome = spool_path(app).and_then(|path| drain_spool(&path, conn));
    match outcome {
        Ok(0) => {}
        Ok(n) => eprintln!("[agent_bridge] 已将 {n} 条离线审批记录回填到工作区审计日志"),
        Err(e) => eprintln!("[agent_bridge] 回填离线审批记录失败（缓冲保留，下次重试）: {e}"),
    }
}

// ------------------------------------------------------------ the recorder --

/// Build a recorder that persists every decision, workspace open or not.
///
/// Resolving the connection lazily (rather than capturing one) means a workspace
/// opened after the agent started still gets its approvals recorded directly.
/// With no workspace, the decision goes to the spool instead of being dropped.
///
/// @param runtime - the workspace runtime holding the current database.
/// @param spool - where to buffer decisions taken with no workspace open.
pub fn workspace_recorder(
    runtime: Arc<crate::workspace::WorkspaceRuntime>,
    spool: PathBuf,
) -> super::approval::ApprovalRecorder {
    Arc::new(move |record| {
        let direct = runtime.require_db().and_then(|conn| insert(&conn, record));
        let Err(reason) = direct else {
            return;
        };
        // No workspace, or the workspace write failed: buffer rather than drop.
        if let Err(e) = append_to_spool(&spool, record) {
            eprintln!(
                "[agent_bridge] 审批记录既未落库也未能缓冲（决定本身不受影响）: {reason} / {e} · request_id={}",
                record.request_id
            );
        } else {
            eprintln!(
                "[agent_bridge] 审批记录暂存至缓冲，待工作区打开后回填: {reason} · request_id={}",
                record.request_id
            );
        }
    })
}
