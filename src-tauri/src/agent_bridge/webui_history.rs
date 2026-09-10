//! Idempotent, lossless backup of retired frontend history; never an authority source.
use super::*;
use rusqlite::{params, OptionalExtension};
use serde_json::{json, Value};
struct LegacyStore(rusqlite::Connection);
impl LegacyStore {
    fn open(path: &std::path::Path) -> Result<Self, String> {
        let conn = rusqlite::Connection::open(path).map_err(|e| e.to_string())?;
        conn.busy_timeout(std::time::Duration::from_secs(5))
            .map_err(|e| e.to_string())?;
        conn.execute_batch("CREATE TABLE IF NOT EXISTS web_legacy_sources (
            id INTEGER PRIMARY KEY, raw TEXT NOT NULL UNIQUE, created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP);
            CREATE TABLE IF NOT EXISTS web_legacy_sessions (
            session_id TEXT PRIMARY KEY, source_id INTEGER NOT NULL, record TEXT NOT NULL);
            CREATE TABLE IF NOT EXISTS web_legacy_import_status (id INTEGER PRIMARY KEY CHECK(id=1), warning TEXT);").map_err(|e|e.to_string())?;
        Ok(Self(conn))
    }
    fn import(&mut self, raw: &str) -> Result<usize, String> {
        // Preserve even malformed input before indexing. Failed indexing never removes
        // the last valid index or the browser's original source.
        self.0.execute("INSERT OR IGNORE INTO web_legacy_sources(raw) VALUES (?1)", [raw])
            .map_err(|e| e.to_string())?;
        let records = match parse_records(raw) {
            Ok(records) => records,
            Err(error) => {
                self.0.execute("INSERT INTO web_legacy_import_status(id, warning) VALUES (1, ?1) ON CONFLICT(id) DO UPDATE SET warning=excluded.warning", [&error]).map_err(|e|e.to_string())?;
                return Err(error);
            }
        };
        let tx = self.0.transaction().map_err(|e| e.to_string())?;
        tx.execute(
            "INSERT OR IGNORE INTO web_legacy_sources(raw) VALUES (?1)",
            [raw],
        )
        .map_err(|e| e.to_string())?;
        let source_id: i64 = tx
            .query_row(
                "SELECT id FROM web_legacy_sources WHERE raw=?1",
                [raw],
                |r| r.get(0),
            )
            .map_err(|e| e.to_string())?;
        for record in &records {
            tx.execute("INSERT INTO web_legacy_sessions(session_id,source_id,record) VALUES (?1,?2,?3)
                ON CONFLICT(session_id) DO UPDATE SET source_id=excluded.source_id,record=excluded.record",
                params![record["id"].as_str().unwrap(),source_id,record.to_string()]).map_err(|e|e.to_string())?;
        }
        tx.execute("INSERT INTO web_legacy_import_status(id, warning) VALUES (1, NULL) ON CONFLICT(id) DO UPDATE SET warning=NULL", []).map_err(|e| e.to_string())?;
        tx.commit().map_err(|e| e.to_string())?;
        Ok(records.len())
    }
    fn records(&self) -> Result<Vec<Value>, String> {
        let mut query = self
            .0
            .prepare("SELECT record FROM web_legacy_sessions ORDER BY rowid DESC")
            .map_err(|e| e.to_string())?;
        let rows = query
            .query_map([], |r| r.get::<_, String>(0))
            .map_err(|e| e.to_string())?;
        rows.map(|row| {
            serde_json::from_str(&row.map_err(|e| e.to_string())?).map_err(|e| e.to_string())
        })
        .collect()
    }
}
fn parse_records(raw: &str) -> Result<Vec<Value>, String> {
    let source: Value =
        serde_json::from_str(raw).map_err(|_| "旧会话数据无法解析；原浏览器记录已保留")?;
    let records = source["sessions"]
        .as_array()
        .ok_or("旧会话格式不受支持；原记录已保留")?;
    let mut ids = std::collections::HashSet::new();
    for record in records {
        let id = record["id"]
        .as_str()
        .filter(|s| !s.is_empty())
        .ok_or("旧会话缺少身份，未迁移")?;
        if !ids.insert(id) || !record["messages"].is_array() {
        return Err("旧会话身份重复或消息损坏，未迁移".into());
        }
    }
    Ok(records.clone())
}

fn store(app: &tauri::AppHandle) -> Result<LegacyStore, String> {
    LegacyStore::open(
        &app.path()
            .app_data_dir()
            .map_err(|e| e.to_string())?
            .join("ai-sessions.sqlite"),
    )
}
#[tauri::command]
pub fn ai_webui_import_history(
    app: tauri::AppHandle,
    window: tauri::WebviewWindow,
    raw: String,
) -> Result<Value, String> {
    if window.label() != "main" {
        return Err("旧记录只能从本机主窗口迁移".into());
    }
    let imported = store(&app)?.import(&raw);
    match imported {
        Ok(count) => Ok(json!({"imported": count})),
        Err(error) if parse_records(&raw).is_err() => Ok(json!({"imported": 0, "warning": error})),
        Err(error) => Err(error),
    }
}
pub fn status(app: &tauri::AppHandle) -> Result<Value, String> {
    let warning: Option<String> = store(app)?.0.query_row(
        "SELECT warning FROM web_legacy_import_status WHERE id=1", [], |row| row.get(0)
    ).optional().map_err(|e| e.to_string())?.flatten();
    Ok(json!({"warning": warning}))
}
pub fn list(
    app: &tauri::AppHandle,
    runtime: &crate::workspace::WorkspaceRuntime,
) -> Result<Value, String> {
    let workspace = runtime.require_workspace()?;
    let cwd = std::fs::canonicalize(&workspace.workspace_root).map_err(|e| e.to_string())?;
    let identities = open_session_store(app)?;
    let mut result = Vec::new();
    for record in store(app)?.records()? {
        let id = record["id"].as_str().ok_or("旧记录索引损坏")?;
        let binding = identities.binding(id)?;
        if binding.as_ref().is_some_and(|binding| {
            binding.workspace_id != workspace.workspace_id || binding.cwd != cwd
        }) {
            continue;
        }
        let harness = if binding.is_some() {
            identities.get(id, &cwd)?
        } else {
            None
        };
        result.push(json!({"id":id,"title":record["title"],"updatedAt":record["updatedAt"],"harnessSessionId":harness,
            "messageCount":record["messages"].as_array().map(Vec::len).unwrap_or(0),
            "status":if harness.is_some(){"有可信映射，可尝试恢复原会话"}else{"只读旧记录；没有可信映射，请另建会话"}}));
    }
    Ok(json!(result))
}
pub fn read(
    app: &tauri::AppHandle,
    runtime: &crate::workspace::WorkspaceRuntime,
    id: &str,
) -> Result<Value, String> {
    if !list(app, runtime)?
        .as_array()
        .unwrap()
        .iter()
        .any(|record| record["id"] == id)
    {
        return Err("旧记录不属于当前工作区或已不存在".into());
    }
    store(app)?
        .records()?
        .into_iter()
        .find(|record| record["id"] == id)
        .ok_or("旧记录不存在".into())
}
pub fn receipt_context(app: &tauri::AppHandle, harness: &str) -> Result<Value, String> {
    let conn = store(app)?.0;
    let raw:Option<String>=conn.query_row("SELECT l.record FROM web_legacy_sessions l JOIN sessions s ON l.session_id=s.session_id WHERE s.harness_id=?1",[harness],|r|r.get(0)).optional().map_err(|e|e.to_string())?;
    let record: Value = raw
        .map(|s| serde_json::from_str(&s))
        .transpose()
        .map_err(|e| e.to_string())?
        .unwrap_or(Value::Null);
    Ok(json!(record["messages"]
        .as_array()
        .into_iter()
        .flatten()
        .filter(|m| m["appReceipt"] == true)
        .rev()
        .take(8)
        .map(|m| m["content"].clone())
        .collect::<Vec<_>>()))
}
#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn legacy_import_is_lossless_repeatable_and_rejects_partial_duplicate_import() {
        let mut db = LegacyStore::open(std::path::Path::new(":memory:")).unwrap();
        let raw = r#"{"version":1,"sessions":[{"id":"a","title":"旧记录🙂","messages":[{"role":"assistant","content":"已生成","appReceipt":true}]}]}"#;
        assert_eq!(db.import(raw).unwrap(), 1);
        assert_eq!(db.import(raw).unwrap(), 1);
        assert_eq!(
            db.0.query_row("SELECT count(*) FROM web_legacy_sources", [], |r| r
                .get::<_, i64>(0))
                .unwrap(),
            1
        );
        assert_eq!(
            db.0.query_row("SELECT raw FROM web_legacy_sources", [], |r| r
                .get::<_, String>(0))
                .unwrap(),
            raw
        );
        assert!(db
            .import(r#"{"sessions":[{"id":"b","messages":[]},{"id":"b","messages":[]}]}"#)
            .is_err());
        assert_eq!(db.records().unwrap().len(), 1);
        assert_eq!(db.records().unwrap()[0]["messages"][0]["appReceipt"], true);
        let malformed = "{broken JSON";
        assert!(db.import(malformed).is_err());
        assert_eq!(db.0.query_row("SELECT raw FROM web_legacy_sources ORDER BY id DESC LIMIT 1", [], |r|r.get::<_, String>(0)).unwrap(), malformed);
        assert_eq!(db.records().unwrap().len(), 1);
        assert!(db.0.query_row("SELECT warning IS NOT NULL FROM web_legacy_import_status WHERE id=1", [], |r|r.get::<_, bool>(0)).unwrap());
        assert_eq!(db.import(raw).unwrap(), 1);
        assert!(db.0.query_row("SELECT warning IS NULL FROM web_legacy_import_status WHERE id=1", [], |r|r.get::<_, bool>(0)).unwrap());
    }
}
