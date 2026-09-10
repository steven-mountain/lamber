//! Desktop action ledger. Dispatch is claimed once; reconnect only reads results.
use super::*;
use rusqlite::{params, OptionalExtension};
use serde_json::{json, Value};

pub const REQUEST_EVENT: &str = "lamber-webui-business-request";
const READS: &[&str] = &[
    "context",
    "documents",
    "lists",
    "reverse-project",
    "images",
    "read-image",
    "demand-targets",
    "quote-image",
];
const WRITES: &[&str] = &[
    "template-action",
    "reverse-action",
    "replace-image",
    "upload-demand",
];
pub struct ActionStore(rusqlite::Connection);
impl ActionStore {
    pub fn open(path: &std::path::Path) -> Result<Self, String> {
        let conn = rusqlite::Connection::open(path).map_err(|e| e.to_string())?;
        conn.busy_timeout(std::time::Duration::from_secs(5))
            .map_err(|e| e.to_string())?;
        conn.execute_batch(
            "CREATE TABLE IF NOT EXISTS web_actions (
            id TEXT PRIMARY KEY, session_id TEXT NOT NULL, method TEXT NOT NULL,
            state TEXT NOT NULL, result TEXT, created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP);",
        )
        .map_err(|e| e.to_string())?;
        let has_deadline = conn
            .prepare("PRAGMA table_info(web_actions)")
            .map_err(|e| e.to_string())?
            .query_map([], |row| row.get::<_, String>(1))
            .map_err(|e| e.to_string())?
            .collect::<Result<Vec<_>, _>>()
            .map_err(|e| e.to_string())?
            .iter()
            .any(|name| name == "expires_at");
        if !has_deadline {
            conn.execute(
                "ALTER TABLE web_actions ADD COLUMN expires_at INTEGER NOT NULL DEFAULT 0",
                [],
            )
            .map_err(|e| e.to_string())?;
        }
        Ok(Self(conn))
    }
    pub fn begin(&self, id: &str, session: &str, method: &str) -> Result<i64, String> {
        uuid::Uuid::parse_str(id).map_err(|_| "业务请求身份无效")?;
        if !READS.contains(&method) && !WRITES.contains(&method) {
            return Err("未开放此业务操作".into());
        }
        let expires_at = chrono::Utc::now().timestamp_millis() + 15_000;
        self.0.execute("INSERT INTO web_actions(id,session_id,method,state,expires_at) VALUES (?1,?2,?3,'queued',?4)",params![id,session,method,expires_at])
            .map_err(|_|"业务请求已登记，请读取原请求结果，不能重复执行".to_string())?;
        Ok(expires_at)
    }
    fn expire_queued(&self) -> Result<(), String> {
        self.0.execute("UPDATE web_actions SET state='settled',result=?1 WHERE state='queued' AND expires_at<=?2",
            params![json!({"ok":false,"error":"主窗口未及时接收，操作已取消且未执行。请重新发起。"}).to_string(),chrono::Utc::now().timestamp_millis()]).map_err(|e|e.to_string())?;
        Ok(())
    }
    pub fn claim(&self, id: &str) -> Result<(), String> {
        self.expire_queued()?;
        if self
            .0
            .execute(
                "UPDATE web_actions SET state='running' WHERE id=?1 AND state='queued'",
                [id],
            )
            .map_err(|e| e.to_string())?
            != 1
        {
            return Err("业务请求已处理或失效，不再执行".into());
        }
        Ok(())
    }
    pub fn finish(&self, id: &str, result: &Value) -> Result<(), String> {
        if self.0.execute("UPDATE web_actions SET state='settled',result=?2 WHERE id=?1 AND state IN ('queued','running')",params![id,result.to_string()]).map_err(|e|e.to_string())? != 1 {
            return Err("业务请求已结束，不能覆盖原回执".into());
        }
        Ok(())
    }
    pub fn read(&self, id: &str, session: &str) -> Result<Value, String> {
        self.expire_queued()?;
        let row: Option<(String, Option<String>)> = self
            .0
            .query_row(
                "SELECT state,result FROM web_actions WHERE id=?1 AND session_id=?2",
                params![id, session],
                |r| Ok((r.get(0)?, r.get(1)?)),
            )
            .optional()
            .map_err(|e| e.to_string())?;
        let (state, result) = row.ok_or("未找到此会话的业务请求")?;
        Ok(
            json!({"state":state,"result":result.map(|s|serde_json::from_str::<Value>(&s)).transpose().map_err(|e|e.to_string())?}),
        )
    }
    pub fn receipts(&self, session: &str) -> Result<Value, String> {
        self.expire_queued()?;
        let mut query=self.0.prepare("SELECT id,method,state,result,created_at FROM web_actions WHERE session_id=?1 AND method IN ('template-action','reverse-action','replace-image','upload-demand') ORDER BY created_at DESC,rowid DESC LIMIT 40").map_err(|e|e.to_string())?;
        let rows = query
            .query_map([session], |r| {
                Ok((
                    r.get::<_, String>(0)?,
                    r.get::<_, String>(1)?,
                    r.get::<_, String>(2)?,
                    r.get::<_, Option<String>>(3)?,
                    r.get::<_, String>(4)?,
                ))
            })
            .map_err(|e| e.to_string())?;
        let mut results = Vec::new();
        for row in rows {
            let (id, method, state, result, created) = row.map_err(|e| e.to_string())?;
            let mut result = result
                .map(|s| serde_json::from_str::<Value>(&s))
                .transpose()
                .map_err(|e| e.to_string())?;
            // Preview permits never travel through history/model context as reusable actions.
            if let Some(value) = result
                .as_mut()
                .and_then(|v| v.get_mut("value"))
                .and_then(Value::as_object_mut)
            {
                value.remove("token");
            }
            results.push(json!({"requestId":id,"method":method,"state":state,"result":result,"createdAt":created}));
        }
        Ok(json!(results))
    }
    pub fn recover(&self) -> Result<(), String> {
        self.0.execute("UPDATE web_actions SET state='interrupted', result=?1 WHERE state IN ('queued','running')",[json!({"ok":false,"error":"应用上次退出时此操作尚无最终回执。请核对项目或输出文件；系统不会自动重做。"}).to_string()]).map_err(|e|e.to_string())?;
        // Read responses, including image bytes, are ephemeral and never historical receipts.
        self.0.execute("DELETE FROM web_actions WHERE method NOT IN ('template-action','reverse-action','replace-image','upload-demand')",[]).map_err(|e|e.to_string())?;
        Ok(())
    }
    pub fn forget_read(&self, id: &str) -> Result<(), String> {
        self.0.execute("DELETE FROM web_actions WHERE id=?1 AND state='settled' AND method NOT IN ('template-action','reverse-action','replace-image','upload-demand')",[id]).map_err(|e|e.to_string())?;
        Ok(())
    }
}
pub fn store(app: &tauri::AppHandle) -> Result<ActionStore, String> {
    ActionStore::open(
        &app.path()
            .app_data_dir()
            .map_err(|e| e.to_string())?
            .join("ai-sessions.sqlite"),
    )
}
#[tauri::command]
pub fn ai_webui_claim_action(
    app: tauri::AppHandle,
    window: tauri::WebviewWindow,
    request_id: String,
) -> Result<(), String> {
    if window.label() != "main" {
        return Err("业务操作只能由主窗口执行".into());
    }
    store(&app)?.claim(&request_id)
}
#[tauri::command]
pub fn ai_webui_complete_action(
    app: tauri::AppHandle,
    window: tauri::WebviewWindow,
    request_id: String,
    result: Value,
) -> Result<(), String> {
    if window.label() != "main" {
        return Err("回执只能由主窗口提交".into());
    }
    store(&app)?.finish(&request_id, &result)
}
#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn unclaimed_expired_action_cannot_execute_but_running_action_can_finish() {
        let store = ActionStore::open(std::path::Path::new(":memory:")).unwrap();
        let expired = uuid::Uuid::new_v4().to_string();
        let running = uuid::Uuid::new_v4().to_string();
        store.begin(&expired, "web:a", "replace-image").unwrap();
        store.begin(&running, "web:a", "replace-image").unwrap();
        store.claim(&running).unwrap();
        store
            .0
            .execute("UPDATE web_actions SET expires_at=0", [])
            .unwrap();
        assert_eq!(
            store.read(&expired, "web:a").unwrap()["result"]["ok"],
            false
        );
        assert!(store.claim(&expired).is_err());
        assert!(store.finish(&expired, &json!({"ok":true})).is_err());
        assert_eq!(store.read(&running, "web:a").unwrap()["state"], "running");
        store.finish(&running, &json!({"ok":true})).unwrap();
    }
    #[test]
    fn action_reconnect_duplicate_and_crash_never_replay() {
        let store = ActionStore::open(std::path::Path::new(":memory:")).unwrap();
        let id = uuid::Uuid::new_v4().to_string();
        store.begin(&id, "web:a", "reverse-action").unwrap();
        assert!(store.begin(&id, "web:a", "reverse-action").is_err());
        assert!(store.read(&id, "web:b").is_err());
        store.claim(&id).unwrap();
        assert!(store.claim(&id).is_err());
        store
            .finish(
                &id,
                &json!({"ok":true,"value":{"status":"preview","token":"single-use"}}),
            )
            .unwrap();
        assert!(store.finish(&id, &json!({"ok":false})).is_err());
        assert!(store.receipts("web:a").unwrap()[0]["result"]["value"]
            .get("token")
            .is_none());
        let interrupted = uuid::Uuid::new_v4().to_string();
        store.begin(&interrupted, "web:a", "replace-image").unwrap();
        store.claim(&interrupted).unwrap();
        store.recover().unwrap();
        store.recover().unwrap();
        assert_eq!(
            store.read(&interrupted, "web:a").unwrap()["state"],
            "interrupted"
        );
        assert!(store.claim(&interrupted).is_err());
        assert_eq!(store.read(&id, "web:a").unwrap()["state"], "settled");
    }
}
