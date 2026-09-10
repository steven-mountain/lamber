//! Durable ACP identity, separate from project/business databases.
use super::project_bindings::ProjectBinding;
use rusqlite::{params, Connection, OptionalExtension};
use std::path::Path;

pub struct SessionStore(Connection);
impl SessionStore {
    pub fn open(path: &Path) -> Result<Self, String> {
        let conn = Connection::open(path).map_err(|e| format!("无法打开 AI 会话映射: {e}"))?;
        conn.busy_timeout(std::time::Duration::from_secs(5)).map_err(|e|e.to_string())?;
        conn.execute_batch(
            "CREATE TABLE IF NOT EXISTS sessions (
            session_id TEXT PRIMARY KEY, harness_id TEXT NOT NULL UNIQUE, cwd TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS project_bindings (session_id TEXT PRIMARY KEY, binding TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS web_selected_sessions (workspace_id TEXT PRIMARY KEY, session_id TEXT NOT NULL);",
        )
        .map_err(|e| format!("无法初始化 AI 会话映射: {e}"))?;
        Ok(Self(conn))
    }
    pub fn selected_web_session(&self, workspace:&str)->Result<Option<String>,String> {
        self.0.query_row("SELECT session_id FROM web_selected_sessions WHERE workspace_id=?1",[workspace],|row|row.get(0)).optional().map_err(|e|e.to_string())
    }
    pub fn select_web_session(&self, workspace:&str, session:&str)->Result<(),String> {
        if session.is_empty() { return Err("缺少会话身份".into()); }
        self.0.execute("INSERT INTO web_selected_sessions VALUES (?1,?2) ON CONFLICT(workspace_id) DO UPDATE SET session_id=excluded.session_id",params![workspace,session]).map_err(|e|e.to_string())?;
        Ok(())
    }
    pub fn binding(&self, session: &str) -> Result<Option<ProjectBinding>, String> {
        let value: Option<String> = self
            .0
            .query_row(
                "SELECT binding FROM project_bindings WHERE session_id = ?1",
                [session],
                |row| row.get(0),
            )
            .optional()
            .map_err(|e| e.to_string())?;
        value
            .map(|json| serde_json::from_str(&json).map_err(|e| format!("AI 项目绑定损坏: {e}")))
            .transpose()
    }
    /// Resolve an old ACP session only from the backend mapping and immutable binding.
    /// Frontend mirror ids and display project names never grant authority.
    pub fn binding_for_harness(&self, harness:&str, cwd:&Path)->Result<Option<ProjectBinding>,String> {
        let row:Option<(String,String)>=self.0.query_row(
            "SELECT b.binding,s.cwd FROM sessions s JOIN project_bindings b ON b.session_id=s.session_id WHERE s.harness_id=?1",
            [harness],|row|Ok((row.get(0)?,row.get(1)?))).optional().map_err(|e|e.to_string())?;
        let Some((raw,saved_cwd))=row else { return Ok(None); };
        let binding:ProjectBinding=serde_json::from_str(&raw).map_err(|e|format!("历史绑定损坏：{e}"))?;
        if Path::new(&saved_cwd)!=cwd || binding.cwd!=cwd { return Err("历史会话属于另一个工作区".into()); }
        Ok(Some(binding))
    }
    pub fn bind(&mut self, session: &str, binding: &ProjectBinding) -> Result<(), String> {
        if session.trim().is_empty()
            || binding
                .project_id
                .as_ref()
                .is_some_and(|id| id.trim().is_empty())
        {
            return Err("会话或项目身份不能为空".into());
        }
        if let Some(existing) = self.binding(session)? {
            return if &existing == binding {
                Ok(())
            } else {
                Err("会话不可改绑，请新建会话".into())
            };
        }
        let tx = self.0.transaction().map_err(|e| e.to_string())?;
        let mapped: bool = tx
            .query_row(
                "SELECT EXISTS(SELECT 1 FROM sessions WHERE session_id=?1)",
                [session],
                |row| row.get(0),
            )
            .map_err(|e| e.to_string())?;
        if mapped {
            return Err("历史 AI 会话不可补绑定，请保留记录并新建会话".into());
        }
        tx.execute(
            "INSERT INTO project_bindings VALUES (?1, ?2)",
            params![
                session,
                serde_json::to_string(binding).map_err(|e| e.to_string())?
            ],
        )
        .map_err(|e| e.to_string())?;
        tx.commit().map_err(|e| e.to_string())
    }
    pub fn get(&self, session: &str, cwd: &Path) -> Result<Option<String>, String> {
        let row: Option<(String, String)> = self
            .0
            .query_row(
                "SELECT harness_id, cwd FROM sessions WHERE session_id = ?1",
                [session],
                |row| Ok((row.get(0)?, row.get(1)?)),
            )
            .optional()
            .map_err(|e| format!("读取 AI 会话映射失败: {e}"))?;
        match row {
            Some((id, saved_cwd)) if Path::new(&saved_cwd) == cwd => Ok(Some(id)),
            Some(_) => Err("此 AI 会话属于另一个工作区，请打开原工作区或新建会话".into()),
            None => Ok(None),
        }
    }
    pub fn insert(&self, session: &str, harness: &str, cwd: &Path) -> Result<(), String> {
        self.0
            .execute(
                "INSERT INTO sessions (session_id, harness_id, cwd) VALUES (?1, ?2, ?3)",
                params![session, harness, cwd.to_string_lossy().as_ref()],
            )
            .map_err(|e| format!("保存 AI 会话映射失败: {e}"))?;
        Ok(())
    }
    pub fn remove(&mut self, session: &str) -> Result<(), String> {
        let tx = self.0.transaction().map_err(|e| e.to_string())?;
        tx.execute("DELETE FROM sessions WHERE session_id = ?1", [session])
            .map_err(|e| e.to_string())?;
        tx.execute(
            "DELETE FROM project_bindings WHERE session_id = ?1",
            [session],
        )
        .map_err(|e| e.to_string())?;
        tx.commit().map_err(|e| e.to_string())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn mapping_survives_reopen_and_never_changes_workspace_or_identity() {
        let path = std::env::temp_dir().join(format!(
            "lamber-session-map-{}.sqlite",
            uuid::Uuid::new_v4()
        ));
        let cwd = std::env::temp_dir().join("workspace-a");
        {
            let store = SessionStore::open(&path).unwrap();
            store.insert("front-a", "acp-a", &cwd).unwrap();
            assert!(store.insert("front-a", "acp-b", &cwd).is_err());
        }
        let mut store = SessionStore::open(&path).unwrap();
        assert_eq!(
            store.get("front-a", &cwd).unwrap().as_deref(),
            Some("acp-a")
        );
        assert!(store
            .get("front-a", &cwd.join("another-workspace"))
            .is_err());
        assert_eq!(store.get("front-b", &cwd).unwrap(), None);
        store.remove("front-a").unwrap();
        assert_eq!(store.get("front-a", &cwd).unwrap(), None);
        drop(store);
        let _ = std::fs::remove_file(path);
    }
}
