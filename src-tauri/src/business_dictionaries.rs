use rusqlite::{params, Connection, OptionalExtension};
use serde::{Deserialize, Serialize};
use std::sync::Arc;
use tauri::State;

fn now_iso() -> String {
    chrono::Utc::now().to_rfc3339()
}

fn generate_item_id() -> String {
    format!("dictionary_item_{}", uuid::Uuid::new_v4().simple())
}

struct SeedDictionary {
    id: &'static str,
    key: &'static str,
    name: &'static str,
    description: &'static str,
    field_keys_json: &'static str,
    sort_order: i64,
    items: &'static [(&'static str, &'static str)],
}

const SEED_DICTIONARIES: &[SeedDictionary] = &[
    SeedDictionary {
        id: "dict_procurement_method",
        key: "procurement_method",
        name: "采购方式",
        description: "立项与采购环节使用的采购方式选项。",
        field_keys_json: r#"["procurement.method"]"#,
        sort_order: 10,
        items: &[
            ("短名单甄选", "短名单甄选"),
            ("采购", "采购"),
            ("其他", "其他"),
        ],
    },
    SeedDictionary {
        id: "dict_funding_source",
        key: "funding_source",
        name: "资金来源",
        description: "ICT 项目 IT 部分资金来源。",
        field_keys_json: r#"["template.it_funding_source"]"#,
        sort_order: 20,
        items: &[
            ("分公司成本开支", "分公司成本开支"),
            ("市公司专项资源", "市公司专项资源"),
        ],
    },
    SeedDictionary {
        id: "dict_business_model",
        key: "business_model",
        name: "商务模式",
        description: "ICT 项目商务与需求导入业务模式。",
        field_keys_json: r#"["template.it_business_mode","demand.it_business_mode"]"#,
        sort_order: 30,
        items: &[
            ("服务购销", "服务购销"),
            ("服务模式", "服务模式"),
            ("集成购销", "集成购销"),
            ("投资", "投资"),
        ],
    },
    SeedDictionary {
        id: "dict_yes_no",
        key: "yes_no",
        name: "是否选项",
        description: "联合体投标、单一来源等布尔业务字段的统一选项。",
        field_keys_json: r#"["tender.is_joint","procurement.single_source"]"#,
        sort_order: 40,
        items: &[("是", "是"), ("否", "否")],
    },
];

pub(crate) fn ensure_schema(conn: &Connection) -> rusqlite::Result<()> {
    conn.execute_batch(
        "CREATE TABLE IF NOT EXISTS business_dictionaries (
            id TEXT PRIMARY KEY,
            scope TEXT NOT NULL DEFAULT 'workspace',
            dictionary_key TEXT NOT NULL,
            name TEXT NOT NULL,
            description TEXT,
            applicable_field_keys_json TEXT NOT NULL DEFAULT '[]',
            enabled INTEGER NOT NULL DEFAULT 1,
            sort_order INTEGER NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            deleted_at TEXT
        );
        CREATE UNIQUE INDEX IF NOT EXISTS idx_business_dictionaries_active_key
            ON business_dictionaries(scope, dictionary_key)
            WHERE deleted_at IS NULL;
        CREATE TABLE IF NOT EXISTS business_dictionary_items (
            id TEXT PRIMARY KEY,
            dictionary_id TEXT NOT NULL,
            value TEXT NOT NULL,
            label TEXT NOT NULL,
            description TEXT,
            enabled INTEGER NOT NULL DEFAULT 1,
            is_default INTEGER NOT NULL DEFAULT 0,
            sort_order INTEGER NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            deleted_at TEXT,
            FOREIGN KEY(dictionary_id) REFERENCES business_dictionaries(id) ON DELETE CASCADE
        );
        CREATE UNIQUE INDEX IF NOT EXISTS idx_business_dictionary_items_active_value
            ON business_dictionary_items(dictionary_id, value)
            WHERE deleted_at IS NULL;
        CREATE INDEX IF NOT EXISTS idx_business_dictionary_items_order
            ON business_dictionary_items(dictionary_id, enabled, sort_order);",
    )?;

    seed_defaults(conn)
}

fn seed_defaults(conn: &Connection) -> rusqlite::Result<()> {
    let now = now_iso();
    for dictionary in SEED_DICTIONARIES {
        conn.execute(
            "INSERT OR IGNORE INTO business_dictionaries (
                id, scope, dictionary_key, name, description,
                applicable_field_keys_json, enabled, sort_order, created_at, updated_at
             ) VALUES (?1, 'workspace', ?2, ?3, ?4, ?5, 1, ?6, ?7, ?7)",
            params![
                dictionary.id,
                dictionary.key,
                dictionary.name,
                dictionary.description,
                dictionary.field_keys_json,
                dictionary.sort_order,
                now,
            ],
        )?;

        for (index, (value, label)) in dictionary.items.iter().enumerate() {
            let item_id = format!("{}_item_{}", dictionary.id, index + 1);
            conn.execute(
                "INSERT OR IGNORE INTO business_dictionary_items (
                    id, dictionary_id, value, label, enabled, is_default,
                    sort_order, created_at, updated_at
                 ) VALUES (?1, ?2, ?3, ?4, 1, 0, ?5, ?6, ?6)",
                params![
                    item_id,
                    dictionary.id,
                    value,
                    label,
                    (index as i64 + 1) * 10,
                    now,
                ],
            )?;
        }
    }
    Ok(())
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct BusinessDictionaryItem {
    pub id: String,
    pub dictionary_id: String,
    pub value: String,
    pub label: String,
    pub description: Option<String>,
    pub enabled: bool,
    pub is_default: bool,
    pub sort_order: i64,
    pub created_at: String,
    pub updated_at: String,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct BusinessDictionary {
    pub id: String,
    pub scope: String,
    pub dictionary_key: String,
    pub name: String,
    pub description: Option<String>,
    pub applicable_field_keys: Vec<String>,
    pub enabled: bool,
    pub sort_order: i64,
    pub created_at: String,
    pub updated_at: String,
    pub items: Vec<BusinessDictionaryItem>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct BusinessDictionaryItemInput {
    pub id: Option<String>,
    pub dictionary_id: String,
    pub value: String,
    pub label: String,
    pub description: Option<String>,
    pub enabled: Option<bool>,
    pub is_default: Option<bool>,
    pub sort_order: Option<i64>,
}

fn row_to_item(row: &rusqlite::Row<'_>) -> rusqlite::Result<BusinessDictionaryItem> {
    Ok(BusinessDictionaryItem {
        id: row.get(0)?,
        dictionary_id: row.get(1)?,
        value: row.get(2)?,
        label: row.get(3)?,
        description: row.get(4)?,
        enabled: row.get::<_, i64>(5)? != 0,
        is_default: row.get::<_, i64>(6)? != 0,
        sort_order: row.get(7)?,
        created_at: row.get(8)?,
        updated_at: row.get(9)?,
    })
}

fn list_items_locked(
    conn: &Connection,
    dictionary_id: &str,
    include_disabled: bool,
) -> Result<Vec<BusinessDictionaryItem>, String> {
    let mut stmt = conn
        .prepare(
            "SELECT id, dictionary_id, value, label, description, enabled, is_default,
                sort_order, created_at, updated_at
             FROM business_dictionary_items
             WHERE dictionary_id = ?1
                AND deleted_at IS NULL
                AND (?2 = 1 OR enabled = 1)
             ORDER BY sort_order ASC, created_at ASC",
        )
        .map_err(|e| e.to_string())?;
    let rows = stmt
        .query_map(
            params![dictionary_id, if include_disabled { 1 } else { 0 }],
            row_to_item,
        )
        .map_err(|e| e.to_string())?;
    rows.collect::<rusqlite::Result<Vec<_>>>()
        .map_err(|e| e.to_string())
}

fn list_dictionaries_locked(
    conn: &Connection,
    include_disabled_items: bool,
) -> Result<Vec<BusinessDictionary>, String> {
    ensure_schema(conn).map_err(|e| e.to_string())?;
    let mut stmt = conn
        .prepare(
            "SELECT id, scope, dictionary_key, name, description,
                applicable_field_keys_json, enabled, sort_order, created_at, updated_at
             FROM business_dictionaries
             WHERE deleted_at IS NULL AND scope = 'workspace'
             ORDER BY sort_order ASC, name ASC",
        )
        .map_err(|e| e.to_string())?;
    let rows = stmt
        .query_map([], |row| {
            let field_keys_raw: String = row.get(5)?;
            Ok(BusinessDictionary {
                id: row.get(0)?,
                scope: row.get(1)?,
                dictionary_key: row.get(2)?,
                name: row.get(3)?,
                description: row.get(4)?,
                applicable_field_keys: serde_json::from_str(&field_keys_raw).unwrap_or_default(),
                enabled: row.get::<_, i64>(6)? != 0,
                sort_order: row.get(7)?,
                created_at: row.get(8)?,
                updated_at: row.get(9)?,
                items: Vec::new(),
            })
        })
        .map_err(|e| e.to_string())?;

    let mut dictionaries = Vec::new();
    for row in rows {
        let mut dictionary = row.map_err(|e| e.to_string())?;
        dictionary.items = list_items_locked(conn, &dictionary.id, include_disabled_items)?;
        dictionaries.push(dictionary);
    }
    Ok(dictionaries)
}

fn get_item_locked(conn: &Connection, id: &str) -> Result<BusinessDictionaryItem, String> {
    conn.query_row(
        "SELECT id, dictionary_id, value, label, description, enabled, is_default,
            sort_order, created_at, updated_at
         FROM business_dictionary_items
         WHERE id = ?1 AND deleted_at IS NULL",
        [id],
        row_to_item,
    )
    .optional()
    .map_err(|e| e.to_string())?
    .ok_or_else(|| format!("BusinessDictionaryItemNotFound::{}", id))
}

fn save_item_locked(
    conn: &Connection,
    input: BusinessDictionaryItemInput,
) -> Result<BusinessDictionaryItem, String> {
    ensure_schema(conn).map_err(|e| e.to_string())?;
    let value = input.value.trim();
    let label = input.label.trim();
    if value.is_empty() {
        return Err("BusinessDictionaryItemValueRequired".to_string());
    }
    if label.is_empty() {
        return Err("BusinessDictionaryItemLabelRequired".to_string());
    }

    let dictionary_exists: bool = conn
        .query_row(
            "SELECT EXISTS(
                SELECT 1 FROM business_dictionaries
                WHERE id = ?1 AND deleted_at IS NULL AND enabled = 1
             )",
            [&input.dictionary_id],
            |row| row.get(0),
        )
        .map_err(|e| e.to_string())?;
    if !dictionary_exists {
        return Err(format!(
            "BusinessDictionaryNotFound::{}",
            input.dictionary_id
        ));
    }

    let now = now_iso();
    let enabled = input.enabled.unwrap_or(true);
    let is_default = input.is_default.unwrap_or(false);
    let sort_order = input.sort_order.unwrap_or(0);
    let description = input
        .description
        .map(|value| value.trim().to_string())
        .filter(|value| !value.is_empty());

    let id = if let Some(id) = input.id {
        let changed = conn
            .execute(
                "UPDATE business_dictionary_items
                 SET dictionary_id = ?1, value = ?2, label = ?3, description = ?4,
                    enabled = ?5, is_default = ?6, sort_order = ?7, updated_at = ?8
                 WHERE id = ?9 AND deleted_at IS NULL",
                params![
                    input.dictionary_id,
                    value,
                    label,
                    description,
                    if enabled { 1 } else { 0 },
                    if is_default { 1 } else { 0 },
                    sort_order,
                    now,
                    id,
                ],
            )
            .map_err(|e| e.to_string())?;
        if changed == 0 {
            return Err(format!("BusinessDictionaryItemNotFound::{}", id));
        }
        id
    } else {
        let id = generate_item_id();
        conn.execute(
            "INSERT INTO business_dictionary_items (
                id, dictionary_id, value, label, description, enabled, is_default,
                sort_order, created_at, updated_at
             ) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?9)",
            params![
                id,
                input.dictionary_id,
                value,
                label,
                description,
                if enabled { 1 } else { 0 },
                if is_default { 1 } else { 0 },
                sort_order,
                now,
            ],
        )
        .map_err(|e| e.to_string())?;
        id
    };
    get_item_locked(conn, &id)
}

fn set_item_enabled_locked(
    conn: &Connection,
    id: &str,
    enabled: bool,
) -> Result<BusinessDictionaryItem, String> {
    ensure_schema(conn).map_err(|e| e.to_string())?;
    let changed = conn
        .execute(
            "UPDATE business_dictionary_items
             SET enabled = ?1, updated_at = ?2
             WHERE id = ?3 AND deleted_at IS NULL",
            params![if enabled { 1 } else { 0 }, now_iso(), id],
        )
        .map_err(|e| e.to_string())?;
    if changed == 0 {
        return Err(format!("BusinessDictionaryItemNotFound::{}", id));
    }
    get_item_locked(conn, id)
}

fn delete_item_locked(conn: &Connection, id: &str) -> Result<(), String> {
    ensure_schema(conn).map_err(|e| e.to_string())?;
    let now = now_iso();
    let changed = conn
        .execute(
            "UPDATE business_dictionary_items
             SET enabled = 0, deleted_at = ?1, updated_at = ?1
             WHERE id = ?2 AND deleted_at IS NULL",
            params![now, id],
        )
        .map_err(|e| e.to_string())?;
    if changed == 0 {
        return Err(format!("BusinessDictionaryItemNotFound::{}", id));
    }
    Ok(())
}

fn reorder_items_locked(
    conn: &mut Connection,
    dictionary_id: &str,
    item_ids: &[String],
) -> Result<Vec<BusinessDictionaryItem>, String> {
    ensure_schema(conn).map_err(|e| e.to_string())?;
    let tx = conn.transaction().map_err(|e| e.to_string())?;
    let now = now_iso();
    for (index, item_id) in item_ids.iter().enumerate() {
        let changed = tx
            .execute(
                "UPDATE business_dictionary_items
                 SET sort_order = ?1, updated_at = ?2
                 WHERE id = ?3 AND dictionary_id = ?4 AND deleted_at IS NULL",
                params![(index as i64 + 1) * 10, now, item_id, dictionary_id],
            )
            .map_err(|e| e.to_string())?;
        if changed == 0 {
            return Err(format!(
                "BusinessDictionaryItemNotFoundInDictionary::{}",
                item_id
            ));
        }
    }
    tx.commit().map_err(|e| e.to_string())?;
    list_items_locked(conn, dictionary_id, true)
}

#[tauri::command]
pub async fn list_business_dictionaries(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    include_disabled_items: Option<bool>,
) -> Result<Vec<BusinessDictionary>, String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    list_dictionaries_locked(&conn, include_disabled_items.unwrap_or(true))
}

#[tauri::command]
pub async fn get_business_dictionary_options(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    dictionary_key: String,
) -> Result<Vec<BusinessDictionaryItem>, String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    ensure_schema(&conn).map_err(|e| e.to_string())?;
    let dictionary_id: Option<String> = conn
        .query_row(
            "SELECT id FROM business_dictionaries
             WHERE scope = 'workspace' AND dictionary_key = ?1
                AND enabled = 1 AND deleted_at IS NULL",
            [dictionary_key.trim()],
            |row| row.get(0),
        )
        .optional()
        .map_err(|e| e.to_string())?;
    match dictionary_id {
        Some(id) => list_items_locked(&conn, &id, false),
        None => Err(format!("BusinessDictionaryNotFound::{}", dictionary_key)),
    }
}

#[tauri::command]
pub async fn save_business_dictionary_item(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    item: BusinessDictionaryItemInput,
) -> Result<BusinessDictionaryItem, String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    save_item_locked(&conn, item)
}

#[tauri::command]
pub async fn set_business_dictionary_item_enabled(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    id: String,
    enabled: bool,
) -> Result<BusinessDictionaryItem, String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    set_item_enabled_locked(&conn, &id, enabled)
}

#[tauri::command]
pub async fn delete_business_dictionary_item(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    id: String,
) -> Result<(), String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let conn = db.lock().map_err(|e| e.to_string())?;
    delete_item_locked(&conn, &id)
}

#[tauri::command]
pub async fn reorder_business_dictionary_items(
    runtime: State<'_, Arc<crate::workspace::WorkspaceRuntime>>,
    dictionary_id: String,
    item_ids: Vec<String>,
) -> Result<Vec<BusinessDictionaryItem>, String> {
    runtime.require_workspace()?;
    let db = runtime.require_db()?;
    let mut conn = db.lock().map_err(|e| e.to_string())?;
    reorder_items_locked(&mut conn, &dictionary_id, &item_ids)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn defaults_and_item_lifecycle_work() {
        let mut conn = Connection::open_in_memory().unwrap();
        conn.execute("PRAGMA foreign_keys = ON", []).unwrap();
        ensure_schema(&conn).unwrap();

        let dictionaries = list_dictionaries_locked(&conn, true).unwrap();
        assert_eq!(dictionaries.len(), 4);
        let procurement = dictionaries
            .iter()
            .find(|item| item.dictionary_key == "procurement_method")
            .unwrap();
        assert_eq!(procurement.items.len(), 3);

        let created = save_item_locked(
            &conn,
            BusinessDictionaryItemInput {
                id: None,
                dictionary_id: procurement.id.clone(),
                value: "公开招标".to_string(),
                label: "公开招标".to_string(),
                description: None,
                enabled: Some(true),
                is_default: Some(false),
                sort_order: Some(5),
            },
        )
        .unwrap();
        assert_eq!(created.value, "公开招标");

        let disabled = set_item_enabled_locked(&conn, &created.id, false).unwrap();
        assert!(!disabled.enabled);
        assert!(!list_items_locked(&conn, &procurement.id, false)
            .unwrap()
            .iter()
            .any(|item| item.id == created.id));
        delete_item_locked(&conn, &created.id).unwrap();
        assert!(get_item_locked(&conn, &created.id).is_err());

        let reordered = reorder_items_locked(
            &mut conn,
            &procurement.id,
            &procurement
                .items
                .iter()
                .map(|item| item.id.clone())
                .rev()
                .collect::<Vec<_>>(),
        )
        .unwrap();
        assert_eq!(reordered[0].id, procurement.items.last().unwrap().id);
    }
}
