use rusqlite::{Connection, Result};
use std::path::Path;

pub fn init_db(db_path: &Path) -> Result<Connection> {
    let mut conn = Connection::open(db_path)?;

    // Enable foreign keys support
    conn.execute("PRAGMA foreign_keys = ON;", [])?;

    // Create tables (Version 4 structure)
    conn.execute(
        "CREATE TABLE IF NOT EXISTS projects (
            id TEXT PRIMARY KEY,
            name TEXT NOT NULL UNIQUE,
            customer_name TEXT NOT NULL,
            status TEXT NOT NULL,
            benefit_status TEXT NOT NULL,
            default_scheme_id TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            total_revenue_incl REAL NOT NULL,
            total_cost_incl REAL NOT NULL,
            project_years INTEGER NOT NULL,
            discount_rate REAL NOT NULL,
            cashflow_model TEXT NOT NULL,
            summary_metrics TEXT,
            folder_path TEXT,
            main_document_path TEXT,
            main_budget_file_path TEXT,
            note TEXT,
            logs TEXT,
            folder_name TEXT,
            relative_path TEXT,
            progress REAL DEFAULT 0.0,
            deadline TEXT,
            linked_folder_type TEXT DEFAULT 'none',
            linked_folder_relative_path TEXT,
            linked_folder_external_path TEXT
        );",
        [],
    )?;

    conn.execute(
        "CREATE TABLE IF NOT EXISTS benefit_schemes (
            id TEXT PRIMARY KEY,
            project_id TEXT NOT NULL,
            name TEXT NOT NULL,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );",
        [],
    )?;

    conn.execute(
        "CREATE TABLE IF NOT EXISTS benefit_snapshots (
            id TEXT PRIMARY KEY,
            scheme_id TEXT NOT NULL,
            project_id TEXT NOT NULL,
            version INTEGER NOT NULL,
            input_params TEXT NOT NULL,
            output_metrics TEXT NOT NULL,
            fingerprint TEXT NOT NULL,
            created_at TEXT NOT NULL,
            FOREIGN KEY(scheme_id) REFERENCES benefit_schemes(id) ON DELETE CASCADE,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );",
        [],
    )?;

    conn.execute(
        "CREATE TABLE IF NOT EXISTS project_roots (
            id TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            root_path TEXT NOT NULL UNIQUE,
            root_alias TEXT,
            is_default INTEGER NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );",
        [],
    )?;

    conn.execute(
        "CREATE TABLE IF NOT EXISTS project_directories (
            id TEXT PRIMARY KEY,
            project_id TEXT NOT NULL,
            root_id TEXT NOT NULL,
            relative_path TEXT NOT NULL,
            dir_name TEXT NOT NULL,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE,
            FOREIGN KEY(root_id) REFERENCES project_roots(id) ON DELETE CASCADE
        );",
        [],
    )?;

    conn.execute(
        "CREATE TABLE IF NOT EXISTS project_files (
            id TEXT PRIMARY KEY,
            project_id TEXT NOT NULL,
            file_name TEXT NOT NULL,
            file_path TEXT NOT NULL,
            original_path TEXT,
            managed_path TEXT,
            file_type TEXT NOT NULL,
            extension TEXT NOT NULL,
            size INTEGER NOT NULL,
            \"exists\" INTEGER NOT NULL,
            last_scanned_at TEXT,
            modified_at TEXT NOT NULL,
            storage_mode TEXT NOT NULL,
            is_main_document INTEGER NOT NULL,
            is_main_budget_file INTEGER NOT NULL,
            note TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            root_id TEXT,
            directory_id TEXT,
            relative_path TEXT,
            absolute_path_snapshot TEXT,
            file_hash TEXT,
            file_role TEXT,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE,
            FOREIGN KEY(root_id) REFERENCES project_roots(id) ON DELETE SET NULL,
            FOREIGN KEY(directory_id) REFERENCES project_directories(id) ON DELETE SET NULL
        );",
        [],
    )?;

    conn.execute(
        "CREATE TABLE IF NOT EXISTS app_settings (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );",
        [],
    )?;

    conn.execute(
        "CREATE TABLE IF NOT EXISTS project_settings (
            project_id TEXT NOT NULL,
            key TEXT NOT NULL,
            value TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            PRIMARY KEY(project_id, key),
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );",
        [],
    )?;

    conn.execute(
        "CREATE TABLE IF NOT EXISTS project_template_assets (
            id TEXT PRIMARY KEY,
            project_id TEXT NOT NULL,
            template_name TEXT NOT NULL,
            template_id TEXT,
            asset_type TEXT NOT NULL,
            usage TEXT,
            original_file_name TEXT,
            stored_file_name TEXT NOT NULL,
            relative_path TEXT NOT NULL,
            absolute_path_snapshot TEXT NOT NULL,
            mime_type TEXT,
            file_size INTEGER NOT NULL,
            width INTEGER,
            height INTEGER,
            file_hash TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            deleted_at TEXT,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );",
        [],
    )?;

    conn.execute(
        "CREATE TABLE IF NOT EXISTS project_lifecycle_states (
            id TEXT PRIMARY KEY,
            project_id TEXT NOT NULL UNIQUE,
            lifecycle_version INTEGER NOT NULL DEFAULT 1,
            profile_json TEXT NOT NULL DEFAULT '{}',
            parameters_json TEXT NOT NULL DEFAULT '{}',
            background_json TEXT NOT NULL DEFAULT '{}',
            input_payload_json TEXT NOT NULL DEFAULT '{}',
            updated_at TEXT NOT NULL,
            created_at TEXT NOT NULL,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );",
        [],
    )?;

    conn.execute(
        "CREATE TABLE IF NOT EXISTS project_cashflow_states (
            id TEXT PRIMARY KEY,
            project_id TEXT NOT NULL UNIQUE,
            cashflow_version INTEGER NOT NULL DEFAULT 1,
            cashflow_model TEXT,
            payment_model_json TEXT NOT NULL DEFAULT '{}',
            yearly_cashflow_json TEXT NOT NULL DEFAULT '{}',
            sector_cashflow_json TEXT NOT NULL DEFAULT '{}',
            assumptions_json TEXT NOT NULL DEFAULT '{}',
            metrics_json TEXT NOT NULL DEFAULT '{}',
            updated_at TEXT NOT NULL,
            created_at TEXT NOT NULL,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );",
        [],
    )?;

    conn.execute(
        "CREATE TABLE IF NOT EXISTS project_template_states (
            id TEXT PRIMARY KEY,
            project_id TEXT NOT NULL,
            template_id TEXT NOT NULL,
            template_name TEXT,
            template_type TEXT,
            template_version INTEGER NOT NULL DEFAULT 1,
            template_path TEXT,
            template_path_type TEXT,
            filled_data_json TEXT NOT NULL DEFAULT '{}',
            field_mapping_json TEXT NOT NULL DEFAULT '{}',
            output_config_json TEXT NOT NULL DEFAULT '{}',
            updated_at TEXT NOT NULL,
            created_at TEXT NOT NULL,
            UNIQUE(project_id, template_id),
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );",
        [],
    )?;

    conn.execute(
        "CREATE TABLE IF NOT EXISTS ict_templates (
            id TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            template_type TEXT NOT NULL,
            file_path TEXT NOT NULL,
            description TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );",
        [],
    )?;

    conn.execute(
        "CREATE TABLE IF NOT EXISTS project_page_contents (
            id TEXT PRIMARY KEY,
            project_id TEXT NOT NULL,
            file_id TEXT NOT NULL,
            page_num INTEGER NOT NULL,
            content TEXT NOT NULL,
            created_at TEXT NOT NULL,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE,
            FOREIGN KEY(file_id) REFERENCES project_files(id) ON DELETE CASCADE
        );",
        [],
    )?;

    conn.execute(
        "CREATE TABLE IF NOT EXISTS ai_knowledge_items (
            id TEXT PRIMARY KEY,
            project_id TEXT,
            title TEXT NOT NULL,
            content TEXT NOT NULL,
            source_type TEXT NOT NULL,
            source_id TEXT,
            embedding BLOB,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );",
        [],
    )?;

    conn.execute(
        "CREATE TABLE IF NOT EXISTS file_summaries (
            file_id TEXT PRIMARY KEY,
            project_id TEXT NOT NULL,
            summary TEXT NOT NULL,
            keywords TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY(file_id) REFERENCES project_files(id) ON DELETE CASCADE,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );",
        [],
    )?;

    crate::common_presets::ensure_schema(&conn)?;

    conn.execute("CREATE INDEX IF NOT EXISTS idx_project_lifecycle_project_id ON project_lifecycle_states(project_id);", [])?;
    conn.execute("CREATE INDEX IF NOT EXISTS idx_project_cashflow_project_id ON project_cashflow_states(project_id);", [])?;
    conn.execute("CREATE INDEX IF NOT EXISTS idx_project_template_states_project_id ON project_template_states(project_id);", [])?;
    conn.execute("CREATE INDEX IF NOT EXISTS idx_project_template_assets_project_template ON project_template_assets(project_id, template_name);", [])?;

    // Set schema_version = 2 if not exists
    {
        let mut stmt =
            conn.prepare("SELECT value FROM app_settings WHERE key = 'schema_version'")?;
        let version_exists = stmt.exists([])?;
        if !version_exists {
            let now = chrono::Utc::now().to_rfc3339();

            // Check if projects table already has version 4 columns (fresh database check)
            let has_folder_name: bool = {
                let mut col_stmt = conn.prepare("PRAGMA table_info(projects)")?;
                let mut rows = col_stmt.query([])?;
                let mut found = false;
                while let Some(row) = rows.next()? {
                    let name: String = row.get(1)?;
                    if name == "folder_name" {
                        found = true;
                        break;
                    }
                }
                found
            };

            let initial_version = if has_folder_name { "7" } else { "2" };
            conn.execute(
                "INSERT INTO app_settings (key, value, updated_at) VALUES ('schema_version', ?1, ?2)",
                [initial_version, &now],
            )?;
        }
    }

    // Run migration checks from Version 1 to 2
    {
        let mut stmt =
            conn.prepare("SELECT value FROM app_settings WHERE key = 'schema_version'")?;
        let mut rows = stmt.query([])?;
        if let Some(row) = rows.next()? {
            let val_str: String = row.get(0)?;
            let version = val_str.parse::<i32>().unwrap_or(1);
            if version == 1 {
                // Recreate project_roots as a global configuration table (safe as it was empty)
                conn.execute("DROP TABLE IF EXISTS project_roots;", [])?;
                conn.execute(
                    "CREATE TABLE project_roots (
                        id TEXT PRIMARY KEY,
                        name TEXT NOT NULL,
                        root_path TEXT NOT NULL UNIQUE,
                        root_alias TEXT,
                        is_default INTEGER NOT NULL DEFAULT 0,
                        created_at TEXT NOT NULL,
                        updated_at TEXT NOT NULL
                    );",
                    [],
                )?;

                // Query columns of project_files to avoid altering if already present
                let columns: Vec<String> = {
                    let mut col_stmt = conn.prepare("PRAGMA table_info(project_files)")?;
                    let col_iter = col_stmt.query_map([], |r| {
                        let col_name: String = r.get(1)?;
                        Ok(col_name)
                    })?;
                    let mut cols = Vec::new();
                    for col in col_iter {
                        cols.push(col?);
                    }
                    cols
                };

                if !columns.contains(&"directory_id".to_string()) {
                    conn.execute(
                        "ALTER TABLE project_files ADD COLUMN directory_id TEXT;",
                        [],
                    )?;
                }
                if !columns.contains(&"file_hash".to_string()) {
                    conn.execute("ALTER TABLE project_files ADD COLUMN file_hash TEXT;", [])?;
                }
                if !columns.contains(&"file_role".to_string()) {
                    conn.execute("ALTER TABLE project_files ADD COLUMN file_role TEXT;", [])?;
                }

                // Recreate project_directories (safe as it was empty)
                conn.execute("DROP TABLE IF EXISTS project_directories;", [])?;
                conn.execute(
                    "CREATE TABLE project_directories (
                        id TEXT PRIMARY KEY,
                        project_id TEXT NOT NULL,
                        root_id TEXT NOT NULL,
                        relative_path TEXT NOT NULL,
                        dir_name TEXT NOT NULL,
                        created_at TEXT NOT NULL,
                        updated_at TEXT NOT NULL,
                        FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE,
                        FOREIGN KEY(root_id) REFERENCES project_roots(id) ON DELETE CASCADE
                    );",
                    [],
                )?;

                // Update version
                let now = chrono::Utc::now().to_rfc3339();
                conn.execute(
                    "UPDATE app_settings SET value = '2', updated_at = ?1 WHERE key = 'schema_version'",
                    [now],
                )?;
            }
        }
    }

    // Run migration checks from Version 2 to 3
    {
        let version = {
            let mut stmt =
                conn.prepare("SELECT value FROM app_settings WHERE key = 'schema_version'")?;
            let mut rows = stmt.query([])?;
            if let Some(row) = rows.next()? {
                let val_str: String = row.get(0)?;
                val_str.parse::<i32>().unwrap_or(1)
            } else {
                1
            }
        };
        if version < 3 {
            let tx = conn.transaction()?;
            tx.execute(
                "CREATE TABLE IF NOT EXISTS project_template_assets (
                    id TEXT PRIMARY KEY,
                    project_id TEXT NOT NULL,
                    template_name TEXT NOT NULL,
                    asset_type TEXT NOT NULL,
                    usage TEXT,
                    original_file_name TEXT,
                    stored_file_name TEXT NOT NULL,
                    relative_path TEXT NOT NULL,
                    absolute_path_snapshot TEXT NOT NULL,
                    mime_type TEXT,
                    file_size INTEGER NOT NULL,
                    width INTEGER,
                    height INTEGER,
                    file_hash TEXT,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    deleted_at TEXT,
                    FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
                );",
                [],
            )?;

            let now = chrono::Utc::now().to_rfc3339();
            tx.execute(
                "UPDATE app_settings SET value = '3', updated_at = ?1 WHERE key = 'schema_version'",
                [now],
            )?;
            tx.commit()?;
        }
    }

    // Run migration checks from Version 3 to 4
    {
        let version = {
            let mut stmt =
                conn.prepare("SELECT value FROM app_settings WHERE key = 'schema_version'")?;
            let mut rows = stmt.query([])?;
            if let Some(row) = rows.next()? {
                let val_str: String = row.get(0)?;
                val_str.parse::<i32>().unwrap_or(1)
            } else {
                1
            }
        };
        if version < 4 {
            let tx = conn.transaction()?;
            tx.execute("ALTER TABLE projects ADD COLUMN folder_name TEXT;", [])?;
            tx.execute("ALTER TABLE projects ADD COLUMN relative_path TEXT;", [])?;
            tx.execute(
                "ALTER TABLE projects ADD COLUMN progress REAL DEFAULT 0.0;",
                [],
            )?;
            tx.execute("ALTER TABLE projects ADD COLUMN deadline TEXT;", [])?;
            tx.execute(
                "ALTER TABLE projects ADD COLUMN linked_folder_type TEXT DEFAULT 'none';",
                [],
            )?;
            tx.execute(
                "ALTER TABLE projects ADD COLUMN linked_folder_relative_path TEXT;",
                [],
            )?;
            tx.execute(
                "ALTER TABLE projects ADD COLUMN linked_folder_external_path TEXT;",
                [],
            )?;

            let now = chrono::Utc::now().to_rfc3339();
            tx.execute(
                "UPDATE app_settings SET value = '4', updated_at = ?1 WHERE key = 'schema_version'",
                [now],
            )?;
            tx.commit()?;
        }
    }

    // Run migration checks from Version 4 to 5
    {
        let version = {
            let mut stmt =
                conn.prepare("SELECT value FROM app_settings WHERE key = 'schema_version'")?;
            let mut rows = stmt.query([])?;
            if let Some(row) = rows.next()? {
                let val_str: String = row.get(0)?;
                val_str.parse::<i32>().unwrap_or(1)
            } else {
                1
            }
        };
        if version < 5 {
            let tx = conn.transaction()?;
            tx.execute(
                "CREATE TABLE IF NOT EXISTS project_template_assets (
                    id TEXT PRIMARY KEY,
                    project_id TEXT NOT NULL,
                    template_name TEXT NOT NULL,
                    template_id TEXT,
                    asset_type TEXT NOT NULL,
                    usage TEXT,
                    original_file_name TEXT,
                    stored_file_name TEXT NOT NULL,
                    relative_path TEXT NOT NULL,
                    absolute_path_snapshot TEXT NOT NULL,
                    mime_type TEXT,
                    file_size INTEGER NOT NULL,
                    width INTEGER,
                    height INTEGER,
                    file_hash TEXT,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    deleted_at TEXT,
                    FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
                );",
                [],
            )?;

            let asset_columns: Vec<String> = {
                let mut col_stmt = tx.prepare("PRAGMA table_info(project_template_assets)")?;
                let col_iter = col_stmt.query_map([], |r| {
                    let col_name: String = r.get(1)?;
                    Ok(col_name)
                })?;
                let mut cols = Vec::new();
                for col in col_iter {
                    cols.push(col?);
                }
                cols
            };
            if !asset_columns.contains(&"template_id".to_string()) {
                tx.execute(
                    "ALTER TABLE project_template_assets ADD COLUMN template_id TEXT;",
                    [],
                )?;
            }

            tx.execute(
                "CREATE TABLE IF NOT EXISTS project_lifecycle_states (
                    id TEXT PRIMARY KEY,
                    project_id TEXT NOT NULL UNIQUE,
                    lifecycle_version INTEGER NOT NULL DEFAULT 1,
                    profile_json TEXT NOT NULL DEFAULT '{}',
                    parameters_json TEXT NOT NULL DEFAULT '{}',
                    background_json TEXT NOT NULL DEFAULT '{}',
                    input_payload_json TEXT NOT NULL DEFAULT '{}',
                    updated_at TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
                );",
                [],
            )?;
            tx.execute(
                "CREATE TABLE IF NOT EXISTS project_cashflow_states (
                    id TEXT PRIMARY KEY,
                    project_id TEXT NOT NULL UNIQUE,
                    cashflow_version INTEGER NOT NULL DEFAULT 1,
                    cashflow_model TEXT,
                    payment_model_json TEXT NOT NULL DEFAULT '{}',
                    yearly_cashflow_json TEXT NOT NULL DEFAULT '{}',
                    sector_cashflow_json TEXT NOT NULL DEFAULT '{}',
                    assumptions_json TEXT NOT NULL DEFAULT '{}',
                    metrics_json TEXT NOT NULL DEFAULT '{}',
                    updated_at TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
                );",
                [],
            )?;
            tx.execute(
                "CREATE TABLE IF NOT EXISTS project_template_states (
                    id TEXT PRIMARY KEY,
                    project_id TEXT NOT NULL,
                    template_id TEXT NOT NULL,
                    template_name TEXT,
                    template_type TEXT,
                    template_version INTEGER NOT NULL DEFAULT 1,
                    template_path TEXT,
                    template_path_type TEXT,
                    filled_data_json TEXT NOT NULL DEFAULT '{}',
                    field_mapping_json TEXT NOT NULL DEFAULT '{}',
                    output_config_json TEXT NOT NULL DEFAULT '{}',
                    updated_at TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    UNIQUE(project_id, template_id),
                    FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
                );",
                [],
            )?;

            tx.execute("CREATE INDEX IF NOT EXISTS idx_project_lifecycle_project_id ON project_lifecycle_states(project_id);", [])?;
            tx.execute("CREATE INDEX IF NOT EXISTS idx_project_cashflow_project_id ON project_cashflow_states(project_id);", [])?;
            tx.execute("CREATE INDEX IF NOT EXISTS idx_project_template_states_project_id ON project_template_states(project_id);", [])?;
            tx.execute("CREATE INDEX IF NOT EXISTS idx_project_template_assets_project_template ON project_template_assets(project_id, template_name);", [])?;

            let now = chrono::Utc::now().to_rfc3339();
            tx.execute(
                "UPDATE app_settings SET value = '5', updated_at = ?1 WHERE key = 'schema_version'",
                [now],
            )?;
            tx.commit()?;
        }
    }

    // Run migration checks from Version 5 to 6
    {
        let version = {
            let mut stmt =
                conn.prepare("SELECT value FROM app_settings WHERE key = 'schema_version'")?;
            let mut rows = stmt.query([])?;
            if let Some(row) = rows.next()? {
                let val_str: String = row.get(0)?;
                val_str.parse::<i32>().unwrap_or(1)
            } else {
                1
            }
        };
        if version < 6 {
            let tx = conn.transaction()?;
            crate::common_presets::ensure_schema(&tx)?;
            let now = chrono::Utc::now().to_rfc3339();
            tx.execute(
                "UPDATE app_settings SET value = '6', updated_at = ?1 WHERE key = 'schema_version'",
                [now],
            )?;
            tx.commit()?;
        }
    }

    // Run migration checks from Version 6 to 7
    {
        let version = {
            let mut stmt =
                conn.prepare("SELECT value FROM app_settings WHERE key = 'schema_version'")?;
            let mut rows = stmt.query([])?;
            if let Some(row) = rows.next()? {
                let val_str: String = row.get(0)?;
                val_str.parse::<i32>().unwrap_or(1)
            } else {
                1
            }
        };
        if version < 7 {
            let tx = conn.transaction()?;
            crate::common_presets::ensure_schema(&tx)?;
            let now = chrono::Utc::now().to_rfc3339();
            tx.execute(
                "UPDATE app_settings SET value = '7', updated_at = ?1 WHERE key = 'schema_version'",
                [now],
            )?;
            tx.commit()?;
        }
    }

    Ok(conn)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn schema_v6_migrates_preset_field_settings_to_v7() {
        let path = std::env::temp_dir().join(format!(
            "lamber-schema-v7-{}.sqlite",
            uuid::Uuid::new_v4().simple()
        ));

        {
            let conn = init_db(&path).unwrap();
            conn.execute(
                "UPDATE app_settings SET value = '6' WHERE key = 'schema_version'",
                [],
            )
            .unwrap();
            conn.execute("DROP TABLE preset_field_settings", [])
                .unwrap();
        }

        {
            let conn = init_db(&path).unwrap();
            let version: String = conn
                .query_row(
                    "SELECT value FROM app_settings WHERE key = 'schema_version'",
                    [],
                    |row| row.get(0),
                )
                .unwrap();
            assert_eq!(version, "7");
            let table_exists: i64 = conn
                .query_row(
                    "SELECT COUNT(*) FROM sqlite_master
                     WHERE type = 'table' AND name = 'preset_field_settings'",
                    [],
                    |row| row.get(0),
                )
                .unwrap();
            assert_eq!(table_exists, 1);
        }

        let _ = std::fs::remove_file(path);
    }
}
