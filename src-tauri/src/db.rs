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

    // Set schema_version = 2 if not exists
    {
        let mut stmt = conn.prepare("SELECT value FROM app_settings WHERE key = 'schema_version'")?;
        let version_exists = stmt.exists([])?;
        if !version_exists {
            let now = chrono::Utc::now().to_rfc3339();
            conn.execute(
                "INSERT INTO app_settings (key, value, updated_at) VALUES ('schema_version', '2', ?1)",
                [now],
            )?;
        }
    }

    // Run migration checks from Version 1 to 2
    {
        let mut stmt = conn.prepare("SELECT value FROM app_settings WHERE key = 'schema_version'")?;
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
                    conn.execute("ALTER TABLE project_files ADD COLUMN directory_id TEXT;", [])?;
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
            let mut stmt = conn.prepare("SELECT value FROM app_settings WHERE key = 'schema_version'")?;
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
            let mut stmt = conn.prepare("SELECT value FROM app_settings WHERE key = 'schema_version'")?;
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
            tx.execute("ALTER TABLE projects ADD COLUMN progress REAL DEFAULT 0.0;", [])?;
            tx.execute("ALTER TABLE projects ADD COLUMN deadline TEXT;", [])?;
            tx.execute("ALTER TABLE projects ADD COLUMN linked_folder_type TEXT DEFAULT 'none';", [])?;
            tx.execute("ALTER TABLE projects ADD COLUMN linked_folder_relative_path TEXT;", [])?;
            tx.execute("ALTER TABLE projects ADD COLUMN linked_folder_external_path TEXT;", [])?;

            let now = chrono::Utc::now().to_rfc3339();
            tx.execute(
                "UPDATE app_settings SET value = '4', updated_at = ?1 WHERE key = 'schema_version'",
                [now],
            )?;
            tx.commit()?;
        }
    }

    Ok(conn)
}
