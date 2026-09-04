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
            linked_folder_external_path TEXT,
            project_type TEXT NOT NULL DEFAULT 'ict'
        );",
        [],
    )?;

    conn.execute(
        "CREATE TABLE IF NOT EXISTS benefit_schemes (
            id TEXT PRIMARY KEY,
            project_id TEXT NOT NULL,
            name TEXT NOT NULL,
            stage TEXT,
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

    // Agent 工具审批审计日志。Append-only：每条记录一次"谁在什么时候批准/拒绝了
    // 哪个工具调用"，用于事后追溯 AI 代理的写操作授权。不引用 projects，因为审批
    // 与具体项目无关，且必须在没有打开项目时也能落库。
    conn.execute(
        "CREATE TABLE IF NOT EXISTS agent_approval_log (
            request_id TEXT PRIMARY KEY,
            tool_name TEXT NOT NULL,
            call_id TEXT,
            reason TEXT,
            args_json TEXT NOT NULL DEFAULT '{}',
            approved INTEGER NOT NULL,
            decided_by TEXT NOT NULL,
            decision_reason TEXT NOT NULL DEFAULT '',
            requested_at TEXT NOT NULL,
            decided_at TEXT NOT NULL
        );",
        [],
    )?;

    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_agent_approval_decided_at
            ON agent_approval_log(decided_at);",
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
            project_id TEXT NOT NULL,
            scheme_id TEXT NOT NULL DEFAULT '',
            lifecycle_version INTEGER NOT NULL DEFAULT 1,
            profile_json TEXT NOT NULL DEFAULT '{}',
            parameters_json TEXT NOT NULL DEFAULT '{}',
            background_json TEXT NOT NULL DEFAULT '{}',
            input_payload_json TEXT NOT NULL DEFAULT '{}',
            updated_at TEXT NOT NULL,
            created_at TEXT NOT NULL,
            UNIQUE(project_id, scheme_id),
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );",
        [],
    )?;

    conn.execute(
        "CREATE TABLE IF NOT EXISTS project_cashflow_states (
            id TEXT PRIMARY KEY,
            project_id TEXT NOT NULL,
            scheme_id TEXT NOT NULL DEFAULT '',
            cashflow_version INTEGER NOT NULL DEFAULT 1,
            cashflow_model TEXT,
            payment_model_json TEXT NOT NULL DEFAULT '{}',
            yearly_cashflow_json TEXT NOT NULL DEFAULT '{}',
            sector_cashflow_json TEXT NOT NULL DEFAULT '{}',
            assumptions_json TEXT NOT NULL DEFAULT '{}',
            metrics_json TEXT NOT NULL DEFAULT '{}',
            updated_at TEXT NOT NULL,
            created_at TEXT NOT NULL,
            UNIQUE(project_id, scheme_id),
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );",
        [],
    )?;

    ensure_intelligent_compute_schema(&conn)?;

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
    crate::business_dictionaries::ensure_schema(&conn)?;
    crate::project_presets::ensure_schema(&conn)?;

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

            let has_project_type: bool = {
                let mut col_stmt = conn.prepare("PRAGMA table_info(projects)")?;
                let mut rows = col_stmt.query([])?;
                let mut found = false;
                while let Some(row) = rows.next()? {
                    let name: String = row.get(1)?;
                    if name == "project_type" {
                        found = true;
                        break;
                    }
                }
                found
            };

            let initial_version = if has_project_type {
                "7"
            } else if has_folder_name {
                "6"
            } else {
                "2"
            };
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
            let project_columns: Vec<String> = {
                let mut col_stmt = conn.prepare("PRAGMA table_info(projects)")?;
                let col_iter = col_stmt.query_map([], |r| r.get::<_, String>(1))?;
                let mut cols = Vec::new();
                for col in col_iter {
                    cols.push(col?);
                }
                cols
            };
            if !project_columns.contains(&"project_type".to_string()) {
                conn.execute(
                    "ALTER TABLE projects ADD COLUMN project_type TEXT NOT NULL DEFAULT 'ict';",
                    [],
                )?;
            }
            ensure_intelligent_compute_schema(&conn)?;
            migrate_legacy_ai_compute_settings(&mut conn)?;

            let now = chrono::Utc::now().to_rfc3339();
            conn.execute(
                "UPDATE app_settings SET value = '7', updated_at = ?1 WHERE key = 'schema_version'",
                [now],
            )?;
        }
    }

    // Run migration checks from Version 7 to 8: add benefit_schemes.stage (甄选阶段标签)
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
        if version < 8 {
            let scheme_columns: Vec<String> = {
                let mut col_stmt = conn.prepare("PRAGMA table_info(benefit_schemes)")?;
                let col_iter = col_stmt.query_map([], |r| r.get::<_, String>(1))?;
                let mut cols = Vec::new();
                for col in col_iter {
                    cols.push(col?);
                }
                cols
            };
            if !scheme_columns.contains(&"stage".to_string()) {
                conn.execute("ALTER TABLE benefit_schemes ADD COLUMN stage TEXT;", [])?;
            }

            let now = chrono::Utc::now().to_rfc3339();
            conn.execute(
                "UPDATE app_settings SET value = '8', updated_at = ?1 WHERE key = 'schema_version'",
                [now],
            )?;
        }
    }

    // Run migration checks from Version 8 to 9: 测算工作副本(lifecycle/cashflow state)改为按方案存储。
    // 两张状态表原本 project_id UNIQUE（每项目一行），现改为 (project_id, scheme_id) 唯一，
    // 既有行归属到项目 default_scheme_id（无默认方案则归入 '' 桶）。因需去掉旧的 project_id
    // 唯一约束，采用 SQLite 标准的“建新表→拷贝→删旧→改名”重建方式。
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
        if version < 9 {
            let table_has_column = |conn: &Connection, table: &str, column: &str| -> Result<bool> {
                let mut col_stmt = conn.prepare(&format!("PRAGMA table_info({})", table))?;
                let col_iter = col_stmt.query_map([], |r| r.get::<_, String>(1))?;
                let mut found = false;
                for col in col_iter {
                    if col? == column {
                        found = true;
                    }
                }
                Ok(found)
            };

            if !table_has_column(&conn, "project_lifecycle_states", "scheme_id")? {
                conn.execute_batch(
                    "CREATE TABLE project_lifecycle_states_v9 (
                        id TEXT PRIMARY KEY,
                        project_id TEXT NOT NULL,
                        scheme_id TEXT NOT NULL DEFAULT '',
                        lifecycle_version INTEGER NOT NULL DEFAULT 1,
                        profile_json TEXT NOT NULL DEFAULT '{}',
                        parameters_json TEXT NOT NULL DEFAULT '{}',
                        background_json TEXT NOT NULL DEFAULT '{}',
                        input_payload_json TEXT NOT NULL DEFAULT '{}',
                        updated_at TEXT NOT NULL,
                        created_at TEXT NOT NULL,
                        UNIQUE(project_id, scheme_id),
                        FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
                    );
                    INSERT INTO project_lifecycle_states_v9 (id, project_id, scheme_id, lifecycle_version, profile_json, parameters_json, background_json, input_payload_json, updated_at, created_at)
                        SELECT ls.id, ls.project_id,
                               COALESCE((SELECT default_scheme_id FROM projects WHERE id = ls.project_id), ''),
                               ls.lifecycle_version, ls.profile_json, ls.parameters_json, ls.background_json, ls.input_payload_json, ls.updated_at, ls.created_at
                        FROM project_lifecycle_states ls;
                    DROP TABLE project_lifecycle_states;
                    ALTER TABLE project_lifecycle_states_v9 RENAME TO project_lifecycle_states;
                    CREATE INDEX IF NOT EXISTS idx_project_lifecycle_project_id ON project_lifecycle_states(project_id);",
                )?;
            }

            if !table_has_column(&conn, "project_cashflow_states", "scheme_id")? {
                conn.execute_batch(
                    "CREATE TABLE project_cashflow_states_v9 (
                        id TEXT PRIMARY KEY,
                        project_id TEXT NOT NULL,
                        scheme_id TEXT NOT NULL DEFAULT '',
                        cashflow_version INTEGER NOT NULL DEFAULT 1,
                        cashflow_model TEXT,
                        payment_model_json TEXT NOT NULL DEFAULT '{}',
                        yearly_cashflow_json TEXT NOT NULL DEFAULT '{}',
                        sector_cashflow_json TEXT NOT NULL DEFAULT '{}',
                        assumptions_json TEXT NOT NULL DEFAULT '{}',
                        metrics_json TEXT NOT NULL DEFAULT '{}',
                        updated_at TEXT NOT NULL,
                        created_at TEXT NOT NULL,
                        UNIQUE(project_id, scheme_id),
                        FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
                    );
                    INSERT INTO project_cashflow_states_v9 (id, project_id, scheme_id, cashflow_version, cashflow_model, payment_model_json, yearly_cashflow_json, sector_cashflow_json, assumptions_json, metrics_json, updated_at, created_at)
                        SELECT cs.id, cs.project_id,
                               COALESCE((SELECT default_scheme_id FROM projects WHERE id = cs.project_id), ''),
                               cs.cashflow_version, cs.cashflow_model, cs.payment_model_json, cs.yearly_cashflow_json, cs.sector_cashflow_json, cs.assumptions_json, cs.metrics_json, cs.updated_at, cs.created_at
                        FROM project_cashflow_states cs;
                    DROP TABLE project_cashflow_states;
                    ALTER TABLE project_cashflow_states_v9 RENAME TO project_cashflow_states;
                    CREATE INDEX IF NOT EXISTS idx_project_cashflow_project_id ON project_cashflow_states(project_id);",
                )?;
            }

            let now = chrono::Utc::now().to_rfc3339();
            conn.execute(
                "UPDATE app_settings SET value = '9', updated_at = ?1 WHERE key = 'schema_version'",
                [now],
            )?;
        }
    }

    // v9 -> v10: Agent 工具审批审计日志。纯新增表，无数据迁移；`CREATE TABLE IF
    // NOT EXISTS` 已在上面的建表段执行过，这里只负责推进 schema_version，使既有
    // 工作区也被标记为已具备该表。
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
        if version < 10 {
            let now = chrono::Utc::now().to_rfc3339();
            conn.execute(
                "UPDATE app_settings SET value = '10', updated_at = ?1 WHERE key = 'schema_version'",
                [now],
            )?;
        }
    }

    Ok(conn)
}

fn ensure_intelligent_compute_schema(conn: &Connection) -> Result<()> {
    conn.execute(
        "CREATE TABLE IF NOT EXISTS project_intelligent_compute_states (
            project_id TEXT PRIMARY KEY,
            state_version INTEGER NOT NULL DEFAULT 1,
            active_amount_source_id TEXT,
            project_years INTEGER NOT NULL DEFAULT 1,
            discount_rate REAL NOT NULL DEFAULT 0.055,
            sync_revision INTEGER NOT NULL DEFAULT 0,
            controlled_subjects_json TEXT NOT NULL DEFAULT '{}',
            last_result_json TEXT NOT NULL DEFAULT '{}',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );",
        [],
    )?;
    conn.execute(
        "CREATE TABLE IF NOT EXISTS intelligent_compute_amount_sources (
            id TEXT PRIMARY KEY,
            project_id TEXT NOT NULL,
            name TEXT NOT NULL,
            description TEXT,
            enabled INTEGER NOT NULL DEFAULT 1,
            source_version INTEGER NOT NULL DEFAULT 1,
            metadata_json TEXT NOT NULL DEFAULT '{}',
            parameter_groups_json TEXT NOT NULL DEFAULT '[]',
            parameters_json TEXT NOT NULL DEFAULT '[]',
            revenue_items_json TEXT NOT NULL DEFAULT '[]',
            cost_items_json TEXT NOT NULL DEFAULT '[]',
            mappings_json TEXT NOT NULL DEFAULT '[]',
            calculation_snapshot_json TEXT NOT NULL DEFAULT '{}',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );",
        [],
    )?;
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_intelligent_amount_sources_project
         ON intelligent_compute_amount_sources(project_id, updated_at);",
        [],
    )?;
    Ok(())
}

fn migrate_legacy_ai_compute_settings(conn: &mut Connection) -> Result<()> {
    let legacy_rows: Vec<(String, String, String)> = {
        let mut stmt = conn.prepare(
            "SELECT project_id, value, updated_at
             FROM project_settings
             WHERE key = 'ai_compute_quote::active'",
        )?;
        let rows = stmt.query_map([], |row| {
            Ok((
                row.get::<_, String>(0)?,
                row.get::<_, String>(1)?,
                row.get::<_, String>(2)?,
            ))
        })?;
        let mut values = Vec::new();
        for row in rows {
            values.push(row?);
        }
        values
    };

    for (project_id, raw, updated_at) in legacy_rows {
        let value: serde_json::Value =
            serde_json::from_str(&raw).unwrap_or_else(|_| serde_json::json!({}));
        let blueprint = value
            .get("blueprint")
            .cloned()
            .unwrap_or_else(|| serde_json::json!({}));
        let source_id = format!("legacy_amount_source_{}", project_id);
        let project_values: (i64, f64) = conn
            .query_row(
                "SELECT project_years, discount_rate FROM projects WHERE id = ?1",
                [&project_id],
                |row| Ok((row.get(0)?, row.get(1)?)),
            )
            .unwrap_or((1, 0.055));
        let name = blueprint
            .get("name")
            .and_then(serde_json::Value::as_str)
            .unwrap_or("历史智算金额来源");
        let description = blueprint
            .get("description")
            .and_then(serde_json::Value::as_str);
        let metadata = serde_json::json!({
            "legacyBlueprintId": blueprint.get("id"),
            "legacyScenarioId": blueprint.get("scenarioId"),
            "legacySavedAt": value.get("savedAt"),
        });
        let snapshot = serde_json::json!({
            "syncState": blueprint.get("syncState").cloned().unwrap_or_else(|| serde_json::json!({})),
        });
        let sync_revision = blueprint
            .get("syncState")
            .and_then(|sync| sync.get("revision"))
            .and_then(serde_json::Value::as_i64)
            .unwrap_or(0);
        let controlled_subjects = blueprint
            .get("syncState")
            .and_then(|sync| sync.get("subjects"))
            .cloned()
            .unwrap_or_else(|| serde_json::json!({}));

        conn.execute(
            "UPDATE projects SET project_type = 'intelligent_compute' WHERE id = ?1",
            [&project_id],
        )?;
        conn.execute(
            "INSERT OR IGNORE INTO project_intelligent_compute_states (
                project_id, state_version, active_amount_source_id, project_years, discount_rate,
                sync_revision, controlled_subjects_json, last_result_json, created_at, updated_at
             ) VALUES (?1, 1, ?2, ?3, ?4, ?5, ?6, '{}', ?7, ?7)",
            rusqlite::params![
                project_id,
                source_id,
                project_values.0,
                project_values.1,
                sync_revision,
                serde_json::to_string(&controlled_subjects).unwrap_or_else(|_| "{}".to_string()),
                updated_at,
            ],
        )?;
        conn.execute(
            "INSERT OR IGNORE INTO intelligent_compute_amount_sources (
                id, project_id, name, description, enabled, source_version, metadata_json,
                parameter_groups_json, parameters_json, revenue_items_json, cost_items_json,
                mappings_json, calculation_snapshot_json, created_at, updated_at
             ) VALUES (?1, ?2, ?3, ?4, 1, 1, ?5, ?6, ?7, ?8, ?9, ?10, ?11, ?12, ?12)",
            rusqlite::params![
                source_id,
                project_id,
                name,
                description,
                serde_json::to_string(&metadata).unwrap_or_else(|_| "{}".to_string()),
                serde_json::to_string(
                    blueprint
                        .get("parameterGroups")
                        .unwrap_or(&serde_json::Value::Array(Vec::new())),
                )
                .unwrap_or_else(|_| "[]".to_string()),
                serde_json::to_string(
                    blueprint
                        .get("parameters")
                        .unwrap_or(&serde_json::Value::Array(Vec::new())),
                )
                .unwrap_or_else(|_| "[]".to_string()),
                serde_json::to_string(
                    blueprint
                        .get("revenueItems")
                        .unwrap_or(&serde_json::Value::Array(Vec::new())),
                )
                .unwrap_or_else(|_| "[]".to_string()),
                serde_json::to_string(
                    blueprint
                        .get("costItems")
                        .unwrap_or(&serde_json::Value::Array(Vec::new())),
                )
                .unwrap_or_else(|_| "[]".to_string()),
                serde_json::to_string(
                    blueprint
                        .get("mappings")
                        .unwrap_or(&serde_json::Value::Array(Vec::new())),
                )
                .unwrap_or_else(|_| "[]".to_string()),
                serde_json::to_string(&snapshot).unwrap_or_else(|_| "{}".to_string()),
                updated_at,
            ],
        )?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use rusqlite::params;

    fn temp_db_path(name: &str) -> std::path::PathBuf {
        std::env::temp_dir().join(format!(
            "lamber-{}-{}.db",
            name,
            uuid::Uuid::new_v4().simple()
        ))
    }

    fn create_v6_database(path: &Path) {
        let conn = Connection::open(path).unwrap();
        conn.execute_batch(
            "
            PRAGMA foreign_keys = ON;
            CREATE TABLE projects (
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
            );
            CREATE TABLE app_settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );
            CREATE TABLE project_settings (
                project_id TEXT NOT NULL,
                key TEXT NOT NULL,
                value TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                PRIMARY KEY(project_id, key),
                FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
            );
            ",
        )
        .unwrap();
        conn.execute(
            "INSERT INTO app_settings (key, value, updated_at)
             VALUES ('schema_version', '6', '2026-06-01T00:00:00Z')",
            [],
        )
        .unwrap();
        conn.execute(
            "INSERT INTO projects (
                id, name, customer_name, status, benefit_status, created_at, updated_at,
                total_revenue_incl, total_cost_incl, project_years, discount_rate,
                cashflow_model, logs
             ) VALUES (
                'legacy-intelligent', '历史智算项目', '客户A', '需求导入', 'normal',
                '2026-06-01T00:00:00Z', '2026-06-01T00:00:00Z',
                100, 80, 5, 0.06, 'model_a', '[]'
             )",
            [],
        )
        .unwrap();
        let legacy = serde_json::json!({
            "version": 4,
            "savedAt": "2026-06-01T00:00:00Z",
            "blueprint": {
                "id": "legacy-blueprint",
                "scenarioId": "legacy-scenario",
                "name": "H200 历史来源",
                "parameterGroups": [{"id": "scale", "name": "规模", "builtin": true}],
                "parameters": [{"id": "device-count", "key": "device_count", "name": "设备数", "value": 8}],
                "revenueItems": [{"id": "revenue-1", "name": "收入", "fundingPlan": {"enabled": true, "yearlyAmounts": {"1": 100}}}],
                "costItems": [{"id": "cost-1", "name": "成本"}],
                "mappings": [{"lineItemId": "revenue-1", "ictSubjectCode": "rev_it_cloud"}],
                "syncState": {
                    "revision": 3,
                    "subjects": {
                        "revenue:rev_it_cloud": {
                            "side": "revenue",
                            "ictSubjectCode": "rev_it_cloud",
                            "amountInclTax": 100,
                            "taxRate": 6,
                            "yearlyAmounts": [100, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                            "sourceLineItemIds": ["revenue-1"]
                        }
                    }
                }
            }
        });
        conn.execute(
            "INSERT INTO project_settings (project_id, key, value, updated_at)
             VALUES (?1, 'ai_compute_quote::active', ?2, ?3)",
            params![
                "legacy-intelligent",
                legacy.to_string(),
                "2026-06-01T00:00:00Z"
            ],
        )
        .unwrap();
    }

    /// v9 工作区升级到 v10 后必须具备审批审计表，且既有数据不受影响。
    #[test]
    fn v9_database_gains_agent_approval_log_and_preserves_rows() {
        let path = temp_db_path("v9-to-v10");
        {
            let conn = init_db(&path).unwrap();
            conn.execute(
                "UPDATE app_settings SET value = '9', updated_at = '2026-01-01T00:00:00Z' WHERE key = 'schema_version'",
                [],
            )
            .unwrap();
            conn.execute("DROP TABLE agent_approval_log", []).unwrap();
            conn.execute(
                "INSERT INTO app_settings (key, value, updated_at) VALUES ('sentinel', 'keep-me', '2026-01-01T00:00:00Z')",
                [],
            )
            .unwrap();
        }

        let conn = init_db(&path).unwrap();
        let version: String = conn
            .query_row(
                "SELECT value FROM app_settings WHERE key = 'schema_version'",
                [],
                |row| row.get(0),
            )
            .unwrap();
        assert_eq!(version, "10");

        // 表存在且可写。
        conn.execute(
            "INSERT INTO agent_approval_log (request_id, tool_name, call_id, reason, args_json, approved, decided_by, decision_reason, requested_at, decided_at)
             VALUES ('r1', 'write_test_marker', 'c1', '需要确认', '{}', 1, 'user', '用户已确认', '2026-01-01T00:00:00Z', '2026-01-01T00:00:05Z')",
            [],
        )
        .unwrap();
        let count: i64 = conn
            .query_row("SELECT COUNT(*) FROM agent_approval_log", [], |row| row.get(0))
            .unwrap();
        assert_eq!(count, 1);

        // 既有数据未被迁移破坏。
        let sentinel: String = conn
            .query_row(
                "SELECT value FROM app_settings WHERE key = 'sentinel'",
                [],
                |row| row.get(0),
            )
            .unwrap();
        assert_eq!(sentinel, "keep-me");
    }

    #[test]
    fn fresh_database_uses_schema_v10_and_defaults_projects_to_ict() {
        let path = temp_db_path("fresh-v10");
        let conn = init_db(&path).unwrap();
        let version: String = conn
            .query_row(
                "SELECT value FROM app_settings WHERE key = 'schema_version'",
                [],
                |row| row.get(0),
            )
            .unwrap();
        assert_eq!(version, "10");
        // v8 引入 benefit_schemes.stage（甄选阶段标签）
        let has_stage: bool = {
            let mut col_stmt = conn.prepare("PRAGMA table_info(benefit_schemes)").unwrap();
            let mut rows = col_stmt.query([]).unwrap();
            let mut found = false;
            while let Some(row) = rows.next().unwrap() {
                let name: String = row.get(1).unwrap();
                if name == "stage" {
                    found = true;
                    break;
                }
            }
            found
        };
        assert!(has_stage, "benefit_schemes 应包含 stage 列");
        // v9 引入 project_lifecycle_states / project_cashflow_states 的 scheme_id 列（工作副本按方案存储）
        let table_has_scheme_id = |table: &str| -> bool {
            let mut col_stmt = conn
                .prepare(&format!("PRAGMA table_info({})", table))
                .unwrap();
            let mut rows = col_stmt.query([]).unwrap();
            while let Some(row) = rows.next().unwrap() {
                let name: String = row.get(1).unwrap();
                if name == "scheme_id" {
                    return true;
                }
            }
            false
        };
        assert!(
            table_has_scheme_id("project_lifecycle_states"),
            "project_lifecycle_states 应包含 scheme_id 列"
        );
        assert!(
            table_has_scheme_id("project_cashflow_states"),
            "project_cashflow_states 应包含 scheme_id 列"
        );
        conn.execute(
            "INSERT INTO projects (
                id, name, customer_name, status, benefit_status, created_at, updated_at,
                total_revenue_incl, total_cost_incl, project_years, discount_rate,
                cashflow_model, logs
             ) VALUES (
                'ict-default', 'ICT 默认项目', '客户', '需求导入', 'not_started',
                '2026-06-01T00:00:00Z', '2026-06-01T00:00:00Z',
                0, 0, 1, 0.055, 'model_a', '[]'
             )",
            [],
        )
        .unwrap();
        let project_type: String = conn
            .query_row(
                "SELECT project_type FROM projects WHERE id = 'ict-default'",
                [],
                |row| row.get(0),
            )
            .unwrap();
        assert_eq!(project_type, "ict");
        drop(conn);
        let _ = std::fs::remove_file(path);
    }

    #[test]
    fn v6_legacy_blueprint_migrates_to_intelligent_amount_source() {
        let path = temp_db_path("v6-to-v7");
        create_v6_database(&path);
        let conn = init_db(&path).unwrap();
        let (project_type, sync_revision, active_source): (String, i64, String) = conn
            .query_row(
                "SELECT p.project_type, s.sync_revision, s.active_amount_source_id
                 FROM projects p
                 JOIN project_intelligent_compute_states s ON s.project_id = p.id
                 WHERE p.id = 'legacy-intelligent'",
                [],
                |row| Ok((row.get(0)?, row.get(1)?, row.get(2)?)),
            )
            .unwrap();
        assert_eq!(project_type, "intelligent_compute");
        assert_eq!(sync_revision, 3);
        let (name, parameters, revenue_items, mappings): (String, String, String, String) = conn
            .query_row(
                "SELECT name, parameters_json, revenue_items_json, mappings_json
                 FROM intelligent_compute_amount_sources WHERE id = ?1",
                [&active_source],
                |row| Ok((row.get(0)?, row.get(1)?, row.get(2)?, row.get(3)?)),
            )
            .unwrap();
        assert_eq!(name, "H200 历史来源");
        assert!(parameters.contains("device-count"));
        assert!(revenue_items.contains("yearlyAmounts"));
        assert!(mappings.contains("rev_it_cloud"));
        let legacy_setting_count: i64 = conn
            .query_row(
                "SELECT COUNT(*) FROM project_settings
                 WHERE project_id = 'legacy-intelligent' AND key = 'ai_compute_quote::active'",
                [],
                |row| row.get(0),
            )
            .unwrap();
        assert_eq!(legacy_setting_count, 1);
        drop(conn);
        let _ = std::fs::remove_file(path);
    }

    #[test]
    fn v7_benefit_schemes_gains_stage_column_and_preserves_rows() {
        // 模拟生产环境中已存在、但 benefit_schemes 缺少 stage 列的 v7 数据库。
        let path = temp_db_path("v7-stage-migration");
        {
            let conn = Connection::open(&path).unwrap();
            conn.execute_batch(
                "
                PRAGMA foreign_keys = ON;
                CREATE TABLE app_settings (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );
                CREATE TABLE benefit_schemes (
                    id TEXT PRIMARY KEY,
                    project_id TEXT NOT NULL,
                    name TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );
                ",
            )
            .unwrap();
            conn.execute(
                "INSERT INTO app_settings (key, value, updated_at)
                 VALUES ('schema_version', '7', '2026-06-01T00:00:00Z')",
                [],
            )
            .unwrap();
            conn.execute(
                "INSERT INTO benefit_schemes (id, project_id, name, created_at, updated_at)
                 VALUES ('scheme-1', 'proj-1', '甄选前方案', '2026-06-01T00:00:00Z', '2026-06-01T00:00:00Z')",
                [],
            )
            .unwrap();
        }

        let conn = init_db(&path).unwrap();

        let version: String = conn
            .query_row(
                "SELECT value FROM app_settings WHERE key = 'schema_version'",
                [],
                |row| row.get(0),
            )
            .unwrap();
        assert_eq!(version, "10");

        // 迁移后可读写 stage，且既有行的 stage 默认为 NULL（未标注）。
        let (name, stage): (String, Option<String>) = conn
            .query_row(
                "SELECT name, stage FROM benefit_schemes WHERE id = 'scheme-1'",
                [],
                |row| Ok((row.get(0)?, row.get(1)?)),
            )
            .unwrap();
        assert_eq!(name, "甄选前方案");
        assert_eq!(stage, None);

        conn.execute(
            "UPDATE benefit_schemes SET stage = 'post_selection' WHERE id = 'scheme-1'",
            [],
        )
        .unwrap();
        let updated_stage: Option<String> = conn
            .query_row(
                "SELECT stage FROM benefit_schemes WHERE id = 'scheme-1'",
                [],
                |row| row.get(0),
            )
            .unwrap();
        assert_eq!(updated_stage.as_deref(), Some("post_selection"));

        drop(conn);
        let _ = std::fs::remove_file(path);
    }

    #[test]
    fn v8_lifecycle_cashflow_states_gain_scheme_id_and_allow_multiple_per_project() {
        // 模拟 v8 数据库：状态表还是 project_id UNIQUE、无 scheme_id，且每项目一行。
        let path = temp_db_path("v8-scheme-state-migration");
        {
            let conn = Connection::open(&path).unwrap();
            conn.execute_batch(
                "
                PRAGMA foreign_keys = ON;
                CREATE TABLE app_settings (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );
                CREATE TABLE projects (
                    id TEXT PRIMARY KEY,
                    name TEXT NOT NULL,
                    customer_name TEXT NOT NULL,
                    project_type TEXT NOT NULL DEFAULT 'ict',
                    status TEXT NOT NULL DEFAULT '需求导入',
                    benefit_status TEXT NOT NULL DEFAULT 'not_started',
                    default_scheme_id TEXT,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    total_revenue_incl REAL NOT NULL DEFAULT 0,
                    total_cost_incl REAL NOT NULL DEFAULT 0,
                    project_years INTEGER NOT NULL DEFAULT 1,
                    discount_rate REAL NOT NULL DEFAULT 0.055,
                    cashflow_model TEXT NOT NULL DEFAULT 'model_a',
                    folder_name TEXT
                );
                CREATE TABLE project_lifecycle_states (
                    id TEXT PRIMARY KEY,
                    project_id TEXT NOT NULL UNIQUE,
                    lifecycle_version INTEGER NOT NULL DEFAULT 1,
                    profile_json TEXT NOT NULL DEFAULT '{}',
                    parameters_json TEXT NOT NULL DEFAULT '{}',
                    background_json TEXT NOT NULL DEFAULT '{}',
                    input_payload_json TEXT NOT NULL DEFAULT '{}',
                    updated_at TEXT NOT NULL,
                    created_at TEXT NOT NULL
                );
                CREATE TABLE project_cashflow_states (
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
                    created_at TEXT NOT NULL
                );
                ",
            )
            .unwrap();
            conn.execute(
                "INSERT INTO app_settings (key, value, updated_at)
                 VALUES ('schema_version', '8', '2026-06-01T00:00:00Z')",
                [],
            )
            .unwrap();
            conn.execute(
                "INSERT INTO projects (id, name, customer_name, default_scheme_id, created_at, updated_at)
                 VALUES ('proj-1', '甄选项目', '客户A', 'scheme-pre', '2026-06-01T00:00:00Z', '2026-06-01T00:00:00Z')",
                [],
            )
            .unwrap();
            conn.execute(
                "INSERT INTO project_lifecycle_states (id, project_id, lifecycle_version, input_payload_json, updated_at, created_at)
                 VALUES ('lc-1', 'proj-1', 3, '{\"amount\":\"pre\"}', 'now', 'now')",
                [],
            )
            .unwrap();
            conn.execute(
                "INSERT INTO project_cashflow_states (id, project_id, cashflow_version, updated_at, created_at)
                 VALUES ('cf-1', 'proj-1', 3, 'now', 'now')",
                [],
            )
            .unwrap();
        }

        let conn = init_db(&path).unwrap();

        let version: String = conn
            .query_row(
                "SELECT value FROM app_settings WHERE key = 'schema_version'",
                [],
                |row| row.get(0),
            )
            .unwrap();
        assert_eq!(version, "10");

        // 既有行被归属到项目 default_scheme_id，且数据保留。
        let (scheme_id, payload): (String, String) = conn
            .query_row(
                "SELECT scheme_id, input_payload_json FROM project_lifecycle_states WHERE id = 'lc-1'",
                [],
                |row| Ok((row.get(0)?, row.get(1)?)),
            )
            .unwrap();
        assert_eq!(scheme_id, "scheme-pre");
        assert_eq!(payload, "{\"amount\":\"pre\"}");

        // 旧的 project_id 唯一约束已解除：同项目可写入第二个方案的工作副本。
        conn.execute(
            "INSERT INTO project_lifecycle_states (id, project_id, scheme_id, lifecycle_version, input_payload_json, updated_at, created_at)
             VALUES ('lc-2', 'proj-1', 'scheme-post', 1, '{\"amount\":\"post\"}', 'now', 'now')",
            [],
        )
        .expect("同项目应能写入不同方案的工作副本");

        let count: i64 = conn
            .query_row(
                "SELECT COUNT(*) FROM project_lifecycle_states WHERE project_id = 'proj-1'",
                [],
                |row| row.get(0),
            )
            .unwrap();
        assert_eq!(count, 2);

        // (project_id, scheme_id) 仍唯一：重复同一方案应冲突。
        let dup = conn.execute(
            "INSERT INTO project_lifecycle_states (id, project_id, scheme_id, lifecycle_version, input_payload_json, updated_at, created_at)
             VALUES ('lc-3', 'proj-1', 'scheme-post', 1, '{}', 'now', 'now')",
            [],
        );
        assert!(dup.is_err(), "同一 (project_id, scheme_id) 应违反唯一约束");

        drop(conn);
        let _ = std::fs::remove_file(path);
    }
}
