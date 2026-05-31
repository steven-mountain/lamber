use super::dto::*;
use base64::Engine;
use rusqlite::{params, Connection, OptionalExtension};
use serde_json::{json, Map, Value};
use std::collections::{HashMap, HashSet};
use std::fs;
use std::path::Path;
use tauri::AppHandle;

const VALID_SOURCES: [&str; 7] = [
    "overview",
    "lifecycle",
    "cashflow",
    "benefit",
    "templates",
    "template_detail",
    "files",
];
const MAX_TEMPLATE_STRING_CHARS: usize = 2000;
const MAX_TEMPLATE_ARRAY_ITEMS: usize = 40;
const MAX_TEMPLATE_OBJECT_KEYS: usize = 120;
const MAX_TEMPLATE_DEPTH: usize = 6;

struct ProjectRow {
    id: String,
    name: String,
    customer_name: String,
    status: String,
    benefit_status: String,
    default_scheme_id: Option<String>,
    updated_at: String,
    summary_metrics: Option<String>,
    folder_path: Option<String>,
    note: Option<String>,
    progress: f64,
    deadline: Option<String>,
    linked_folder_type: Option<String>,
    linked_folder_relative_path: Option<String>,
    linked_folder_external_path: Option<String>,
}

pub fn build_ai_project_context(
    db: &Connection,
    workspace_root: &str,
    request: AiProjectContextRequest,
) -> Result<AiProjectContextBundle, String> {
    let project_id = request.project_id.trim();
    if project_id.is_empty() {
        return Err("ProjectIdRequired".to_string());
    }

    let project = load_project(db, project_id)?
        .ok_or_else(|| format!("ProjectNotFoundInCurrentWorkspace::{}", project_id))?;

    let (sources_to_load, explicit_request, mut warnings) =
        normalize_requested_sources(request.requested_sources.as_ref());
    let detailed = |source: &str| explicit_request && sources_to_load.contains(source);

    let mut source_meta = vec![AiContextSourceMeta {
        source_type: "projects".to_string(),
        source_id: Some(project.id.clone()),
        updated_at: Some(project.updated_at.clone()),
    }];

    let overview = AiProjectOverview {
        name: project.name.clone(),
        customer_name: Some(project.customer_name.clone()),
        status: Some(project.status.clone()),
        phase: None,
        deadline: project.deadline.clone(),
        description: project.note.clone(),
        progress: Some(project.progress),
        benefit_status: Some(project.benefit_status.clone()),
        folder_linked: is_folder_linked(&project),
        updated_at: Some(project.updated_at.clone()),
    };

    let lifecycle = if sources_to_load.contains("lifecycle") {
        let value =
            load_ai_lifecycle_context(db, &project.id, detailed("lifecycle"), &mut warnings)?;
        if let Some(ctx) = &value {
            source_meta.push(AiContextSourceMeta {
                source_type: "project_lifecycle_states".to_string(),
                source_id: Some(project.id.clone()),
                updated_at: ctx.updated_at.clone(),
            });
        }
        value
    } else {
        None
    };

    let cashflow = if sources_to_load.contains("cashflow") {
        let value = load_ai_cashflow_context(db, &project.id, detailed("cashflow"), &mut warnings)?;
        if let Some(ctx) = &value {
            source_meta.push(AiContextSourceMeta {
                source_type: "project_cashflow_states".to_string(),
                source_id: Some(project.id.clone()),
                updated_at: ctx.updated_at.clone(),
            });
        }
        value
    } else {
        None
    };

    let benefit = if sources_to_load.contains("benefit") {
        let value = load_ai_benefit_context(db, &project, detailed("benefit"), &mut source_meta)?;
        Some(value)
    } else {
        None
    };

    let templates = if sources_to_load.contains("templates") {
        let value = load_ai_template_summaries(
            db,
            &project.id,
            request.active_template_id.as_deref(),
            &mut source_meta,
            &mut warnings,
        )?;
        Some(value)
    } else {
        None
    };

    let template_detail = if sources_to_load.contains("template_detail") {
        if let Some(template_id) = request.active_template_id.as_deref() {
            let template_id = template_id.trim();
            if template_id.is_empty() {
                warnings.push("TemplateDetailRequestedWithoutTemplateId".to_string());
                None
            } else {
                let value = load_ai_template_detail(
                    db,
                    Path::new(workspace_root),
                    &project.id,
                    template_id,
                )?;
                source_meta.push(AiContextSourceMeta {
                    source_type: value.source.clone(),
                    source_id: Some(value.template_id.clone()),
                    updated_at: value.updated_at.clone(),
                });
                Some(value)
            }
        } else {
            warnings.push("TemplateDetailRequestedWithoutTemplateId".to_string());
            None
        }
    } else {
        None
    };

    let files = if sources_to_load.contains("files") {
        let value = load_ai_file_summary(db, &project.id, detailed("files"), &mut source_meta)?;
        Some(value)
    } else {
        None
    };

    Ok(AiProjectContextBundle {
        project_id: project.id,
        project_name: project.name,
        overview,
        lifecycle,
        cashflow,
        benefit,
        templates,
        template_detail,
        files,
        sources: source_meta,
        warnings,
    })
}

pub fn list_ai_workspace_projects(
    db: &Connection,
) -> Result<Vec<AiWorkspaceProjectIndexItem>, String> {
    let mut stmt = db
        .prepare(
            "SELECT
                p.id,
                p.name,
                p.customer_name,
                p.status,
                p.updated_at,
                EXISTS(SELECT 1 FROM project_lifecycle_states l WHERE l.project_id = p.id),
                EXISTS(SELECT 1 FROM project_cashflow_states c WHERE c.project_id = p.id),
                EXISTS(SELECT 1 FROM project_template_states t WHERE t.project_id = p.id)
                    OR EXISTS(SELECT 1 FROM project_settings s WHERE s.project_id = p.id AND s.key LIKE 'template_form_data::%'),
                EXISTS(SELECT 1 FROM benefit_schemes b WHERE b.project_id = p.id)
             FROM projects p
             ORDER BY p.updated_at DESC, p.name COLLATE NOCASE ASC",
        )
        .map_err(|e| e.to_string())?;

    let rows = stmt
        .query_map([], |row| {
            Ok(AiWorkspaceProjectIndexItem {
                project_id: row.get(0)?,
                project_name: row.get(1)?,
                customer_name: row.get(2)?,
                status: row.get(3)?,
                updated_at: row.get(4)?,
                has_lifecycle_state: row.get::<_, i64>(5)? != 0,
                has_cashflow_state: row.get::<_, i64>(6)? != 0,
                has_template_state: row.get::<_, i64>(7)? != 0,
                template_names: Vec::new(),
                has_benefit_schemes: row.get::<_, i64>(8)? != 0,
            })
        })
        .map_err(|e| e.to_string())?;

    let mut projects = Vec::new();
    for row in rows {
        let mut item = row.map_err(|e| e.to_string())?;
        item.template_names = load_project_template_names(db, &item.project_id)?;
        projects.push(item);
    }
    Ok(projects)
}

fn load_project_template_names(db: &Connection, project_id: &str) -> Result<Vec<String>, String> {
    let mut names = Vec::new();
    let mut seen = HashSet::new();

    let mut stmt = db
        .prepare(
            "SELECT template_id, template_name
             FROM project_template_states
             WHERE project_id = ?1
             ORDER BY updated_at DESC
             LIMIT 20",
        )
        .map_err(|e| e.to_string())?;
    let rows = stmt
        .query_map([project_id], |row| {
            Ok((row.get::<_, String>(0)?, row.get::<_, Option<String>>(1)?))
        })
        .map_err(|e| e.to_string())?;
    for row in rows {
        let (template_id, template_name) = row.map_err(|e| e.to_string())?;
        let label = template_name.unwrap_or(template_id);
        if seen.insert(label.clone()) {
            names.push(label);
        }
    }
    drop(stmt);

    let mut legacy_stmt = db
        .prepare(
            "SELECT key
             FROM project_settings
             WHERE project_id = ?1 AND key LIKE 'template_form_data::%'
             ORDER BY updated_at DESC
             LIMIT 20",
        )
        .map_err(|e| e.to_string())?;
    let legacy_rows = legacy_stmt
        .query_map([project_id], |row| row.get::<_, String>(0))
        .map_err(|e| e.to_string())?;
    for row in legacy_rows {
        let key = row.map_err(|e| e.to_string())?;
        let label = key
            .strip_prefix("template_form_data::")
            .unwrap_or(&key)
            .to_string();
        if seen.insert(label.clone()) {
            names.push(label);
        }
    }

    Ok(names)
}

fn normalize_requested_sources(
    requested: Option<&Vec<String>>,
) -> (HashSet<&'static str>, bool, Vec<String>) {
    let mut warnings = Vec::new();
    let valid: HashSet<&'static str> = VALID_SOURCES.iter().copied().collect();
    let explicit = requested
        .map(|sources| !sources.is_empty())
        .unwrap_or(false);

    if !explicit {
        return (
            [
                "overview",
                "lifecycle",
                "cashflow",
                "benefit",
                "templates",
                "files",
            ]
            .iter()
            .copied()
            .collect(),
            false,
            warnings,
        );
    }

    let mut selected = HashSet::new();
    selected.insert("overview");
    if let Some(items) = requested {
        for item in items {
            let normalized = item.trim().to_ascii_lowercase();
            if normalized.is_empty() {
                continue;
            }
            if let Some(source) = valid.get(normalized.as_str()) {
                selected.insert(*source);
            } else {
                warnings.push(format!("UnsupportedAiContextSource::{}", item));
            }
        }
    }
    (selected, true, warnings)
}

fn load_project(db: &Connection, project_id: &str) -> Result<Option<ProjectRow>, String> {
    db.query_row(
        "SELECT id, name, customer_name, status, benefit_status, default_scheme_id, updated_at,
            summary_metrics, folder_path, note, progress, deadline, linked_folder_type,
            linked_folder_relative_path, linked_folder_external_path
         FROM projects WHERE id = ?1",
        [project_id],
        |row| {
            Ok(ProjectRow {
                id: row.get(0)?,
                name: row.get(1)?,
                customer_name: row.get(2)?,
                status: row.get(3)?,
                benefit_status: row.get(4)?,
                default_scheme_id: row.get(5)?,
                updated_at: row.get(6)?,
                summary_metrics: row.get(7)?,
                folder_path: row.get(8)?,
                note: row.get(9)?,
                progress: row.get(10).unwrap_or(0.0),
                deadline: row.get(11)?,
                linked_folder_type: row.get(12)?,
                linked_folder_relative_path: row.get(13)?,
                linked_folder_external_path: row.get(14)?,
            })
        },
    )
    .optional()
    .map_err(|e| e.to_string())
}

fn is_folder_linked(project: &ProjectRow) -> bool {
    let linked_type = project
        .linked_folder_type
        .as_deref()
        .unwrap_or("none")
        .trim()
        .to_ascii_lowercase();
    linked_type != "none"
        || non_empty(project.folder_path.as_deref())
        || non_empty(project.linked_folder_relative_path.as_deref())
        || non_empty(project.linked_folder_external_path.as_deref())
}

fn load_ai_lifecycle_context(
    db: &Connection,
    project_id: &str,
    include_full_json: bool,
    warnings: &mut Vec<String>,
) -> Result<Option<AiLifecycleContext>, String> {
    let row: Option<(i64, String, String, String, String, String)> = db
        .query_row(
            "SELECT lifecycle_version, profile_json, parameters_json, background_json,
                input_payload_json, updated_at
             FROM project_lifecycle_states WHERE project_id = ?1",
            [project_id],
            |row| {
                Ok((
                    row.get(0)?,
                    row.get(1)?,
                    row.get(2)?,
                    row.get(3)?,
                    row.get(4)?,
                    row.get(5)?,
                ))
            },
        )
        .optional()
        .map_err(|e| e.to_string())?;

    let Some((version, profile_raw, parameters_raw, background_raw, input_payload_raw, updated_at)) =
        row
    else {
        warnings.push("LifecycleStateMissing".to_string());
        return Ok(None);
    };

    let profile = parse_json_column("project_lifecycle_states.profile_json", &profile_raw)?;
    let parameters =
        parse_json_column("project_lifecycle_states.parameters_json", &parameters_raw)?;
    let background =
        parse_json_column("project_lifecycle_states.background_json", &background_raw)?;
    let input_payload = parse_json_column(
        "project_lifecycle_states.input_payload_json",
        &input_payload_raw,
    )?;

    let summary_json = json!({
        "profileKeys": object_keys(&profile),
        "parameterKeys": object_keys(&parameters),
        "backgroundKeys": object_keys(&background),
        "inputPayloadKeys": object_keys(&input_payload),
        "backgroundPreview": first_string_value(&background, &["projectBackground", "project_background", "background"]),
        "inputProjectYears": input_payload.get("project_years").or_else(|| input_payload.get("projectYears")).cloned(),
        "inputCashflowModel": input_payload.get("cashflow_model").or_else(|| input_payload.get("cashflowModel")).cloned(),
    });

    Ok(Some(AiLifecycleContext {
        has_saved_state: true,
        lifecycle_version: Some(version),
        updated_at: Some(updated_at),
        summary_json,
        profile_json: include_full_json.then_some(profile),
        parameters_json: include_full_json.then_some(parameters),
        background_json: include_full_json.then_some(background),
        input_payload_json: include_full_json.then_some(input_payload),
    }))
}

fn load_ai_cashflow_context(
    db: &Connection,
    project_id: &str,
    include_full_json: bool,
    warnings: &mut Vec<String>,
) -> Result<Option<AiCashflowContext>, String> {
    let row: Option<(
        i64,
        Option<String>,
        String,
        String,
        String,
        String,
        String,
        String,
    )> = db
        .query_row(
            "SELECT cashflow_version, cashflow_model, payment_model_json, yearly_cashflow_json,
                sector_cashflow_json, assumptions_json, metrics_json, updated_at
             FROM project_cashflow_states WHERE project_id = ?1",
            [project_id],
            |row| {
                Ok((
                    row.get(0)?,
                    row.get(1)?,
                    row.get(2)?,
                    row.get(3)?,
                    row.get(4)?,
                    row.get(5)?,
                    row.get(6)?,
                    row.get(7)?,
                ))
            },
        )
        .optional()
        .map_err(|e| e.to_string())?;

    let Some((
        version,
        cashflow_model,
        payment_raw,
        yearly_raw,
        sector_raw,
        assumptions_raw,
        metrics_raw,
        updated_at,
    )) = row
    else {
        warnings.push("CashflowStateMissing".to_string());
        return Ok(None);
    };

    let payment_model =
        parse_json_column("project_cashflow_states.payment_model_json", &payment_raw)?;
    let yearly_cashflow =
        parse_json_column("project_cashflow_states.yearly_cashflow_json", &yearly_raw)?;
    let sector_cashflow =
        parse_json_column("project_cashflow_states.sector_cashflow_json", &sector_raw)?;
    let assumptions =
        parse_json_column("project_cashflow_states.assumptions_json", &assumptions_raw)?;
    let metrics = parse_json_column("project_cashflow_states.metrics_json", &metrics_raw)?;
    let year_count = collection_len(&yearly_cashflow);
    let has_yearly_cashflow = year_count.unwrap_or(0) > 0;

    let summary_json = json!({
        "paymentModelKeys": object_keys(&payment_model),
        "yearlyCashflowKeys": object_keys(&yearly_cashflow),
        "sectorCashflowKeys": object_keys(&sector_cashflow),
        "assumptionKeys": object_keys(&assumptions),
        "metricKeys": object_keys(&metrics),
        "metrics": remove_large_cashflow_fields(metrics.clone()),
    });

    Ok(Some(AiCashflowContext {
        has_saved_state: true,
        cashflow_version: Some(version),
        cashflow_model,
        has_yearly_cashflow,
        year_count,
        updated_at: Some(updated_at),
        summary_json,
        payment_model_json: include_full_json.then_some(payment_model),
        yearly_cashflow_json: include_full_json.then_some(yearly_cashflow),
        sector_cashflow_json: include_full_json.then_some(sector_cashflow),
        assumptions_json: include_full_json.then_some(assumptions),
        metrics_json: include_full_json.then_some(metrics),
    }))
}

fn load_ai_benefit_context(
    db: &Connection,
    project: &ProjectRow,
    include_input_params: bool,
    source_meta: &mut Vec<AiContextSourceMeta>,
) -> Result<AiBenefitContext, String> {
    let scheme_count: i64 = db
        .query_row(
            "SELECT COUNT(*) FROM benefit_schemes WHERE project_id = ?1",
            [&project.id],
            |row| row.get(0),
        )
        .map_err(|e| e.to_string())?;

    let default_scheme = if let Some(default_id) = &project.default_scheme_id {
        load_scheme_summary(db, &project.id, default_id, true)?
    } else {
        None
    };

    let latest_scheme = db
        .query_row(
            "SELECT id, name, updated_at FROM benefit_schemes
             WHERE project_id = ?1 ORDER BY updated_at DESC LIMIT 1",
            [&project.id],
            |row| {
                let id: String = row.get(0)?;
                Ok(AiBenefitSchemeSummary {
                    is_default: project.default_scheme_id.as_deref() == Some(id.as_str()),
                    id,
                    name: row.get(1)?,
                    updated_at: row.get(2)?,
                })
            },
        )
        .optional()
        .map_err(|e| e.to_string())?;

    if let Some(scheme) = &latest_scheme {
        source_meta.push(AiContextSourceMeta {
            source_type: "benefit_schemes".to_string(),
            source_id: Some(scheme.id.clone()),
            updated_at: scheme.updated_at.clone(),
        });
    }

    let latest_snapshot = db
        .query_row(
            "SELECT id, scheme_id, version, input_params, output_metrics, created_at
             FROM benefit_snapshots WHERE project_id = ?1
             ORDER BY created_at DESC, version DESC LIMIT 1",
            [&project.id],
            |row| {
                let input_raw: String = row.get(3)?;
                let output_raw: String = row.get(4)?;
                let input_params =
                    parse_json_column_sql("benefit_snapshots.input_params", &input_raw)?;
                let output_metrics =
                    parse_json_column_sql("benefit_snapshots.output_metrics", &output_raw)?;
                Ok(AiBenefitSnapshotSummary {
                    id: row.get(0)?,
                    scheme_id: row.get(1)?,
                    version: row.get(2)?,
                    created_at: row.get(5)?,
                    output_metrics_summary: Some(remove_large_cashflow_fields(output_metrics)),
                    input_params: include_input_params.then_some(input_params),
                })
            },
        )
        .optional()
        .map_err(|e| e.to_string())?;

    if let Some(snapshot) = &latest_snapshot {
        source_meta.push(AiContextSourceMeta {
            source_type: "benefit_snapshots".to_string(),
            source_id: Some(snapshot.id.clone()),
            updated_at: snapshot.created_at.clone(),
        });
    }

    let project_summary_metrics = project
        .summary_metrics
        .as_deref()
        .map(|raw| parse_json_column("projects.summary_metrics", raw))
        .transpose()?;

    Ok(AiBenefitContext {
        scheme_count: scheme_count.max(0) as usize,
        default_scheme,
        latest_scheme,
        latest_snapshot,
        project_summary_metrics,
    })
}

fn load_scheme_summary(
    db: &Connection,
    project_id: &str,
    scheme_id: &str,
    is_default: bool,
) -> Result<Option<AiBenefitSchemeSummary>, String> {
    db.query_row(
        "SELECT id, name, updated_at FROM benefit_schemes WHERE project_id = ?1 AND id = ?2",
        params![project_id, scheme_id],
        |row| {
            Ok(AiBenefitSchemeSummary {
                id: row.get(0)?,
                name: row.get(1)?,
                updated_at: row.get(2)?,
                is_default,
            })
        },
    )
    .optional()
    .map_err(|e| e.to_string())
}

fn load_ai_template_summaries(
    db: &Connection,
    project_id: &str,
    active_template_id: Option<&str>,
    source_meta: &mut Vec<AiContextSourceMeta>,
    warnings: &mut Vec<String>,
) -> Result<Vec<AiTemplateContextSummary>, String> {
    let asset_counts = load_template_asset_counts(db, project_id, source_meta)?;
    let mut summaries = Vec::new();
    let mut seen_ids = HashSet::new();

    let sql = if active_template_id.is_some() {
        "SELECT template_id, template_name, filled_data_json, field_mapping_json, output_config_json, updated_at
         FROM project_template_states
         WHERE project_id = ?1 AND template_id = ?2
         ORDER BY updated_at DESC"
    } else {
        "SELECT template_id, template_name, filled_data_json, field_mapping_json, output_config_json, updated_at
         FROM project_template_states
         WHERE project_id = ?1
         ORDER BY updated_at DESC"
    };
    let mut stmt = db.prepare(sql).map_err(|e| e.to_string())?;
    if let Some(template_id) = active_template_id {
        let rows = stmt
            .query_map(params![project_id, template_id], |row| {
                map_template_state_row(row, &asset_counts)
            })
            .map_err(|e| e.to_string())?;
        collect_template_rows(rows, &mut summaries, &mut seen_ids)?;
    } else {
        let rows = stmt
            .query_map([project_id], |row| {
                map_template_state_row(row, &asset_counts)
            })
            .map_err(|e| e.to_string())?;
        collect_template_rows(rows, &mut summaries, &mut seen_ids)?;
    }
    drop(stmt);

    for summary in &summaries {
        source_meta.push(AiContextSourceMeta {
            source_type: "project_template_states".to_string(),
            source_id: Some(summary.template_id.clone()),
            updated_at: summary.updated_at.clone(),
        });
    }

    let legacy_like = if let Some(template_id) = active_template_id {
        format!("template_form_data::{}", escape_like(template_id))
    } else {
        "template_form_data::%".to_string()
    };
    let mut legacy_stmt = db
        .prepare(
            "SELECT key, value, updated_at FROM project_settings
             WHERE project_id = ?1 AND key LIKE ?2 ESCAPE '\\'
             ORDER BY updated_at DESC",
        )
        .map_err(|e| e.to_string())?;
    let legacy_rows = legacy_stmt
        .query_map(params![project_id, legacy_like], |row| {
            let key: String = row.get(0)?;
            let value_raw: String = row.get(1)?;
            let updated_at: Option<String> = row.get(2)?;
            Ok((key, value_raw, updated_at))
        })
        .map_err(|e| e.to_string())?;

    for row in legacy_rows {
        let (key, value_raw, updated_at) = row.map_err(|e| e.to_string())?;
        let template_id = key
            .strip_prefix("template_form_data::")
            .unwrap_or(&key)
            .to_string();
        if seen_ids.contains(&template_id) {
            continue;
        }
        let value = parse_json_column("project_settings.value", &value_raw)?;
        let asset_count = asset_count_for(&asset_counts, &template_id, Some(&template_id));
        summaries.push(AiTemplateContextSummary {
            template_id: template_id.clone(),
            template_name: Some(template_id.clone()),
            has_saved_state: true,
            updated_at: updated_at.clone(),
            field_count: Some(count_value_fields(&value)),
            asset_count: Some(asset_count),
            source: "project_settings".to_string(),
        });
        source_meta.push(AiContextSourceMeta {
            source_type: "project_settings".to_string(),
            source_id: Some(key),
            updated_at,
        });
        seen_ids.insert(template_id);
    }

    if summaries.is_empty() {
        warnings.push("TemplateStateMissing".to_string());
    }

    Ok(summaries)
}

fn load_ai_template_detail(
    db: &Connection,
    workspace_root: &Path,
    project_id: &str,
    template_id: &str,
) -> Result<AiTemplateDetailContext, String> {
    let mut warnings = Vec::new();
    let state: Option<(
        Option<String>,
        String,
        String,
        String,
        Option<String>,
    )> = db
        .query_row(
            "SELECT template_name, filled_data_json, field_mapping_json, output_config_json, updated_at
             FROM project_template_states
             WHERE project_id = ?1 AND template_id = ?2",
            params![project_id, template_id],
            |row| {
                Ok((
                    row.get(0)?,
                    row.get(1)?,
                    row.get(2)?,
                    row.get(3)?,
                    row.get(4)?,
                ))
            },
        )
        .optional()
        .map_err(|e| e.to_string())?;

    let (template_name, source, has_saved_state, updated_at, fields, field_mapping, output_config) =
        if let Some((template_name, fields_raw, mapping_raw, output_raw, updated_at)) = state {
            let fields =
                parse_json_column("project_template_states.filled_data_json", &fields_raw)?;
            let mapping =
                parse_json_column("project_template_states.field_mapping_json", &mapping_raw)?;
            let output =
                parse_json_column("project_template_states.output_config_json", &output_raw)?;
            (
                template_name,
                "project_template_states".to_string(),
                true,
                updated_at,
                sanitize_template_value(fields, "filledDataJson", 0, &mut warnings),
                Some(sanitize_template_value(
                    mapping,
                    "fieldMappingJson",
                    0,
                    &mut warnings,
                )),
                Some(sanitize_template_value(
                    output,
                    "outputConfigJson",
                    0,
                    &mut warnings,
                )),
            )
        } else {
            let legacy_key = format!("template_form_data::{}", template_id);
            let legacy_state: Option<(String, Option<String>)> = db
                .query_row(
                    "SELECT value, updated_at FROM project_settings WHERE project_id = ?1 AND key = ?2",
                    params![project_id, legacy_key],
                    |row| Ok((row.get(0)?, row.get(1)?)),
                )
                .optional()
                .map_err(|e| e.to_string())?;

            if let Some((value_raw, updated_at)) = legacy_state {
                let fields = parse_json_column("project_settings.value", &value_raw)?;
                (
                    Some(template_id.to_string()),
                    "project_settings".to_string(),
                    true,
                    updated_at,
                    sanitize_template_value(fields, "legacyFilledDataJson", 0, &mut warnings),
                    None,
                    None,
                )
            } else {
                warnings.push("TemplateSavedStateMissing".to_string());
                (
                    Some(template_id.to_string()),
                    "workspace_sqlite".to_string(),
                    false,
                    None,
                    Value::Object(Default::default()),
                    None,
                    None,
                )
            }
        };

    let assets = load_ai_template_asset_references(
        db,
        workspace_root,
        project_id,
        template_id,
        template_name.as_deref(),
        &mut warnings,
    )?;

    Ok(AiTemplateDetailContext {
        project_id: project_id.to_string(),
        template_id: template_id.to_string(),
        template_name,
        source,
        has_saved_state,
        updated_at,
        fields,
        field_mapping,
        output_config,
        assets,
        warnings,
    })
}

fn load_ai_template_asset_references(
    db: &Connection,
    workspace_root: &Path,
    project_id: &str,
    template_id: &str,
    template_name: Option<&str>,
    warnings: &mut Vec<String>,
) -> Result<Vec<AiTemplateAssetReference>, String> {
    let mut stmt = db
        .prepare(
            "SELECT id, usage, original_file_name, mime_type, file_size, width, height,
                relative_path, updated_at
             FROM project_template_assets
             WHERE project_id = ?1
                AND deleted_at IS NULL
                AND (template_id = ?2 OR template_name = ?2 OR template_name = ?3)
             ORDER BY updated_at DESC",
        )
        .map_err(|e| e.to_string())?;
    let rows = stmt
        .query_map(
            params![
                project_id,
                template_id,
                template_name.unwrap_or(template_id)
            ],
            |row| {
                Ok((
                    row.get::<_, String>(0)?,
                    row.get::<_, Option<String>>(1)?,
                    row.get::<_, Option<String>>(2)?,
                    row.get::<_, Option<String>>(3)?,
                    row.get::<_, i64>(4)?,
                    row.get::<_, Option<i32>>(5)?,
                    row.get::<_, Option<i32>>(6)?,
                    row.get::<_, String>(7)?,
                    row.get::<_, Option<String>>(8)?,
                ))
            },
        )
        .map_err(|e| e.to_string())?;

    let mut assets = Vec::new();
    for row in rows {
        let (
            asset_id,
            usage,
            file_name,
            mime_type,
            file_size,
            width,
            height,
            relative_path,
            updated_at,
        ) = row.map_err(|e| e.to_string())?;
        let exists = safe_workspace_asset_exists(workspace_root, &relative_path);
        if exists == Some(false) {
            warnings.push(format!("TemplateAssetFileMissing::{}", asset_id));
        }
        assets.push(AiTemplateAssetReference {
            asset_id,
            field_key: usage,
            file_name,
            mime_type,
            file_size,
            width,
            height,
            exists,
            updated_at,
        });
    }
    Ok(assets)
}

pub fn load_ai_template_asset(
    app_handle: &AppHandle,
    db: &Connection,
    workspace_root: &str,
    project_id: &str,
    asset_id: &str,
) -> Result<AiTemplateAssetImageInput, String> {
    if project_id.is_empty() {
        return Err("ProjectIdRequired".to_string());
    }
    if asset_id.is_empty() {
        return Err("AssetIdRequired".to_string());
    }

    let asset: Option<(
        String,
        Option<String>,
        Option<String>,
        i64,
        Option<i32>,
        Option<i32>,
        String,
    )> = db
        .query_row(
            "SELECT id, original_file_name, mime_type, file_size, width, height, relative_path
             FROM project_template_assets
             WHERE project_id = ?1 AND id = ?2 AND deleted_at IS NULL AND asset_type = 'image'",
            params![project_id, asset_id],
            |row| {
                Ok((
                    row.get(0)?,
                    row.get(1)?,
                    row.get(2)?,
                    row.get(3)?,
                    row.get(4)?,
                    row.get(5)?,
                    row.get(6)?,
                ))
            },
        )
        .optional()
        .map_err(|e| e.to_string())?;

    let Some((id, original_file_name, mime_type, file_size, width, height, relative_path)) = asset
    else {
        return Err("TemplateAssetNotFoundOrProjectMismatch".to_string());
    };

    let mime_type = mime_type.ok_or_else(|| "TemplateAssetMimeTypeMissing".to_string())?;
    if !matches!(
        mime_type.as_str(),
        "image/png" | "image/jpeg" | "image/jpg" | "image/webp"
    ) {
        return Err("UnsupportedTemplateAssetMimeType".to_string());
    }
    if file_size > 20 * 1024 * 1024 {
        return Err("TemplateAssetTooLarge".to_string());
    }

    let workspace_root_path = Path::new(workspace_root);
    let relative_path = Path::new(&relative_path);
    if relative_path.is_absolute()
        || relative_path
            .components()
            .any(|component| matches!(component, std::path::Component::ParentDir))
    {
        return Err("UnsafeTemplateAssetPath".to_string());
    }

    let mut resolved = workspace_root_path.join(relative_path);
    if !crate::workspace::is_inside_workspace(workspace_root_path, &resolved) {
        return Err("TemplateAssetOutsideWorkspace".to_string());
    }
    if !resolved.exists() {
        let fallback = crate::project_files::assets::get_template_asset_path_internal(
            app_handle,
            db,
            workspace_root,
            &id,
        )?;
        let fallback_path = Path::new(&fallback);
        if !crate::workspace::is_inside_workspace(workspace_root_path, fallback_path) {
            return Err("TemplateAssetOutsideWorkspace".to_string());
        }
        resolved = fallback_path.to_path_buf();
    }

    let bytes = fs::read(&resolved).map_err(|_| "TemplateAssetFileUnavailable".to_string())?;
    if bytes.len() > 20 * 1024 * 1024 {
        return Err("TemplateAssetTooLarge".to_string());
    }
    let encoded = base64::engine::general_purpose::STANDARD.encode(bytes);
    Ok(AiTemplateAssetImageInput {
        id,
        project_id: project_id.to_string(),
        name: original_file_name.unwrap_or_else(|| asset_id.to_string()),
        mime_type: mime_type.clone(),
        size: file_size,
        width,
        height,
        data_url: format!("data:{};base64,{}", mime_type, encoded),
        source: "workspace_sqlite_template_asset".to_string(),
    })
}

fn map_template_state_row(
    row: &rusqlite::Row<'_>,
    asset_counts: &HashMap<String, usize>,
) -> rusqlite::Result<AiTemplateContextSummary> {
    let template_id: String = row.get(0)?;
    let template_name: Option<String> = row.get(1)?;
    let filled_raw: String = row.get(2)?;
    let mapping_raw: String = row.get(3)?;
    let output_raw: String = row.get(4)?;
    let updated_at: Option<String> = row.get(5)?;
    let filled = parse_json_column_sql("project_template_states.filled_data_json", &filled_raw)?;
    let mapping =
        parse_json_column_sql("project_template_states.field_mapping_json", &mapping_raw)?;
    let output = parse_json_column_sql("project_template_states.output_config_json", &output_raw)?;

    Ok(AiTemplateContextSummary {
        asset_count: Some(asset_count_for(
            asset_counts,
            &template_id,
            template_name.as_deref(),
        )),
        template_id,
        template_name,
        has_saved_state: true,
        updated_at,
        field_count: Some(
            count_value_fields(&filled)
                + count_value_fields(&mapping)
                + count_value_fields(&output),
        ),
        source: "project_template_states".to_string(),
    })
}

fn collect_template_rows<F>(
    rows: rusqlite::MappedRows<'_, F>,
    summaries: &mut Vec<AiTemplateContextSummary>,
    seen_ids: &mut HashSet<String>,
) -> Result<(), String>
where
    F: FnMut(&rusqlite::Row<'_>) -> rusqlite::Result<AiTemplateContextSummary>,
{
    for row in rows {
        let summary = row.map_err(|e| e.to_string())?;
        seen_ids.insert(summary.template_id.clone());
        summaries.push(summary);
    }
    Ok(())
}

fn load_template_asset_counts(
    db: &Connection,
    project_id: &str,
    source_meta: &mut Vec<AiContextSourceMeta>,
) -> Result<HashMap<String, usize>, String> {
    let mut stmt = db
        .prepare(
            "SELECT COALESCE(template_id, ''), template_name, COUNT(*), MAX(updated_at)
             FROM project_template_assets
             WHERE project_id = ?1 AND deleted_at IS NULL
             GROUP BY COALESCE(template_id, ''), template_name",
        )
        .map_err(|e| e.to_string())?;
    let rows = stmt
        .query_map([project_id], |row| {
            Ok((
                row.get::<_, String>(0)?,
                row.get::<_, String>(1)?,
                row.get::<_, i64>(2)?,
                row.get::<_, Option<String>>(3)?,
            ))
        })
        .map_err(|e| e.to_string())?;

    let mut counts = HashMap::new();
    for row in rows {
        let (template_id, template_name, count, updated_at) = row.map_err(|e| e.to_string())?;
        let safe_count = count.max(0) as usize;
        if !template_id.is_empty() {
            counts.insert(format!("id:{}", template_id), safe_count);
        }
        counts.insert(format!("name:{}", template_name), safe_count);
        source_meta.push(AiContextSourceMeta {
            source_type: "project_template_assets".to_string(),
            source_id: Some(if template_id.is_empty() {
                template_name
            } else {
                template_id
            }),
            updated_at,
        });
    }
    Ok(counts)
}

fn load_ai_file_summary(
    db: &Connection,
    project_id: &str,
    include_files: bool,
    source_meta: &mut Vec<AiContextSourceMeta>,
) -> Result<AiFileContextSummary, String> {
    let (total, existing, main_doc, main_budget, updated_at): (i64, i64, i64, i64, Option<String>) =
        db.query_row(
            "SELECT COUNT(*), COALESCE(SUM(CASE WHEN \"exists\" = 1 THEN 1 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN is_main_document = 1 THEN 1 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN is_main_budget_file = 1 THEN 1 ELSE 0 END), 0),
                MAX(updated_at)
             FROM project_files WHERE project_id = ?1",
            [project_id],
            |row| {
                Ok((
                    row.get(0)?,
                    row.get(1)?,
                    row.get(2)?,
                    row.get(3)?,
                    row.get(4)?,
                ))
            },
        )
        .map_err(|e| e.to_string())?;

    source_meta.push(AiContextSourceMeta {
        source_type: "project_files".to_string(),
        source_id: Some(project_id.to_string()),
        updated_at,
    });

    let file_type_counts = load_named_counts(
        db,
        "SELECT file_type, COUNT(*) FROM project_files WHERE project_id = ?1 GROUP BY file_type ORDER BY COUNT(*) DESC, file_type",
        project_id,
    )?;
    let storage_mode_counts = load_named_counts(
        db,
        "SELECT storage_mode, COUNT(*) FROM project_files WHERE project_id = ?1 GROUP BY storage_mode ORDER BY COUNT(*) DESC, storage_mode",
        project_id,
    )?;

    let files = if include_files {
        let mut stmt = db
            .prepare(
                "SELECT id, file_name, file_type, extension, size, \"exists\", storage_mode,
                    is_main_document, is_main_budget_file, file_role, modified_at, updated_at
                 FROM project_files WHERE project_id = ?1 ORDER BY updated_at DESC, file_name",
            )
            .map_err(|e| e.to_string())?;
        let rows = stmt
            .query_map([project_id], |row| {
                let exists: i64 = row.get(5)?;
                let is_main_document: i64 = row.get(7)?;
                let is_main_budget_file: i64 = row.get(8)?;
                Ok(AiProjectFileSummary {
                    id: row.get(0)?,
                    file_name: row.get(1)?,
                    file_type: row.get(2)?,
                    extension: row.get(3)?,
                    size: row.get(4)?,
                    exists: exists == 1,
                    storage_mode: row.get(6)?,
                    is_main_document: is_main_document == 1,
                    is_main_budget_file: is_main_budget_file == 1,
                    file_role: row.get(9)?,
                    modified_at: row.get(10)?,
                    updated_at: row.get(11)?,
                })
            })
            .map_err(|e| e.to_string())?;
        let mut list = Vec::new();
        for row in rows {
            list.push(row.map_err(|e| e.to_string())?);
        }
        Some(list)
    } else {
        None
    };

    Ok(AiFileContextSummary {
        total_files: total.max(0) as usize,
        existing_files: existing.max(0) as usize,
        missing_files: (total - existing).max(0) as usize,
        file_type_counts,
        storage_mode_counts,
        main_document_count: main_doc.max(0) as usize,
        main_budget_file_count: main_budget.max(0) as usize,
        files,
    })
}

fn load_named_counts(
    db: &Connection,
    sql: &str,
    project_id: &str,
) -> Result<Vec<AiNamedCount>, String> {
    let mut stmt = db.prepare(sql).map_err(|e| e.to_string())?;
    let rows = stmt
        .query_map([project_id], |row| {
            Ok(AiNamedCount {
                name: row
                    .get::<_, Option<String>>(0)?
                    .unwrap_or_else(|| "unknown".to_string()),
                count: row.get::<_, i64>(1)?.max(0) as usize,
            })
        })
        .map_err(|e| e.to_string())?;
    let mut counts = Vec::new();
    for row in rows {
        counts.push(row.map_err(|e| e.to_string())?);
    }
    Ok(counts)
}

fn parse_json_column(source: &str, raw: &str) -> Result<Value, String> {
    serde_json::from_str(raw).map_err(|e| format!("InvalidJsonColumn::{}::{}", source, e))
}

fn parse_json_column_sql(source: &str, raw: &str) -> rusqlite::Result<Value> {
    serde_json::from_str(raw).map_err(|e| {
        rusqlite::Error::FromSqlConversionFailure(
            0,
            rusqlite::types::Type::Text,
            Box::new(std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                format!("InvalidJsonColumn::{}::{}", source, e),
            )),
        )
    })
}

fn object_keys(value: &Value) -> Vec<String> {
    match value {
        Value::Object(map) => map.keys().take(40).cloned().collect(),
        _ => Vec::new(),
    }
}

fn collection_len(value: &Value) -> Option<usize> {
    match value {
        Value::Array(items) => Some(items.len()),
        Value::Object(map) => Some(map.len()),
        _ => None,
    }
}

fn count_value_fields(value: &Value) -> usize {
    match value {
        Value::Object(map) => map
            .values()
            .map(|child| match child {
                Value::Object(_) | Value::Array(_) => count_value_fields(child),
                Value::Null => 0,
                _ => 1,
            })
            .sum(),
        Value::Array(items) => items.iter().map(count_value_fields).sum(),
        Value::Null => 0,
        _ => 1,
    }
}

fn first_string_value(value: &Value, keys: &[&str]) -> Option<String> {
    let Value::Object(map) = value else {
        return None;
    };
    for key in keys {
        if let Some(Value::String(text)) = map.get(*key) {
            if !text.trim().is_empty() {
                return Some(truncate_text(text, 500));
            }
        }
    }
    None
}

fn remove_large_cashflow_fields(mut value: Value) -> Value {
    if let Value::Object(map) = &mut value {
        remove_keys_case_insensitive(map, &["cashflow", "cashflowRows", "rows", "yearlyCashflow"]);
    }
    value
}

fn remove_keys_case_insensitive(map: &mut Map<String, Value>, keys: &[&str]) {
    let targets: HashSet<String> = keys.iter().map(|key| key.to_ascii_lowercase()).collect();
    let to_remove: Vec<String> = map
        .keys()
        .filter(|key| targets.contains(&key.to_ascii_lowercase()))
        .cloned()
        .collect();
    for key in to_remove {
        map.remove(&key);
    }
}

fn asset_count_for(
    counts: &HashMap<String, usize>,
    template_id: &str,
    template_name: Option<&str>,
) -> usize {
    counts
        .get(&format!("id:{}", template_id))
        .or_else(|| template_name.and_then(|name| counts.get(&format!("name:{}", name))))
        .or_else(|| counts.get(&format!("name:{}", template_id)))
        .copied()
        .unwrap_or(0)
}

fn escape_like(value: &str) -> String {
    value
        .replace('\\', "\\\\")
        .replace('%', "\\%")
        .replace('_', "\\_")
}

fn non_empty(value: Option<&str>) -> bool {
    value.map(|text| !text.trim().is_empty()).unwrap_or(false)
}

fn truncate_text(value: &str, max_chars: usize) -> String {
    if value.chars().count() <= max_chars {
        return value.to_string();
    }
    value.chars().take(max_chars).collect::<String>()
}

fn sanitize_template_value(
    value: Value,
    key: &str,
    depth: usize,
    warnings: &mut Vec<String>,
) -> Value {
    match value {
        Value::String(text) => sanitize_template_string(text, key, warnings),
        Value::Array(items) => {
            if depth >= MAX_TEMPLATE_DEPTH {
                warnings.push(format!("TemplateContextTruncatedDepth::{}", key));
                return Value::String("[truncated nested array]".to_string());
            }
            let original_len = items.len();
            let mut sanitized = items
                .into_iter()
                .take(MAX_TEMPLATE_ARRAY_ITEMS)
                .map(|item| sanitize_template_value(item, key, depth + 1, warnings))
                .collect::<Vec<_>>();
            if original_len > MAX_TEMPLATE_ARRAY_ITEMS {
                warnings.push(format!(
                    "TemplateContextTruncatedArray::{}::{}",
                    key,
                    original_len - MAX_TEMPLATE_ARRAY_ITEMS
                ));
                sanitized.push(Value::String(format!(
                    "[truncated {} items]",
                    original_len - MAX_TEMPLATE_ARRAY_ITEMS
                )));
            }
            Value::Array(sanitized)
        }
        Value::Object(map) => {
            if depth >= MAX_TEMPLATE_DEPTH {
                warnings.push(format!("TemplateContextTruncatedDepth::{}", key));
                return Value::String("[truncated nested object]".to_string());
            }
            let total_keys = map.len();
            let mut sanitized = Map::new();
            for (child_key, child_value) in map.into_iter().take(MAX_TEMPLATE_OBJECT_KEYS) {
                if is_sensitive_template_key(&child_key) {
                    if should_preserve_asset_key(&child_key) {
                        sanitized.insert(
                            child_key.clone(),
                            sanitize_template_value(child_value, &child_key, depth + 1, warnings),
                        );
                    } else {
                        warnings.push(format!(
                            "TemplateContextOmittedSensitiveField::{}",
                            child_key
                        ));
                        sanitized.insert(
                            child_key,
                            Value::String(
                                "[omitted binary preview or sensitive field]".to_string(),
                            ),
                        );
                    }
                    continue;
                }
                sanitized.insert(
                    child_key.clone(),
                    sanitize_template_value(child_value, &child_key, depth + 1, warnings),
                );
            }
            if total_keys > MAX_TEMPLATE_OBJECT_KEYS {
                warnings.push(format!(
                    "TemplateContextTruncatedObjectKeys::{}::{}",
                    key,
                    total_keys - MAX_TEMPLATE_OBJECT_KEYS
                ));
                sanitized.insert(
                    "__truncatedKeys".to_string(),
                    Value::Number((total_keys - MAX_TEMPLATE_OBJECT_KEYS).into()),
                );
            }
            Value::Object(sanitized)
        }
        other => other,
    }
}

fn sanitize_template_string(text: String, key: &str, warnings: &mut Vec<String>) -> Value {
    if text.starts_with("data:") || text.starts_with("blob:") {
        warnings.push(format!("TemplateContextOmittedPreviewString::{}", key));
        return Value::String("[omitted binary preview or temporary URL]".to_string());
    }
    if key.to_ascii_lowercase().contains("path") && looks_like_absolute_path(&text) {
        warnings.push(format!("TemplateContextOmittedAbsolutePath::{}", key));
        return Value::String("[omitted absolute path]".to_string());
    }
    if text.chars().count() > MAX_TEMPLATE_STRING_CHARS {
        warnings.push(format!("TemplateContextTruncatedText::{}", key));
        return Value::String(format!(
            "{}... [truncated]",
            text.chars()
                .take(MAX_TEMPLATE_STRING_CHARS)
                .collect::<String>()
        ));
    }
    Value::String(text)
}

fn is_sensitive_template_key(key: &str) -> bool {
    let normalized = key.to_ascii_lowercase();
    normalized.contains("base64")
        || normalized.contains("dataurl")
        || normalized.contains("data_url")
        || normalized.contains("preview")
        || normalized == "src"
        || normalized == "data"
}

fn should_preserve_asset_key(key: &str) -> bool {
    matches!(
        key.to_ascii_lowercase().as_str(),
        "assetid" | "asset_id" | "width" | "height"
    )
}

fn looks_like_absolute_path(value: &str) -> bool {
    value.starts_with('/')
        || value.starts_with("\\\\")
        || value
            .get(1..3)
            .map(|slice| slice == ":\\" || slice == ":/")
            .unwrap_or(false)
}

fn safe_workspace_asset_exists(workspace_root: &Path, relative_path: &str) -> Option<bool> {
    let path = Path::new(relative_path);
    if path.is_absolute()
        || path
            .components()
            .any(|component| matches!(component, std::path::Component::ParentDir))
    {
        return None;
    }
    let resolved = workspace_root.join(path);
    if !crate::workspace::is_inside_workspace(workspace_root, &resolved) {
        return None;
    }
    Some(resolved.exists())
}

#[cfg(test)]
mod tests {
    use super::*;
    use rusqlite::Connection;

    fn setup_index_db() -> Connection {
        let conn = Connection::open_in_memory().unwrap();
        conn.execute_batch(
            "
            CREATE TABLE projects (
                id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                customer_name TEXT NOT NULL,
                status TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );
            CREATE TABLE project_lifecycle_states (project_id TEXT NOT NULL);
            CREATE TABLE project_cashflow_states (project_id TEXT NOT NULL);
            CREATE TABLE project_template_states (
                project_id TEXT NOT NULL,
                template_id TEXT NOT NULL,
                template_name TEXT,
                updated_at TEXT NOT NULL
            );
            CREATE TABLE project_settings (
                project_id TEXT NOT NULL,
                key TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );
            CREATE TABLE benefit_schemes (project_id TEXT NOT NULL);
            ",
        )
        .unwrap();
        conn
    }

    #[test]
    fn workspace_project_index_is_lightweight_and_state_aware() {
        let conn = setup_index_db();
        conn.execute(
            "INSERT INTO projects (id, name, customer_name, status, updated_at) VALUES (?1, ?2, ?3, ?4, ?5)",
            params!["p1", "test3", "customer", "active", "2026-06-01T00:00:00Z"],
        )
        .unwrap();
        conn.execute(
            "INSERT INTO project_lifecycle_states (project_id) VALUES ('p1')",
            [],
        )
        .unwrap();
        conn.execute("INSERT INTO project_settings (project_id, key, updated_at) VALUES ('p1', 'template_form_data::立项签批表.docx', '2026-06-01T00:00:00Z')", [])
            .unwrap();

        let items = list_ai_workspace_projects(&conn).unwrap();
        assert_eq!(items.len(), 1);
        assert_eq!(items[0].project_id, "p1");
        assert_eq!(items[0].project_name, "test3");
        assert!(items[0].has_lifecycle_state);
        assert!(!items[0].has_cashflow_state);
        assert!(items[0].has_template_state);
        assert_eq!(items[0].template_names, vec!["立项签批表.docx"]);
        assert!(!items[0].has_benefit_schemes);
    }
}
