#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

mod ai_context;
mod benefit;
mod common_presets;
mod config_manager;
mod db;
mod docfill;
mod migration;
mod project_files;
mod project_state;
mod workspace;
mod workspace_maintenance;

use config_manager::{AppConfig, ConfigManager};
use std::sync::Mutex;
use tauri::{AppHandle, Manager, State};

#[tauri::command]
fn open_file(path: String) -> Result<(), String> {
    #[cfg(target_os = "windows")]
    {
        std::process::Command::new("cmd")
            .args(["/C", "start", "", &path])
            .spawn()
            .map_err(|e| e.to_string())?;
    }
    #[cfg(target_os = "macos")]
    {
        std::process::Command::new("open")
            .arg(&path)
            .spawn()
            .map_err(|e| e.to_string())?;
    }
    Ok(())
}

#[tauri::command]
fn get_module_path(state: State<'_, Mutex<AppConfig>>, module_id: String) -> Option<String> {
    let config = state.lock().unwrap();
    config.module_paths.get(&module_id).cloned()
}

#[tauri::command]
async fn set_module_path(
    app: AppHandle,
    state: State<'_, Mutex<AppConfig>>,
    module_id: String,
) -> Result<String, String> {
    use tauri_plugin_dialog::DialogExt;

    let folder = app.dialog().file().blocking_pick_folder();

    if let Some(folder_path) = folder {
        let path_str = folder_path.to_string();
        let mut config = state.lock().unwrap();
        config
            .module_paths
            .insert(module_id.clone(), path_str.clone());

        let manager = ConfigManager::new(&app);
        manager.save(&config)?;

        // Now that it's saved, we can ensure the structure
        manager.ensure_workspace_structure(&module_id)?;

        Ok(path_str)
    } else {
        Err("用户取消了选择".to_string())
    }
}

fn main() {
    tauri::Builder::default()
        .plugin(tauri_plugin_http::init())
        .plugin(tauri_plugin_dialog::init())
        .setup(|app| {
            let manager = ConfigManager::new(app.handle());
            let config = manager.load();
            app.manage(Mutex::new(config.clone()));

            let workspace_runtime = std::sync::Arc::new(workspace::WorkspaceRuntime::new());
            workspace::try_restore_last_workspace(app.handle(), &workspace_runtime, &config);
            app.manage(workspace_runtime);

            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            benefit::calculate_ict_benefit,
            ai_context::commands::build_ai_project_context,
            ai_context::commands::list_ai_workspace_projects,
            ai_context::commands::load_ai_template_asset,
            common_presets::list_common_presets,
            common_presets::save_common_preset,
            common_presets::set_common_preset_enabled,
            common_presets::delete_common_preset,
            common_presets::mark_common_preset_used,
            benefit::reverse_calc_ict_target,
            benefit::reverse_calc_ict_revenue_target,
            benefit::calculate_selection_fee,
            benefit::reverse_calculate_selection_fee,
            docfill::generate_lifecycle_docs,
            open_file,
            docfill::get_available_templates,
            get_module_path,
            set_module_path,
            benefit::get_projects,
            benefit::get_project,
            benefit::create_project,
            benefit::update_project,
            benefit::delete_project,
            benefit::delete_benefit_scheme,
            benefit::get_schemes,
            benefit::get_snapshots,
            benefit::save_benefit_scheme,
            benefit::parse_benefit_excel,
            benefit::create_project_in_workspace,
            benefit::list_workspace_projects,
            benefit::inspect_workspace_projects,
            project_files::commands::get_project_files,
            project_files::commands::bind_project_folder,
            project_files::commands::create_project_folder,
            project_files::commands::unbind_project_folder,
            project_files::commands::scan_project_folder,
            project_files::commands::add_project_file,
            project_files::commands::remove_project_file_record,
            project_files::commands::delete_managed_project_file,
            project_files::commands::mark_main_document,
            project_files::commands::mark_main_budget_file,
            project_files::commands::open_project_folder,
            project_files::commands::open_project_file,
            project_files::commands::reveal_project_file,
            project_files::commands::select_local_folder,
            project_files::commands::select_local_file,
            migration::check_db_migration,
            migration::run_db_migration,
            migration::skip_db_migration,
            project_files::roots::get_project_roots,
            project_files::roots::create_project_root,
            project_files::roots::update_project_root,
            project_files::roots::delete_project_root,
            project_files::roots::set_default_project_root,
            project_files::health::run_file_health_check,
            project_files::relocation::get_relocation_preview,
            project_files::relocation::execute_bulk_relocation,
            project_files::import_scanner::scan_import_candidates,
            project_files::import_scanner::execute_bulk_import,
            project_files::commands::save_template_asset,
            project_files::commands::get_template_asset_path,
            project_files::commands::delete_template_asset,
            project_files::commands::cleanup_orphan_template_assets,
            project_files::commands::get_project_setting,
            project_files::commands::save_project_setting,
            project_state::save_project_detail,
            project_state::get_project_detail,
            project_state::save_lifecycle_state,
            project_state::get_lifecycle_state,
            project_state::save_cashflow_state,
            project_state::get_cashflow_state,
            project_state::save_benefit_analysis,
            project_state::get_benefit_schemes,
            project_state::save_template_state,
            project_state::get_template_state,
            project_state::list_template_states,
            project_state::list_template_assets,
            project_state::get_project_full_state,
            workspace::get_workspace_state,
            workspace::inspect_workspace_path,
            workspace::select_workspace_folder,
            workspace::create_workspace,
            workspace::open_workspace,
            workspace::forget_workspace,
            workspace::close_current_workspace,
            workspace::clear_workspace,
            workspace::initialize_workspace_from_existing_directory,
            workspace::scan_and_import_all_workspace_calculations,
            workspace_maintenance::export_workspace,
            workspace_maintenance::import_workspace,
            workspace_maintenance::validate_workspace_archive,
            workspace_maintenance::reveal_in_file_manager,
            workspace_maintenance::create_workspace_backup,
            workspace_maintenance::list_workspace_backups,
            workspace_maintenance::restore_workspace_backup,
            workspace_maintenance::delete_workspace_backup,
            workspace_maintenance::run_workspace_health_check,
            workspace_maintenance::repair_workspace_issue,
            workspace_maintenance::repair_workspace_issues,
            workspace_maintenance::inspect_workspace_paths,
            workspace_maintenance::list_external_paths,
            workspace_maintenance::convert_internal_absolute_paths_to_relative,
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
