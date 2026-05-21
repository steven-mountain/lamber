#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

mod benefit;
mod config_manager;
mod docfill;
mod project_files;

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
            app.manage(Mutex::new(config));

            let app_data_dir = app
                .path()
                .app_data_dir()
                .expect("Failed to get app data dir");
            if !app_data_dir.exists() {
                std::fs::create_dir_all(&app_data_dir).expect("Failed to create app data dir");
            }
            let store_path = app_data_dir.join("projects_store.json");

            let repository = Box::new(benefit::repository::JsonProjectRepository::new(
                store_path.clone(),
            ));
            let service = std::sync::Arc::new(benefit::service::ProjectService::new(repository));
            app.manage(service);

            let file_repository: std::sync::Arc<
                dyn project_files::repository::ProjectFileRepository + Send + Sync,
            > = std::sync::Arc::new(project_files::repository::JsonProjectFileRepository::new(
                store_path,
            ));
            let file_service = std::sync::Arc::new(
                project_files::service::ProjectFileService::new(file_repository),
            );
            app.manage(file_service);

            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            benefit::calculate_benefit,
            benefit::calculate_ict_benefit,
            benefit::reverse_calc_ict_target,
            benefit::reverse_calc_ict_revenue_target,
            benefit::process_excel_batch,
            benefit::generate_excel_template,
            benefit::calculate_selection_fee,
            benefit::reverse_calculate_selection_fee,
            docfill::extract_docx_variables,
            docfill::generate_docx,
            docfill::batch_generate_docx_from_excel,
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
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
