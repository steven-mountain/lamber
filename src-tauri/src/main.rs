#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

mod docfill;
mod benefit;
mod config_manager;

use std::sync::Mutex;
use tauri::{AppHandle, Manager, State};
use config_manager::{ConfigManager, AppConfig};

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
        config.module_paths.insert(module_id.clone(), path_str.clone());
        
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
        .plugin(tauri_plugin_dialog::init())
        .setup(|app| {
            let manager = ConfigManager::new(app.handle());
            let config = manager.load();
            app.manage(Mutex::new(config));
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
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
