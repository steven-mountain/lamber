#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

mod docfill;

mod benefit;

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

fn main() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![
            benefit::calculate_benefit, 
            benefit::calculate_ict_benefit,
            benefit::reverse_calc_ict_target,
            benefit::reverse_calc_ict_revenue_target,
            benefit::process_excel_batch,
            benefit::generate_excel_template,
            docfill::extract_docx_variables,
            docfill::generate_docx,
            docfill::batch_generate_docx_from_excel,
            docfill::generate_lifecycle_docs,
            open_file,
            docfill::get_available_templates,
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
