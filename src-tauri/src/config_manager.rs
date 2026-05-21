use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::fs;
use std::path::PathBuf;
use tauri::AppHandle;
use tauri::Manager;

#[derive(Serialize, Deserialize, Default, Clone)]
pub struct AppConfig {
    pub module_paths: HashMap<String, String>,
}

pub struct ConfigManager {
    config_path: PathBuf,
}

impl ConfigManager {
    pub fn new(app_handle: &AppHandle) -> Self {
        let mut config_path = app_handle
            .path()
            .app_data_dir()
            .expect("Failed to get app data dir");
        // Ensure the directory exists
        if !config_path.exists() {
            fs::create_dir_all(&config_path).unwrap();
        }
        config_path.push("config.json");
        Self { config_path }
    }

    pub fn load(&self) -> AppConfig {
        if !self.config_path.exists() {
            return AppConfig::default();
        }

        let content = fs::read_to_string(&self.config_path).unwrap_or_default();
        serde_json::from_str(&content).unwrap_or_default()
    }

    pub fn save(&self, config: &AppConfig) -> Result<(), String> {
        let content = serde_json::to_string_pretty(config).map_err(|e| e.to_string())?;
        fs::write(&self.config_path, content).map_err(|e| e.to_string())?;
        Ok(())
    }

    /// Resolves a path within a module's workspace and ensures the directory exists.
    pub fn resolve_module_path(&self, module_id: &str, sub_dir: &str) -> Result<PathBuf, String> {
        let config = self.load();
        let base_path_str = config
            .module_paths
            .get(module_id)
            .ok_or_else(|| format!("未设置模块 {} 的工作目录", module_id))?;

        let mut path = PathBuf::from(base_path_str);
        if !sub_dir.is_empty() {
            path.push(sub_dir);
        }

        if !path.exists() {
            fs::create_dir_all(&path).map_err(|e| format!("无法创建目录 {:?}: {}", path, e))?;
        }

        Ok(path)
    }

    /// Ensures the workspace structure (templates and output) for a module.
    pub fn ensure_workspace_structure(&self, module_id: &str) -> Result<(), String> {
        self.resolve_module_path(module_id, "templates")?;
        self.resolve_module_path(module_id, "output")?;
        Ok(())
    }
}
