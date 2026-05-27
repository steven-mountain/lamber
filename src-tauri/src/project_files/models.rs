use serde::{Deserialize, Serialize};

#[derive(Serialize, Deserialize, Clone, Debug)]
#[serde(rename_all = "camelCase")]
pub struct ProjectFile {
    pub id: String,
    pub project_id: String,
    pub file_name: String,
    pub file_path: String,
    pub original_path: Option<String>,
    pub managed_path: Option<String>,
    pub file_type: String, // "word" | "excel" | "pdf" | "ppt" | "image" | "other"
    pub extension: String,
    pub size: u64,
    pub exists: bool,
    pub last_scanned_at: Option<String>,
    pub modified_at: String,
    pub storage_mode: String, // "linked" | "copied"
    pub is_main_document: bool,
    pub is_main_budget_file: bool,
    pub note: Option<String>,
    pub created_at: String,
    pub updated_at: String,

    // Phase 2 path resilience fields
    #[serde(default)]
    pub root_id: Option<String>,
    #[serde(default)]
    pub directory_id: Option<String>,
    #[serde(default)]
    pub relative_path: Option<String>,
    #[serde(default)]
    pub absolute_path_snapshot: Option<String>,
    #[serde(default)]
    pub file_hash: Option<String>,
    #[serde(default)]
    pub file_role: Option<String>,
}
