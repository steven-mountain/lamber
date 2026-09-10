//! Locate the packaged dsh runtime and initialize its writable per-user home.

use crate::config_manager::AiAgentSettings;
use std::fs;
use std::path::{Path, PathBuf};
use tauri::{AppHandle, Manager};

const RESOURCE_DIRECTORY: &str = "agent-runtime";
const DSH_ENTRY: &str = "node_modules/@deepseek-ai/dsh/lib/bin.js";
const PATCH_FILE: &str = "patch.yml";
const HOME_TEMPLATE: &str = "dsh-home-template";
const LAMBER_PLUGIN: &str = "dsh-tool-lamber";
const RUNTIME_HOME: &str = "dsh-runtime";
const SETTINGS_PATCH: &str = "lamber-settings.patch.yml";

#[derive(Debug, Clone)]
pub struct AgentDistribution {
    pub node_bin: PathBuf,
    pub dsh_entry: PathBuf,
    pub base_patch: PathBuf,
    pub home_template: PathBuf,
    pub lamber_plugin: PathBuf,
}

impl AgentDistribution {
    fn packaged(root: PathBuf) -> Result<Self, String> {
        let node_bin = root.join(if cfg!(windows) { "node.exe" } else { "node" });
        Self::validated(root, node_bin)
    }

    pub(super) fn development(repo_root: &Path) -> Result<Self, String> {
        let root = repo_root.join("agent-bridge");
        let node_bin = std::env::var_os("LAMBER_NODE_BIN")
            .map(PathBuf::from)
            .unwrap_or_else(|| PathBuf::from("node"));
        Self::validated(root, node_bin)
    }

    fn validated(root: PathBuf, node_bin: PathBuf) -> Result<Self, String> {
        let distribution = Self {
            dsh_entry: root.join(DSH_ENTRY),
            base_patch: root.join(PATCH_FILE),
            home_template: root.join(HOME_TEMPLATE),
            lamber_plugin: root.join(LAMBER_PLUGIN),
            node_bin,
        };
        let mut missing = Vec::new();
        for path in [
            &distribution.dsh_entry,
            &distribution.base_patch,
            &distribution.home_template.join("profiles/acp/package.json"),
            &distribution.lamber_plugin.join("lib/index.js"),
            &root.join("webui/lamber-brand/lib/client.js"),
            &root.join("webui/lamber-host/host-policy.js"),
            &root.join("webui/lamber-host/gateway.js"),
            &root.join("webui/lamber-host/business-presentation.generated.js"),
        ] {
            if !path.is_file() && !path.is_dir() {
                missing.push(path.display().to_string());
            }
        }
        if distribution.node_bin.components().count() > 1 && !distribution.node_bin.is_file() {
            missing.push(distribution.node_bin.display().to_string());
        }
        if missing.is_empty() {
            Ok(distribution)
        } else {
            Err(format!("缺少运行文件：{}", missing.join("、")))
        }
    }
}

/// Locate dsh in product-first order: packaged resources, explicit developer
/// root, then the repository layout surrounding a development executable.
pub fn locate(app: &AppHandle) -> Result<AgentDistribution, String> {
    let packaged_root = app
        .path()
        .resource_dir()
        .map(|resources| resources.join(RESOURCE_DIRECTORY))
        .map_err(|error| format!("安装资源目录不可用：{error}"));
    let explicit_root = std::env::var("LAMBER_REPO_ROOT")
        .ok()
        .filter(|value| !value.is_empty())
        .map(PathBuf::from);
    let executable =
        std::env::current_exe().map_err(|error| format!("无法定位当前可执行文件：{error}"));
    locate_from_candidates(packaged_root, explicit_root, executable)
}

fn locate_from_candidates(
    packaged_root: Result<PathBuf, String>,
    explicit_root: Option<PathBuf>,
    executable: Result<PathBuf, String>,
) -> Result<AgentDistribution, String> {
    let mut attempts = Vec::new();

    match packaged_root {
        Ok(candidate) => match AgentDistribution::packaged(candidate.clone()) {
            Ok(distribution) => return Ok(distribution),
            Err(error) => attempts.push(format!("安装资源 {}：{error}", candidate.display())),
        },
        Err(error) => attempts.push(error),
    }

    if let Some(candidate) = explicit_root {
        return AgentDistribution::development(&candidate).map_err(|error| {
            format!(
                "LAMBER_REPO_ROOT 指向的开发运行资源不完整（{}）：{error}",
                candidate.display()
            )
        });
    }

    match executable {
        Ok(executable) => {
            for ancestor in executable.ancestors() {
                if ancestor.join("agent-bridge").join(PATCH_FILE).is_file() {
                    return AgentDistribution::development(ancestor).map_err(|error| {
                        format!("开发运行资源不完整（{}）：{error}", ancestor.display())
                    });
                }
            }
            attempts.push(format!(
                "可执行文件 {} 周边没有开发运行资源",
                executable.display()
            ));
        }
        Err(error) => attempts.push(error),
    }

    Err(format!(
        "AI 运行组件不完整，请重新安装 Lamber。定位结果：{}",
        attempts.join("；")
    ))
}

/// Copy immutable template/plugin files into the user's writable app data and
/// write the model/endpoint overlay consumed on the next dsh launch.
pub fn prepare_user_home(
    app: &AppHandle,
    distribution: &AgentDistribution,
    settings: &AiAgentSettings,
) -> Result<(PathBuf, PathBuf), String> {
    let app_data = app
        .path()
        .app_data_dir()
        .map_err(|error| format!("无法定位应用数据目录：{error}"))?;
    prepare_home_at(&app_data, distribution, settings)
}

pub(super) fn prepare_home_at(
    app_data: &Path,
    distribution: &AgentDistribution,
    settings: &AiAgentSettings,
) -> Result<(PathBuf, PathBuf), String> {
    let home = app_data.join(RUNTIME_HOME);
    let profile_manifest = home.join("profiles/acp/package.json");
    if !profile_manifest.is_file() {
        copy_tree(&distribution.home_template, &home).map_err(|error| {
            format!("初始化 AI 用户运行目录失败（{}）：{error}", home.display())
        })?;
    }

    // The plugin is application code, not user state. Refresh it on every dsh
    // launch so an app upgrade cannot keep executing an older copied plugin.
    let plugin_target = home.join("profiles/acp/node_modules/dsh-tool-lamber");
    fs::create_dir_all(&plugin_target)
        .map_err(|error| format!("创建 Lamber AI 工具目录失败：{error}"))?;
    fs::copy(
        distribution.lamber_plugin.join("package.json"),
        plugin_target.join("package.json"),
    )
    .map_err(|error| format!("复制 Lamber AI 工具清单失败：{error}"))?;
    copy_tree(
        &distribution.lamber_plugin.join("lib"),
        &plugin_target.join("lib"),
    )
    .map_err(|error| {
        format!(
            "同步 Lamber AI 工具失败（{}）：{error}",
            plugin_target.display()
        )
    })?;

    let patch_path = home.join(SETTINGS_PATCH);
    fs::create_dir_all(&home).map_err(|error| format!("创建 AI 用户目录失败：{error}"))?;
    let model = serde_json::to_string(&settings.model)
        .map_err(|error| format!("序列化 AI 模型配置失败：{error}"))?;
    let base_url = serde_json::to_string(&settings.base_url)
        .map_err(|error| format!("序列化 AI 服务地址失败：{error}"))?;
    let patch = format!(
        "- id: llm-deepseek\n  config:\n    baseURL: {base_url}\n\n- id: acp\n  config:\n    provider: deepseek-official\n    model: {model}\n"
    );
    fs::write(&patch_path, patch)
        .map_err(|error| format!("写入 AI 运行配置失败（{}）：{error}", patch_path.display()))?;

    Ok((home, patch_path))
}

pub(super) fn copy_tree(source: &Path, destination: &Path) -> Result<(), String> {
    if !source.is_dir() {
        return Err(format!("源目录不存在：{}", source.display()));
    }
    fs::create_dir_all(destination)
        .map_err(|error| format!("创建目录 {} 失败：{error}", destination.display()))?;
    for entry in fs::read_dir(source)
        .map_err(|error| format!("读取目录 {} 失败：{error}", source.display()))?
    {
        let entry = entry.map_err(|error| format!("读取目录项失败：{error}"))?;
        let name = entry.file_name();
        if name == "node_modules.lock" {
            continue;
        }
        let source_path = entry.path();
        let destination_path = destination.join(name);
        let file_type = entry
            .file_type()
            .map_err(|error| format!("读取文件类型 {} 失败：{error}", source_path.display()))?;
        if file_type.is_dir() {
            copy_tree(&source_path, &destination_path)?;
        } else if file_type.is_file() {
            fs::copy(&source_path, &destination_path).map_err(|error| {
                format!(
                    "复制 {} 到 {} 失败：{error}",
                    source_path.display(),
                    destination_path.display()
                )
            })?;
        } else {
            return Err(format!(
                "运行模板包含不支持的链接：{}",
                source_path.display()
            ));
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn development_distribution_requires_every_runtime_component() {
        let root = std::env::temp_dir().join(format!(
            "lamber-agent-distribution-{}",
            uuid::Uuid::new_v4().simple()
        ));
        let error = AgentDistribution::development(&root).expect_err("empty root must fail");
        assert!(error.contains("dsh-home-template"));
        assert!(error.contains("dsh-tool-lamber"));
        assert!(error.contains("patch.yml"));
    }

    #[test]
    fn missing_packaged_component_reports_reinstall_without_developer_instructions() {
        let root = std::env::temp_dir().join(format!(
            "lamber-agent-incomplete-package-{}",
            uuid::Uuid::new_v4().simple()
        ));
        for file in [
            if cfg!(windows) { "node.exe" } else { "node" },
            DSH_ENTRY,
            PATCH_FILE,
            "dsh-home-template/profiles/acp/package.json",
            "dsh-tool-lamber/lib/index.js",
            "webui/lamber-brand/lib/client.js",
            "webui/lamber-host/host-policy.js",
            "webui/lamber-host/gateway.js",
            "webui/lamber-host/business-presentation.generated.js",
        ] {
            let path = root.join(file);
            fs::create_dir_all(path.parent().expect("runtime file parent"))
                .expect("create runtime directory");
            fs::write(path, "probe").expect("write runtime file");
        }
        fs::remove_file(root.join(PATCH_FILE)).expect("remove one packaged component");

        let error = locate_from_candidates(
            Ok(root.clone()),
            None,
            Ok(root.join("unrelated/bin/lamber")),
        )
        .expect_err("incomplete packaged runtime must fail");
        assert!(error.contains("请重新安装 Lamber"));
        assert!(error.contains("patch.yml"));
        assert!(!error.contains("npm install"));
        assert!(!error.contains("provision"));
        assert!(!error.contains("agent-bridge/ 目录"));
    }

    #[test]
    fn copy_tree_rejects_links_in_runtime_templates() {
        #[cfg(unix)]
        {
            use std::os::unix::fs::symlink;
            let root = std::env::temp_dir().join(format!(
                "lamber-agent-template-{}",
                uuid::Uuid::new_v4().simple()
            ));
            let source = root.join("source");
            let destination = root.join("destination");
            fs::create_dir_all(&source).expect("create source");
            symlink("/tmp", source.join("unexpected-link")).expect("create link");
            let error = copy_tree(&source, &destination).expect_err("link must be rejected");
            assert!(error.contains("不支持的链接"));
        }
    }

    #[test]
    fn clean_template_initialization_writes_profile_plugin_and_settings_patch() {
        let repo_root = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .parent()
            .expect("repo root")
            .to_path_buf();
        let distribution = AgentDistribution::development(&repo_root).expect("dev distribution");
        let app_data = std::env::temp_dir().join(format!(
            "lamber-agent-app-data-{}",
            uuid::Uuid::new_v4().simple()
        ));
        let settings = AiAgentSettings {
            api_key: Some("must-not-be-written".to_string()),
            model: "deepseek-v4-flash-vision-exp".to_string(),
            base_url: "https://example.invalid/v1".to_string(),
        };

        let (home, patch) =
            prepare_home_at(&app_data, &distribution, &settings).expect("prepare home");
        assert!(home.join("profiles/acp/package.json").is_file());
        assert!(home
            .join("profiles/acp/node_modules/dsh-tool-lamber/lib/index.js")
            .is_file());
        assert!(!home.join("profiles/node_modules.lock").exists());
        let patch = fs::read_to_string(patch).expect("settings patch");
        assert!(patch.contains("deepseek-v4-flash-vision-exp"));
        assert!(patch.contains("https://example.invalid/v1"));
        assert!(!patch.contains("must-not-be-written"));
    }
}
