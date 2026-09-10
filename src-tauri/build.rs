fn main() {
    println!("cargo:rerun-if-changed=../agent-bridge/bridge-contract.json");
    println!("cargo:rerun-if-changed=../scripts/bridge-contract.mjs");
    for plugin in [
        "../agent-bridge/dsh-tool-lamber",
        "resources/agent-runtime/dsh-tool-lamber",
    ] {
        let compiled = std::path::Path::new(plugin).join("lib/contract.generated.js");
        println!("cargo:rerun-if-changed={}", compiled.display());
        // Development requires the source plugin; staged resources, when
        // present, must also match before Tauri is allowed to bundle them.
        if plugin.starts_with("../") || std::path::Path::new(plugin).exists() {
            let node = std::env::var_os("LAMBER_NODE_BIN").unwrap_or_else(|| "node".into());
            let status = std::process::Command::new(node)
                .args([
                    "../scripts/bridge-contract.mjs",
                    plugin,
                    "../agent-bridge/bridge-contract.json",
                ])
                .status()
                .expect("无法校验 AI 组件，请先安装构建所需的 Node");
            assert!(
                status.success(),
                "AI 组件契约不匹配，请先构建插件并重新准备分发资源"
            );
        }
    }
    tauri_build::build()
}
