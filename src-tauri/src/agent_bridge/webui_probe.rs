//! Stage 0 only: run the complete upstream WebUI in an isolated native window.
//! No application configuration, workspace, business commands or IPC grants are
//! installed here. The verification runner owns the isolated Host lifecycle.

use tauri::{WebviewUrl, WebviewWindowBuilder};

fn probe_url(raw: &str) -> Result<tauri::Url, String> {
    let url = tauri::Url::parse(raw).map_err(|_| "Invalid WebUI probe URL")?;
    if url.scheme() != "http"
        || url.host_str() != Some("127.0.0.1")
        || url.port().is_none()
        || !url.username().is_empty()
        || url.password().is_some()
        || url.path() != "/"
        || url.fragment().is_some()
    {
        return Err("WebUI probe requires an explicit IPv4 loopback origin".into());
    }
    Ok(url)
}

pub fn run(raw: &str, mut context: tauri::Context<tauri::Wry>) -> Result<(), String> {
    let url = probe_url(raw)?;
    let origin = url.origin();
    // Prevent restoration of the regular application and its webview storage.
    context.config_mut().identifier = "com.lamber.webui-probe".into();
    context.config_mut().app.windows.clear();
    context.config_mut().app.with_global_tauri = false;
    tauri::Builder::default()
        .setup(move |app| {
            WebviewWindowBuilder::new(app, "webui-probe", WebviewUrl::External(url.clone()))
                .title("Lamber AI · 官方 WebUI 验证")
                .inner_size(1120.0, 820.0)
                .min_inner_size(420.0, 480.0)
                .on_navigation(move |target| target.origin() == origin)
                .build()?;
            Ok(())
        })
        .run(context)
        .map_err(|error| format!("WebUI probe failed: {error}"))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn probe_accepts_only_an_explicit_loopback_host() {
        assert!(probe_url("http://127.0.0.1:54321/?token=synthetic").is_ok());
        for url in [
            "https://example.com/",
            "http://localhost:3000/",
            "http://127.0.0.1/",
            "http://127.0.0.1:3000/api",
            "http://user@127.0.0.1:3000/",
            "http://127.0.0.1.example.com:3000/",
            "file:///tmp/index.html",
        ] {
            assert!(probe_url(url).is_err(), "must reject {url}");
        }
    }
}
