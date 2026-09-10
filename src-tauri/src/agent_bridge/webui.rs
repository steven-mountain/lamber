//! Product WebUI host: official transport and UI, existing Rust business services.
use super::*;
use serde_json::{json, Value};
use std::{
    io::{BufRead, BufReader},
    path::PathBuf,
    process::{Child, Command, Stdio},
    sync::mpsc,
    time::Duration,
};

#[derive(Default)]
pub struct WebUiRuntime {
    inner: Mutex<Option<WebHost>>,
    opening: Mutex<()>,
    cancelled: Arc<Mutex<HashMap<(String, String), std::time::Instant>>>,
    pending: Arc<Mutex<HashMap<String, ApprovalPrompt>>>,
}
struct WebHost {
    child: Child,
    url: tauri::Url,
    cwd: PathBuf,
    settings: AiAgentSettings,
    _bridge: BridgeServer,
}
impl Drop for WebHost {
    fn drop(&mut self) {
        let _ = self.child.kill();
        let _ = self.child.wait();
    }
}
impl WebUiRuntime {
    pub fn stop(&self) -> Result<(), String> {
        *self.inner.lock().map_err(|_| "AI WebUI 运行时不可用")? = None;
        self.pending
            .lock()
            .map_err(|_| "AI 审批队列不可用")?
            .clear();
        Ok(())
    }
    fn start(&self, app: &tauri::AppHandle) -> Result<tauri::Url, String> {
        let runtime = app
            .state::<Arc<crate::workspace::WorkspaceRuntime>>()
            .inner()
            .clone();
        let (workspace, audit_db) = runtime.require_context()?;
        let cwd = std::fs::canonicalize(&workspace.workspace_root).map_err(|e| e.to_string())?;
        let settings = ConfigManager::new(app).load().ai_agent;
        let mut owned = self.inner.lock().map_err(|_| "AI WebUI 运行时不可用")?;
        if let Some(host) = owned.as_mut() {
            if host.cwd == cwd
                && host.settings.base_url == settings.base_url
                && host.settings.api_key == settings.api_key
                && host.child.try_wait().map_err(|e| e.to_string())?.is_none()
            {
                app.state::<Arc<AgentRuntime>>().gate.reopen();
                return Ok(host.url.clone());
            }
        }
        let agent = app.state::<Arc<AgentRuntime>>().inner().clone();
        agent.gate.shutdown();
        *owned = None;
        agent.gate.reopen();
        let distribution = locate_distribution(app)?;
        let (home, patch) = prepare_web_home(app, &distribution, &settings)?;
        let bindings = Arc::new(project_bindings::ProjectBindings::default());
        agent.gate.set_recorder(approval_log::bound_workspace_recorder(
            audit_db,
            cwd.join(approval_log::WORKSPACE_SPOOL_FILE),
        ));
        let changed = app.clone();
        let services = benefit_simulation::handler(
            runtime.clone(),
            bindings.clone(),
            benefit_simulation::app_preparer(app.clone(), agent.simulations.clone()),
            workspace_handler(runtime.clone(), bindings.clone()),
        );
        let services = template_write::handler(
            runtime.clone(),
            bindings.clone(),
            agent.gate.clone(),
            Arc::new(move |receipt| {
                let _ = changed.emit(template_write::CHANGED_EVENT, receipt);
            }),
            agent.gate.reviewed.handler(services),
        );
        let bridge = BridgeServer::start(handler(
            app.clone(),
            workspace.workspace_id.clone(),
            runtime,
            bindings,
            agent.gate.clone(),
            self.pending.clone(),
            self.cancelled.clone(),
            services,
        ))?;
        let host = WebHost::launch(&distribution, &home, &patch, cwd, &settings, bridge)?;
        let url = host.url.clone();
        *owned = Some(host);
        Ok(url)
    }
}
impl WebHost {
    fn launch(
        distribution: &distribution::AgentDistribution,
        home: &std::path::Path,
        patch: &std::path::Path,
        cwd: PathBuf,
        settings: &AiAgentSettings,
        bridge: BridgeServer,
    ) -> Result<Self, String> {
        let mut command = Command::new(&distribution.node_bin);
        command
            .arg(&distribution.dsh_entry)
            .args(["--profile", "web", "--patch"])
            .arg(&patch)
            .args(["--no-open", "--host", "127.0.0.1", "--port", "0"])
            .current_dir(&cwd)
            .env("DSH_HOME", &home)
            .env("DSH_TELEMETRY_MODE", "DISABLED")
            .env("LAMBER_PARENT_PID", std::process::id().to_string())
            .env("LAMBER_BRIDGE_URL", bridge.origin())
            .env("LAMBER_BRIDGE_TOKEN", bridge.token())
            .env(
                "DEEPSEEK_API_KEY",
                settings.api_key.as_deref().unwrap_or(""),
            )
            .env_remove("LAMBER_STREAM_DISPLAY")
            .stdin(Stdio::null())
            .stdout(Stdio::piped())
            .stderr(Stdio::piped());
        #[cfg(windows)]
        {
            use std::os::windows::process::CommandExt;
            command.creation_flags(0x08000000);
        }
        let mut child = command
            .spawn()
            .map_err(|e| format!("无法启动 AI WebUI：{e}"))?;
        let (tx, rx) = mpsc::channel();
        let diagnostics = Arc::new(Mutex::new(std::collections::VecDeque::new()));
        let secrets = vec![
            bridge.token().to_string(),
            settings.api_key.clone().unwrap_or_default(),
        ];
        for stream in [
            child
                .stdout
                .take()
                .map(|s| Box::new(s) as Box<dyn std::io::Read + Send>),
            child
                .stderr
                .take()
                .map(|s| Box::new(s) as Box<dyn std::io::Read + Send>),
        ]
        .into_iter()
        .flatten()
        {
            let tx = tx.clone();
            let diagnostics = diagnostics.clone();
            let secrets = secrets.clone();
            std::thread::spawn(move || {
                for line in BufReader::new(stream).lines().map_while(Result::ok) {
                    for word in line.split_whitespace() {
                        if let Ok(url) = tauri::Url::parse(word) {
                            if url.scheme() == "http"
                                && url.host_str() == Some("127.0.0.1")
                                && url.port().is_some()
                                && url.path() == "/"
                                && url.query_pairs().any(|(key, _)| key == "token")
                            {
                                let _ = tx.send(url);
                            }
                        }
                    }
                    // Keep bounded actionable diagnostics; credentials/launch URLs are never exposed.
                    if !line.contains("token=")
                        && (line.contains("Error")
                            || line.contains("error")
                            || line.contains("failed")
                            || line.contains("Cannot")
                            || line.contains("unknown"))
                    {
                        let mut safe = line;
                        for secret in &secrets {
                            if !secret.is_empty() {
                                safe = safe.replace(secret, "[redacted]");
                            }
                        }
                        if let Ok(mut lines) = diagnostics.lock() {
                            lines.push_back(safe);
                            if lines.len() > 12 {
                                lines.pop_front();
                            }
                        }
                    }
                }
            });
        }
        drop(tx);
        let url = match rx.recv_timeout(Duration::from_secs(30)) {
            Ok(url) => url,
            Err(_) => {
                let _ = child.kill();
                let _ = child.wait();
                let detail = diagnostics
                    .lock()
                    .map(|lines| lines.iter().cloned().collect::<Vec<_>>().join("\n"))
                    .unwrap_or_default();
                return Err(format!(
                    "AI WebUI 启动失败，请检查运行资源和模型设置后重试\n{detail}"
                ));
            }
        };
        Ok(WebHost {
            child,
            url,
            cwd,
            settings: settings.clone(),
            _bridge: bridge,
        })
    }
}

fn binding_view(
    app: &tauri::AppHandle,
    runtime: &crate::workspace::WorkspaceRuntime,
    session: &str,
) -> Result<Value, String> {
    let mut store = open_session_store(app)?;
    let key = format!("web:{session}");
    runtime.with_locked_context(|workspace, conn| {
        let cwd = std::fs::canonicalize(&workspace.workspace_root).map_err(|e| e.to_string())?;
        let binding = store
            .binding(&key)?
            .or(store.binding_for_harness(session, &cwd)?);
        let Some(binding) = binding else {
            return Ok(Value::Null);
        };
        if binding.cwd != cwd || binding.workspace_id != workspace.workspace_id {
            return Err("此会话属于另一个工作区".into());
        }
        let project = binding
            .project_id
            .as_deref()
            .map(|id| crate::project_state::get_project_locked(conn, id))
            .transpose()?
            .flatten();
        if binding.project_id.is_some() && project.is_none() {
            return Err("绑定项目已不存在".into());
        }
        store.bind(&key, &binding)?;
        Ok(
            json!({ "workspaceId": binding.workspace_id, "projectId": binding.project_id,
            "projectName": project.map(|p| p.name) }),
        )
    })
}
fn handler(
    app: tauri::AppHandle,
    workspace_id: String,
    runtime: Arc<crate::workspace::WorkspaceRuntime>,
    bindings: Arc<project_bindings::ProjectBindings>,
    gate: Arc<ApprovalGate>,
    pending: Arc<Mutex<HashMap<String, ApprovalPrompt>>>,
    cancelled: Arc<Mutex<HashMap<(String, String), std::time::Instant>>>,
    fallback: BridgeHandler,
) -> BridgeHandler {
    Arc::new(move |path, body| {
        let Some(method) = path.strip_prefix("/lamber-webui/") else {
            return fallback(path, body);
        };
        let result = (|| -> Result<Value, String> {
            if runtime.require_workspace()?.workspace_id != workspace_id {
                return Err("工作区已变更，请重新打开 AI 窗口".into());
            }
            let args: Value = serde_json::from_str(body).map_err(|_| "WebUI 请求格式错误")?;
            let session = args["sessionId"].as_str().unwrap_or("");
            match method {
                "health" => Ok(json!({"alive":true})),
                "bootstrap" => runtime.with_locked_context(|workspace, conn| {
                    let mut query = conn.prepare("SELECT id, name FROM projects ORDER BY name").map_err(|e| e.to_string())?;
                    let projects = query.query_map([], |row| Ok(json!({"id":row.get::<_,String>(0)?,"name":row.get::<_,String>(1)?})))
                        .map_err(|e|e.to_string())?.collect::<Result<Vec<_>,_>>().map_err(|e|e.to_string())?;
                    Ok(json!({"cwd": std::fs::canonicalize(&workspace.workspace_root).map_err(|e|e.to_string())?,
                        "workspaceName": workspace.workspace_name, "projects": projects}))
                }),
                "legacy-status" => webui_history::status(&app),
                "selected-session" => Ok(json!({"sessionId":open_session_store(&app)?.selected_web_session(&workspace_id)?})),
                "select-session" => {
                    open_session_store(&app)?.select_web_session(&workspace_id,session)?;
                    Ok(Value::Null)
                },
                "binding" => binding_view(&app, &runtime, session),
                "bind" => {
                    if session.is_empty() { return Err("缺少会话身份".into()); }
                    runtime.with_locked_context(|workspace, conn| {
                        let project = args.get("projectId").ok_or("请选择项目或显式通用聊天")?;
                        let project_id = if project.is_null() { None } else { Some(project.as_str().ok_or("项目身份无效")?.to_string()) };
                        if let Some(id) = &project_id {
                            crate::project_state::get_project_locked(conn,id)?.ok_or("项目不存在")?;
                        }
                        let binding = project_bindings::ProjectBinding { workspace_id:workspace.workspace_id.clone(),
                            cwd:std::fs::canonicalize(&workspace.workspace_root).map_err(|e|e.to_string())?, project_id };
                        let key = format!("web:{session}");
                        open_session_store(&app)?.bind(&key, &binding)?;
                        bindings.register(session,binding)
                    })?;
                    binding_view(&app,&runtime,session)
                },
                "inherit" => {
                    let parent = args["parentSessionId"].as_str().ok_or("缺少原会话身份")?;
                    if parent == session || session.is_empty() { return Err("分支身份无效".into()); }
                    binding_view(&app,&runtime,parent)?;
                    let binding = open_session_store(&app)?.binding(&format!("web:{parent}"))?.ok_or("原会话没有可信绑定")?;
                    if args["cwd"].as_str() != binding.cwd.to_str() { return Err("分支目录不一致".into()); }
                    open_session_store(&app)?.bind(&format!("web:{session}"),&binding)?;
                    bindings.register(session,binding)?;
                    binding_view(&app,&runtime,session)
                },
                "admit" => {
                    let view = binding_view(&app,&runtime,session)?;
                    if view.is_null() { return Err("请先在输入区下方选择项目或通用聊天，再发送消息".into()); }
                    let binding = open_session_store(&app)?.binding(&format!("web:{session}"))?.ok_or("缺少可信绑定")?;
                    if args["cwd"].as_str() != binding.cwd.to_str() { return Err("会话目录与可信绑定不一致".into()); }
                    bindings.register(session,binding)?;
                    Ok(view)
                },
                "cancel-approval" => {
                    let call=args["callId"].as_str().ok_or("缺少工具调用身份")?;
                    {
                        let mut aborted=cancelled.lock().map_err(|_|"审批取消状态不可用")?;
                        aborted.retain(|_,at|at.elapsed()<Duration::from_secs(610));
                        aborted.insert((session.into(),call.into()),std::time::Instant::now());
                    }
                    let ids=pending.lock().map_err(|_|"审批队列不可用")?.values()
                        .filter(|p|p.session_id.as_deref()==Some(session)&&p.call_id.as_deref()==Some(call))
                        .map(|p|p.request_id.clone()).collect::<Vec<_>>();
                    for id in ids { let _=gate.cancel_request(&id,"本次生成已停止，未完成的审批已拒绝"); }
                    Ok(Value::Null)
                },
                "approval" => {
                    if app.get_webview_window("ai-assistant").is_none() { return Err("AI 窗口已关闭，本次写入已拒绝".into()); }
                    let question = approval::ApprovalQuestion { session_id:Some(session.into()), call_id:args["callId"].as_str().map(str::to_owned),
                        tool_name:args["toolName"].as_str().ok_or("缺少工具名称")?.into(), args:args["args"].clone(), reason:None, intent:None };
                    let mut request = None;
                    let result = template_write::request_approval(&gate,&runtime,&bindings,question, |prompt| {
                        request = Some(prompt.request_id.clone());
                        if let Ok(mut queue) = pending.lock() { queue.insert(prompt.request_id.clone(),prompt.clone()); }
                        let key=(session.to_string(),args["callId"].as_str().unwrap_or("").to_string());
                        if cancelled.lock().map(|aborted|aborted.contains_key(&key)).unwrap_or(true) {
                            let _=gate.cancel_request(&prompt.request_id,"本次生成已停止，未完成的审批已拒绝");
                        }

                    });
                    if let Some(id) = request { pending.lock().map_err(|_| "审批队列不可用")?.remove(&id); }
                    if let Ok(mut aborted)=cancelled.lock() { aborted.remove(&(session.to_string(),args["callId"].as_str().unwrap_or("").to_string())); }
                    serde_json::to_value(result).map_err(|e|e.to_string())
                },
                "pending" => {
                    let mut queue=pending.lock().map_err(|_|"审批队列不可用")?.values().cloned().collect::<Vec<_>>();
                    queue.sort_by(|a,b|a.expires_at.cmp(&b.expires_at).then(a.request_id.cmp(&b.request_id)));
                    Ok(json!(queue))
                },
                "resolve" => {
                    let id = args["requestId"].as_str().ok_or("缺少审批身份")?;
                    let queue = pending.lock().map_err(|_| "审批队列不可用")?;
                    let prompt = queue.get(id).ok_or("审批已结束或过期")?;
                    let session = prompt.session_id.as_deref().ok_or("审批缺少会话")?;
                    binding_view(&app,&runtime,session)?;
                    gate.resolve_with_args(id,args["approved"].as_bool().ok_or("缺少明确决定")?, args.get("modifiedArgs").filter(|v|!v.is_null()).cloned())?;
                    Ok(Value::Null)
                },
                "legacy-list" => webui_history::list(&app,&runtime),
                "legacy-read" => webui_history::read(&app,&runtime,args["id"].as_str().ok_or("缺少旧记录身份")?),
                "legacy-context" => { binding_view(&app,&runtime,session)?; webui_history::receipt_context(&app,session) },
                "legacy-restore" => {
                    let id=args["id"].as_str().ok_or("缺少旧记录身份")?;
                    webui_history::read(&app,&runtime,id)?;
                    let workspace=runtime.require_workspace()?;
                    let cwd=std::fs::canonicalize(workspace.workspace_root).map_err(|e|e.to_string())?;
                    let session=open_session_store(&app)?.get(id,&cwd)?.ok_or("没有可信映射，只能读取旧记录")?;
                    let binding=binding_view(&app,&runtime,&session)?;
                    if binding.is_null() { return Err("没有可信绑定，请另建会话".into()); }
                    Ok(json!({"sessionId":session}))
                },
                "business" => {
                    let session = session.strip_prefix("web:").ok_or("业务会话身份无效")?;
                    let binding = binding_view(&app,&runtime,session)?;
                    if binding.is_null() || (binding["projectId"].is_null() && args["method"]!="context") { return Err("此业务操作需要绑定项目".into()); }
                    let payload = &args["payload"];
                    let key = format!("web:{session}");
                    if payload["sessionId"] != key { return Err("业务请求会话不一致".into()); }
                    for target in [payload.get("target"),payload.get("input")].into_iter().flatten() {
                        if target["sessionId"] != key || target["projectId"] != binding["projectId"] || target["workspaceId"] != binding["workspaceId"] {
                            return Err("业务目标与可信绑定不一致".into());
                        }
                    }
                    let id=args["requestId"].as_str().ok_or("缺少业务请求身份")?;
                    let method=args["method"].as_str().ok_or("缺少业务操作")?;
                    let store=webui_actions::store(&app)?;
                    let expires_at=store.begin(id,&key,method)?;
                    let event=json!({"requestId":id,"method":method,"payload":payload,"expiresAt":expires_at});
                    if let Err(error)=app.emit_to("main",webui_actions::REQUEST_EVENT,event) {
                        store.finish(id,&json!({"ok":false,"error":format!("主窗口未收到操作：{error}")}))?;
                    }
                    Ok(json!({"requestId":id}))
                },
                "business-result" | "receipts" | "release-read" => {
                    let session = session.strip_prefix("web:").ok_or("业务会话身份无效")?;
                    binding_view(&app,&runtime,session)?;
                    let key=format!("web:{session}");
                    let store=webui_actions::store(&app)?;
                    if method=="receipts" { store.receipts(&key) }
                    else {
                        let id=args["requestId"].as_str().ok_or("缺少业务请求身份")?;
                        let result=store.read(id,&key)?;
                        if method=="release-read" { store.forget_read(id)?; }
                        Ok(result)
                    }
                },
                "select-model" => {
                    let model = args["model"].as_str().ok_or("缺少模型")?;
                    if !AI_AGENT_MODELS.iter().any(|item| item.id == model) { return Err("此部署不支持该模型".into()); }
                    let manager = ConfigManager::new(&app);
                    let mut config = manager.load(); config.ai_agent.model = model.into(); manager.save(&config)?;
                    Ok(Value::Null)
                },
                "save-settings" => {
                    let update:AiAgentSettingsUpdate=serde_json::from_value(args).map_err(|_|"模型设置格式无效")?;
                    serde_json::to_value(save_settings(&app,update)?).map_err(|e|e.to_string())
                },
                "restart" => {
                    let restarting=app.clone();
                    tauri::async_runtime::spawn(async move {
                        if let Err(error) = restart_webui(restarting.clone()).await {
                            let _ = restarting.emit_to("main", "lamber-ai-startup-error", error);
                            if let Some(main) = restarting.get_webview_window("main") { let _ = main.show(); let _ = main.set_focus(); }
                        }
                    });
                    Ok(Value::Null)
                },
                "settings" => serde_json::to_value(AiAgentSettingsView::from(ConfigManager::new(&app).load().ai_agent)).map_err(|e|e.to_string()),
                _ => Err("未开放此 WebUI 操作".into()),
            }
        })();
        match result {
            Ok(value) => BridgeReply::ok(value.to_string()),
            Err(error) => BridgeReply::error(422, &error),
        }
    })
}

fn prepare_web_home(
    app: &tauri::AppHandle,
    distribution: &distribution::AgentDistribution,
    settings: &AiAgentSettings,
) -> Result<(PathBuf, PathBuf), String> {
    let app_data = app.path().app_data_dir().map_err(|e| e.to_string())?;
    prepare_web_home_at(&app_data, distribution, settings)
}
fn prepare_web_home_at(
    app_data: &std::path::Path,
    distribution: &distribution::AgentDistribution,
    settings: &AiAgentSettings,
) -> Result<(PathBuf, PathBuf), String> {
    let (home, _) = distribution::prepare_home_at(app_data, distribution, settings)?;
    let profile = home.join("profiles/web");
    std::fs::create_dir_all(&profile).map_err(|e| e.to_string())?;
    std::fs::write(profile.join("package.json"), r#"{"name":"dsh-profile-web","private":true,"dsh":{"profile":{"bundles":["@deepseek-ai/dsh-base","@deepseek-ai/dsh-web-app"],"patchReload":"startup"}}}"#).map_err(|e|e.to_string())?;
    std::fs::write(profile.join("cordis.yml"), "[]\n").map_err(|e| e.to_string())?;
    let tool_target = profile.join("node_modules/dsh-tool-lamber");
    std::fs::create_dir_all(&tool_target).map_err(|e| e.to_string())?;
    std::fs::copy(
        distribution.lamber_plugin.join("package.json"),
        tool_target.join("package.json"),
    )
    .map_err(|e| e.to_string())?;
    distribution::copy_tree(
        &distribution.lamber_plugin.join("lib"),
        &tool_target.join("lib"),
    )?;
    let brand = distribution
        .base_patch
        .parent()
        .ok_or("缺少运行根目录")?
        .join("webui/lamber-brand");
    distribution::copy_tree(
        &brand,
        &profile.join("node_modules/dsh-client-ui-lamber-brand"),
    )?;
    distribution::copy_tree(
        &brand.parent().unwrap().join("lamber-host"),
        &profile.join("node_modules/dsh-lamber-web-host"),
    )?;
    let preset = home.join("lamber-presets/lamber");
    std::fs::create_dir_all(&preset).map_err(|e| e.to_string())?;
    std::fs::write(
        preset.join("preset.yml"),
        "name: Lamber\ndescription: Lamber 业务助手\n",
    )
    .map_err(|e| e.to_string())?;
    std::fs::write(preset.join("agent.cordis.yml"), "- id: persona\n  name: '@deepseek-ai/dsh-persona'\n  config:\n    text: 你是 Lamber 的中文业务助手。项目权限以可信上下文为准；先读模板再提议改文。所有写入必须经用户审核，金额、税率、计划和图片不能直接修改。\n    complete: true\n    includeRuntimeContext: false\n").map_err(|e|e.to_string())?;
    let patch = home.join("lamber-web.patch.yml");
    let text = format!("- id: llm-deepseek\n  config:\n    baseURL: {}\n- insert:\n    - id: lamber-default-model\n      name: dsh-lamber-web-host/default-model\n      config:\n        provider: deepseek-official\n        model: {}\n- id: agent-presets\n  config:\n    default: lamber\n    includeShippedRoot: false\n    includeUserRoot: false\n    roots:\n      - path: {}\n        trust: system\n",json!(settings.base_url),json!(settings.model),json!(home.join("lamber-presets")));
    let disabled = [
        "typert-gateway",
        "session-controller",
        "workspace-controller",
        "agent-default-model",
        "ui-brand-official",
        "directory-picker",
        "cordis-host-runner",
        "cordis-client-runner",
        "ui-cordis",
        "plugin-inventory",
        "ui-settings-plugins",
        "ui-settings-models",
        "ui-permission",
        "ui-agent-preset",
        "file-reference-local",
        "agent-instructions",
        "ui-skill",
        "ui-goal",
        "ui-jobs",
        "ui-schedule",
    ];
    let mut text = text;
    for id in disabled {
        text.push_str(&format!("- id: {id}\n  disabled: true\n"));
    }
    text.push_str("- id: permission\n  config:\n    defaultPreset: workspace-write\n    presets:\n      workspace-write:\n        sandbox: workspace-write\n        approval: ask\n        name: 写入需审核\n        description: 模板改文需用户审核；金额、计划和图片仅通过业务卡片确认。\n");
    // A dual-face upstream row owns both the Host and the browser module. Mount
    // a deployment entry with the original client identity and byte-identical
    // client artifact, but our Host export. Never patch installed upstream files.
    for (adapter, upstream) in [
        ("gateway", "dsh-api-gateway"),
        ("session-controller", "dsh-api-session-controller"),
        ("workspace-controller", "dsh-api-workspace-controller"),
    ] {
        let source = distribution.base_patch.parent().unwrap()
            .join("node_modules/@deepseek-ai").join(upstream);
        let manifest: serde_json::Value = serde_json::from_slice(
            &std::fs::read(source.join("package.json")).map_err(|e| e.to_string())?
        ).map_err(|e| e.to_string())?;
        let entry = profile.join("lamber-client-faces").join(adapter);
        std::fs::create_dir_all(&entry).map_err(|e| e.to_string())?;
        std::fs::write(entry.join("package.json"), json!({
            "name": manifest["name"], "version": manifest["version"],
            "private": true, "type": "module", "main": "index.js",
            "exports": {".": "./index.js", "./client": "./client.js"},
            "dsh": {"client": manifest["dsh"]["client"]}
        }).to_string()).map_err(|e| e.to_string())?;
        std::fs::copy(source.join("lib/client.js"), entry.join("client.js"))
            .map_err(|e| e.to_string())?;
        std::fs::write(entry.join("index.js"), format!(
            "export {{ default }} from 'dsh-lamber-web-host/{adapter}';\n"
        )).map_err(|e| e.to_string())?;
        let module_url = tauri::Url::from_file_path(entry.join("index.js"))
            .map_err(|_| "部署插件路径无效")?;
        text.push_str(&format!("- insert:\n    - id: lamber-{adapter}\n      name: {}\n", json!(module_url.as_str())));
    }
    text.push_str("- insert:\n    - id: lamber-tools\n      name: dsh-tool-lamber\n    - id: lamber-brand\n      name: dsh-client-ui-lamber-brand\n    - id: lamber-web-policy\n      name: dsh-lamber-web-host/host-policy\n");
    std::fs::write(&patch, text).map_err(|e| e.to_string())?;
    Ok((home, patch))
}

/// Serialize native workspace activation with Host creation. Approval callbacks
/// retain the original database; no wait, heartbeat or frontend event is needed.
pub(crate) fn workspace_transition<T>(app: &tauri::AppHandle, transition: impl FnOnce() -> Result<T, String>) -> Result<T, String> {
    let web = app.try_state::<Arc<WebUiRuntime>>();
    let _opening = web.as_ref().map(|state| state.opening.lock().map_err(|_| "AI 窗口创建不可用")).transpose()?;
    if let Some(agent) = app.try_state::<Arc<AgentRuntime>>() {
        agent.stop()?;
        agent.gate.shutdown();
        agent.gate.clear_recorder();
    }
    if let Some(web) = web.as_ref() { web.stop()?; }
    // Keep the window reusable: the next normal AI click navigates to the new Host.
    if let Some(window) = app.get_webview_window("ai-assistant") { window.hide().map_err(|e| e.to_string())?; }
    transition()
}

#[tauri::command]
pub async fn ai_open_webui(app: tauri::AppHandle) -> Result<(), String> {
    tauri::async_runtime::spawn_blocking(move || open_window(&app, false))
        .await
        .map_err(|e| e.to_string())?
}
pub async fn restart_webui(app: tauri::AppHandle) -> Result<(), String> {
    tauri::async_runtime::spawn_blocking(move || open_window(&app, true))
        .await
        .map_err(|e| e.to_string())?
}
fn open_window(app: &tauri::AppHandle, restart: bool) -> Result<(), String> {
    let state = app.state::<Arc<WebUiRuntime>>();
    let _opening = state.opening.lock().map_err(|_| "AI 窗口创建不可用")?;
    if restart {
        app.state::<Arc<AgentRuntime>>().gate.shutdown();
        state.stop()?;
    }
    let url = state.start(app)?;
    if let Some(window) = app.get_webview_window("ai-assistant") {
        if window.url().map_err(|e| e.to_string())?.origin() != url.origin() {
            window.navigate(url).map_err(|e| e.to_string())?;
        }
        window.show().map_err(|e| e.to_string())?;
        return window.set_focus().map_err(|e| e.to_string());
    }
    let navigation = app.clone();
    let window =
        tauri::WebviewWindowBuilder::new(app, "ai-assistant", tauri::WebviewUrl::External(url))
            .title("Lamber AI 工作区")
            .inner_size(1000.0, 760.0)
            .min_inner_size(400.0, 480.0)
            .on_navigation(move |target| {
                navigation
                    .state::<Arc<WebUiRuntime>>()
                    .inner
                    .lock()
                    .ok()
                    .and_then(|host| {
                        host.as_ref()
                            .map(|host| host.url.origin() == target.origin())
                    })
                    .unwrap_or(false)
            })
            .build()
            .map_err(|e| e.to_string())?;
    let closing = app.clone();
    window.on_window_event(move |event| {
        if matches!(event, tauri::WindowEvent::CloseRequested { .. }) {
            closing.state::<Arc<AgentRuntime>>().gate.shutdown();
        }
        if matches!(event, tauri::WindowEvent::Destroyed) {
            closing.state::<Arc<AgentRuntime>>().gate.shutdown();
            let _ = closing.state::<Arc<WebUiRuntime>>().stop();
        }
    });
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn official_web_host_starts_with_product_plugins_and_reaps_child() {
        let root = std::env::temp_dir().join(format!("lamber-web-host-{}", uuid::Uuid::new_v4()));
        let cwd = root.join("合成工作区 With Spaces");
        std::fs::create_dir_all(&cwd).unwrap();
        let repo = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .parent()
            .unwrap();
        let distribution = distribution::AgentDistribution::development(repo).unwrap();
        let settings = AiAgentSettings {
            api_key: None,
            base_url: "http://127.0.0.1:1".into(),
            model: "deepseek-v4-flash".into(),
        };
        let (home, patch) = prepare_web_home_at(&root, &distribution, &settings).unwrap();
        let contract = root.join("host-contract.mjs");
        std::fs::write(&contract, include_str!("../../../scripts/fixtures/webui-host-contract.mjs")).unwrap();
        let mut patch_text = std::fs::read_to_string(&patch).unwrap();
        patch_text.push_str(&format!("- insert:\n    - id: test-host-contract\n      name: {}\n", json!(tauri::Url::from_file_path(&contract).unwrap().as_str())));
        std::fs::write(&patch, patch_text).unwrap();
        let fixture_cwd = cwd.clone();
        let bridge = BridgeServer::start(Arc::new(move |path, _body| {
            if path == "/lamber-webui/bootstrap" {
                BridgeReply::ok(
                    json!({"cwd":fixture_cwd,"workspaceName":"合成工作区","projects":[]})
                        .to_string(),
                )
            } else if path == "/lamber-webui/health" {
                BridgeReply::ok(json!({"alive":true}).to_string())
            } else {
                BridgeReply::error(403, "测试不开放业务操作")
            }
        }))
        .unwrap();
        let mut host =
            WebHost::launch(&distribution, &home, &patch, cwd, &settings, bridge).unwrap();
        assert!(host.child.try_wait().unwrap().is_none());
        assert!(root.join("host-contract-passed.json").is_file(), "effective Gateway and stream contract did not finish");
        let browser_contract = Command::new(&distribution.node_bin)
            .args(["--input-type=module", "-e", r#"
                import assert from 'node:assert/strict';
                const url = new URL(process.env.LAMBER_TEST_WEB_URL);
                const auth = await fetch(url, {redirect:'manual'});
                const cookie = auth.headers.getSetCookie().map(x=>x.split(';')[0]).join('; ');
                const response = auth.status === 302 || auth.status === 303
                    ? await fetch(new URL(auth.headers.get('location'),url), {headers:{cookie}}) : auth;
                const html = await response.text();
                assert.equal(response.status,200);
                for (const name of ['dsh-api-gateway','dsh-api-session-controller','dsh-api-workspace-controller','dsh-client-ui-conversation']) {
                    assert.ok(html.includes('"id":"@deepseek-ai/'+name+'"'), 'missing original browser module '+name);
                }
            "#])
            .env("LAMBER_TEST_WEB_URL", host.url.as_str())
            .output().unwrap();
        assert!(browser_contract.status.success(), "{}", String::from_utf8_lossy(&browser_contract.stderr));
        for (adapter, upstream) in [("gateway", "dsh-api-gateway"), ("session-controller", "dsh-api-session-controller"), ("workspace-controller", "dsh-api-workspace-controller")] {
            assert_eq!(std::fs::read(home.join("profiles/web/lamber-client-faces").join(adapter).join("client.js")).unwrap(),
                std::fs::read(repo.join("agent-bridge/node_modules/@deepseek-ai").join(upstream).join("lib/client.js")).unwrap());
        }
        let port = host.url.port().unwrap();
        drop(host);
        assert!(std::net::TcpStream::connect(("127.0.0.1", port)).is_err());
        std::fs::remove_dir_all(root).unwrap();
    }
}
