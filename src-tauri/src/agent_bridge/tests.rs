//! Integration coverage for the agent bridge.
//!
//! Two layers, deliberately separated so a failure says which half broke:
//!
//! * `calculate_route_*` — the bridge server and calculation route on their own,
//!   over a temporary workspace database. No Node, no network, always runs.
//! * `dsh_*` — the full loop through a real `dsh` child process. These are
//!   `#[ignore]`d because they need `agent-bridge/` provisioned (`npm install`
//!   plus `npm run provision`); run them with
//!   `cargo test agent_bridge -- --ignored --nocapture`.

use super::bridge_server::{BridgeReply, BridgeServer, BRIDGE_TOKEN_HEADER};
use super::calculation::{run_calculation, CalculateRequest, CALCULATE_ROUTE};
use super::dsh_session::{DshLaunchConfig, DshSession};
use crate::benefit::models::{
    BenefitAnalysisScheme, BenefitAnalysisSnapshot, IctInput, IctItem, IctResult,
};
use crate::benefit::repository::{ProjectRepository, SqliteProjectRepository};
use crate::benefit::service::ProjectService;
use serde_json::Value;
use std::sync::atomic::{AtomicUsize, Ordering};
use std::sync::{Arc, Condvar, Mutex};
use std::time::{Duration, Instant};

// ---------------------------------------------------------------- fixtures --

fn temp_db_path(name: &str) -> std::path::PathBuf {
    std::env::temp_dir().join(format!(
        "lamber-agent-bridge-{}-{}.db",
        name,
        uuid::Uuid::new_v4().simple()
    ))
}

fn item(incl_tax: &str, tax_rate: &str) -> IctItem {
    IctItem {
        incl_tax: incl_tax.to_string(),
        tax_rate: tax_rate.to_string(),
        custom_subject_name: None,
        billing_subject_name: None,
    }
}

fn zero_item() -> IctItem {
    item("0", "0.06")
}

/// A minimal but realistic single-year ICT project: integration revenue against
/// device cost, so NPV and 利润率 are both non-zero and easy to eyeball.
fn sample_input(project_name: &str) -> IctInput {
    let mut distribution = vec![0.0; 10];
    distribution[0] = 1.0;
    IctInput {
        project_name: project_name.to_string(),
        customer_name: Some("测试客户".to_string()),
        property_rights: "self".to_string(),
        discount_rate: "0.055".to_string(),
        project_years: Some(1),
        cashflow_model: Some("A".to_string()),
        cashflow_calculation_source: None,
        cashflow_segment_value_mode: None,
        cashflow_segments: None,
        subject_funding_plans: None,
        subject_funding_plan_migration_version: None,
        project_background: None,
        revenue_balance_rule: None,
        investment_balance_rule: None,
        ignore_tail_difference: None,
        tail_difference_value: None,
        rev_distribution: distribution.clone(),
        cost_distribution: distribution,
        rev_cashflow_excl: None,
        cost_cashflow_excl: None,
        it_rev_cashflow_excl: None,
        it_cost_cashflow_excl: None,
        selection_fee_quote: None,
        selection_fee_markup: None,
        selection_fee_actual_cost: None,
        selection_fee_amount: None,
        selection_fee_limit: None,
        selection_fee_anchor: None,
        rev_it_integration: item("1060000", "0.06"),
        rev_it_maintenance: zero_item(),
        rev_it_device_sales: zero_item(),
        rev_it_device_lease: zero_item(),
        rev_it_other: zero_item(),
        rev_it_cloud: zero_item(),
        rev_ct_line: zero_item(),
        rev_ct_product: zero_item(),
        rev_non_it_ct: zero_item(),
        cost_it_device: item("678000", "0.13"),
        cost_it_construction: zero_item(),
        cost_it_survey: zero_item(),
        cost_it_integration: zero_item(),
        cost_it_other: zero_item(),
        cost_it_maintenance: zero_item(),
        cost_it_running: zero_item(),
        cost_it_bidding: zero_item(),
        cost_it_design_eval: zero_item(),
        cost_it_audit: zero_item(),
        cost_ct_construction: zero_item(),
        cost_ct_maintenance: zero_item(),
        cost_ct_other: zero_item(),
        cost_ct_bandwidth: zero_item(),
        cost_ct_renewal: zero_item(),
        cost_non_it_ct: zero_item(),
        cost_mix_marketing: zero_item(),
        cost_mix_channel: zero_item(),
        cost_mix_other: zero_item(),
    }
}

/// One project carrying two schemes: a 甄选前 baseline and a cheaper 甄选后 revision.
struct Fixture {
    _db_path: std::path::PathBuf,
    service: ProjectService,
    project_id: String,
    pre_scheme_id: String,
    post_scheme_id: String,
}

fn build_fixture(name: &str) -> Fixture {
    let db_path = temp_db_path(name);
    let conn = crate::db::init_db(&db_path).expect("init_db");
    let repo = SqliteProjectRepository::new(Arc::new(std::sync::Mutex::new(conn)));

    let mut project = blank_project();
    project.id = uuid::Uuid::new_v4().to_string();
    project.name = format!("桥接测试项目-{name}");
    project.customer_name = "测试客户".to_string();
    // Schemes carry a foreign key onto `projects`, so the project row lands first.
    repo.save_project(&project).expect("save_project");

    let pre = save_scheme(&repo, &project.id, "甄选前", "pre_selection", 1, "1060000");
    let post = save_scheme(&repo, &project.id, "甄选后", "post_selection", 2, "1272000");
    project.default_scheme_id = Some(pre.clone());
    repo.save_project(&project).expect("update default scheme");

    Fixture {
        _db_path: db_path,
        service: ProjectService::new(Box::new(repo)),
        project_id: project.id.clone(),
        pre_scheme_id: pre,
        post_scheme_id: post,
    }
}

fn blank_project() -> crate::benefit::models::Project {
    crate::benefit::models::Project {
        id: String::new(),
        name: String::new(),
        customer_name: String::new(),
        project_type: "ict".to_string(),
        status: "需求导入".to_string(),
        benefit_status: "normal".to_string(),
        default_scheme_id: None,
        created_at: "2026-01-01T00:00:00Z".to_string(),
        updated_at: "2026-01-01T00:00:00Z".to_string(),
        total_revenue_incl: 0.0,
        total_cost_incl: 0.0,
        project_years: 1,
        discount_rate: 0.055,
        cashflow_model: "A".to_string(),
        summary_metrics: None,
        folder_path: None,
        main_document_path: None,
        main_budget_file_path: None,
        note: None,
        logs: vec![],
        folder_name: None,
        relative_path: None,
        progress: 0.0,
        deadline: None,
        linked_folder_type: None,
        linked_folder_relative_path: None,
        linked_folder_external_path: None,
    }
}

fn save_scheme(
    repo: &SqliteProjectRepository,
    project_id: &str,
    name: &str,
    stage: &str,
    order: i32,
    revenue_incl: &str,
) -> String {
    let scheme_id = uuid::Uuid::new_v4().to_string();
    let stamp = format!("2026-01-0{order}T00:00:00Z");
    repo.save_scheme(&BenefitAnalysisScheme {
        id: scheme_id.clone(),
        project_id: project_id.to_string(),
        name: name.to_string(),
        stage: Some(stage.to_string()),
        created_at: stamp.clone(),
        updated_at: stamp.clone(),
    })
    .expect("save_scheme");

    let mut input = sample_input(name);
    input.rev_it_integration = item(revenue_incl, "0.06");
    let output = crate::benefit::calculator::calculate_ict_benefit(input.clone()).expect("calc");
    repo.save_snapshot(&BenefitAnalysisSnapshot {
        id: uuid::Uuid::new_v4().to_string(),
        scheme_id: scheme_id.clone(),
        project_id: project_id.to_string(),
        version: 1,
        input_params: input,
        output_metrics: output,
        fingerprint: "test".to_string(),
        created_at: stamp,
    })
    .expect("save_snapshot");
    scheme_id
}

/// Ground truth: what the engine produces for the same inputs, computed directly.
fn expected(revenue_incl: &str) -> IctResult {
    let mut input = sample_input("expected");
    input.rev_it_integration = item(revenue_incl, "0.06");
    crate::benefit::calculator::calculate_ict_benefit(input).expect("calc")
}

// ------------------------------------------------ route-level (always runs) --

#[test]
fn calculate_route_defaults_to_the_project_default_scheme() {
    let fx = build_fixture("default-scheme");
    let response = run_calculation(
        &fx.service,
        &CalculateRequest {
            project_id: fx.project_id.clone(),
            scenario: None,
        },
    )
    .expect("calculation succeeds");

    assert_eq!(response.scheme_id, fx.pre_scheme_id);
    assert_eq!(response.stage, "pre_selection");
    assert_eq!(response.metrics.npv, expected("1060000").npv);
    assert_eq!(response.cashflow.len(), expected("1060000").cashflow.len());
}

#[test]
fn calculate_route_selects_a_scheme_by_stage() {
    let fx = build_fixture("by-stage");
    let response = run_calculation(
        &fx.service,
        &CalculateRequest {
            project_id: fx.project_id.clone(),
            scenario: Some("post_selection".to_string()),
        },
    )
    .expect("calculation succeeds");

    assert_eq!(response.scheme_id, fx.post_scheme_id);
    assert_eq!(response.metrics.npv, expected("1272000").npv);
    // The two schemes must not collapse onto the same numbers.
    assert_ne!(response.metrics.npv, expected("1060000").npv);
}

#[test]
fn calculate_route_selects_a_scheme_by_name_and_id() {
    let fx = build_fixture("by-name");
    let by_name = run_calculation(
        &fx.service,
        &CalculateRequest {
            project_id: fx.project_id.clone(),
            scenario: Some("甄选后".to_string()),
        },
    )
    .expect("by name");
    let by_id = run_calculation(
        &fx.service,
        &CalculateRequest {
            project_id: fx.project_id.clone(),
            scenario: Some(fx.post_scheme_id.clone()),
        },
    )
    .expect("by id");

    assert_eq!(by_name.scheme_id, fx.post_scheme_id);
    assert_eq!(by_id.scheme_id, fx.post_scheme_id);
}

#[test]
fn calculate_route_rejects_an_unknown_scenario_instead_of_guessing() {
    let fx = build_fixture("unknown-scenario");
    let error = run_calculation(
        &fx.service,
        &CalculateRequest {
            project_id: fx.project_id.clone(),
            scenario: Some("不存在的方案".to_string()),
        },
    )
    .expect_err("unknown scenario must fail");
    assert!(error.contains("没有匹配"), "unexpected error: {error}");
}

#[test]
fn bridge_server_requires_its_token_and_serves_json() {
    let hits = Arc::new(AtomicUsize::new(0));
    let counter = Arc::clone(&hits);
    let server = BridgeServer::start(Arc::new(move |path, body| {
        counter.fetch_add(1, Ordering::SeqCst);
        BridgeReply::ok(serde_json::json!({ "path": path, "body": body }).to_string())
    }))
    .expect("bridge starts");

    let unauthorized = http_post(&server.origin(), CALCULATE_ROUTE, "wrong-token", "{}");
    assert_eq!(unauthorized.0, 401, "body: {}", unauthorized.1);
    assert_eq!(hits.load(Ordering::SeqCst), 0, "handler ran without a token");

    let (status, body) = http_post(&server.origin(), CALCULATE_ROUTE, server.token(), "{\"a\":1}");
    assert_eq!(status, 200, "body: {body}");
    let parsed: Value = serde_json::from_str(&body).expect("json body");
    assert_eq!(parsed["path"], CALCULATE_ROUTE);
    assert_eq!(parsed["body"], "{\"a\":1}");
    assert_eq!(hits.load(Ordering::SeqCst), 1);
}

#[test]
fn bridge_server_end_to_end_returns_engine_numbers() {
    let fx = build_fixture("end-to-end");
    let service = Arc::new(fx.service);
    let server = BridgeServer::start(Arc::new(move |path, body| {
        assert_eq!(path, CALCULATE_ROUTE);
        let request: CalculateRequest = serde_json::from_str(body).expect("request");
        match run_calculation(&service, &request) {
            Ok(response) => BridgeReply::ok(serde_json::to_string(&response).unwrap()),
            Err(e) => BridgeReply::error(422, &e),
        }
    }))
    .expect("bridge starts");

    let payload = serde_json::json!({
        "projectId": fx.project_id,
        "scenario": "post_selection",
    })
    .to_string();
    let (status, body) = http_post(&server.origin(), CALCULATE_ROUTE, server.token(), &payload);
    assert_eq!(status, 200, "body: {body}");

    let parsed: Value = serde_json::from_str(&body).expect("json body");
    assert_eq!(parsed["schemeId"], fx.post_scheme_id);
    assert_eq!(parsed["metrics"]["npv"], expected("1272000").npv);
    assert_eq!(parsed["metrics"]["marginRate"], expected("1272000").margin_rate);
}

/// Minimal blocking HTTP/1.1 client, so the tests exercise the real socket path
/// without pulling a client crate into the build.
fn http_post(origin: &str, path: &str, token: &str, body: &str) -> (u16, String) {
    use std::io::{Read, Write};
    let addr = origin.trim_start_matches("http://");
    let mut stream = std::net::TcpStream::connect(addr).expect("connect to bridge");
    let request = format!(
        "POST {path} HTTP/1.1\r\nHost: {addr}\r\n{BRIDGE_TOKEN_HEADER}: {token}\r\nContent-Type: application/json\r\nContent-Length: {}\r\nConnection: close\r\n\r\n{body}",
        body.len()
    );
    stream.write_all(request.as_bytes()).expect("write request");
    let mut raw = String::new();
    stream.read_to_string(&mut raw).expect("read response");

    let (head, payload) = raw.split_once("\r\n\r\n").expect("http response");
    let status = head
        .lines()
        .next()
        .and_then(|line| line.split_whitespace().nth(1))
        .and_then(|code| code.parse().ok())
        .expect("status code");
    (status, payload.to_string())
}

// ------------------------------------------- full loop through dsh (ignored) --

/// Collects notifications off the dsh reader thread so a test can wait on them.
#[derive(Default)]
struct EventLog {
    state: Mutex<Vec<(String, Value)>>,
    signal: Condvar,
}

impl EventLog {
    fn record(&self, method: &str, params: &Value) {
        if let Ok(mut events) = self.state.lock() {
            events.push((method.to_string(), params.clone()));
            self.signal.notify_all();
        }
    }

    /// Block until `predicate` matches a recorded event, or the deadline passes.
    fn wait_for<T>(
        &self,
        timeout: Duration,
        label: &str,
        predicate: impl Fn(&str, &Value) -> Option<T>,
    ) -> T {
        let deadline = Instant::now() + timeout;
        let mut cursor = 0usize;
        let mut events = self.state.lock().expect("event log lock");
        loop {
            while cursor < events.len() {
                let (method, params) = &events[cursor];
                cursor += 1;
                if let Some(found) = predicate(method, params) {
                    return found;
                }
            }
            let remaining = deadline.saturating_duration_since(Instant::now());
            assert!(!remaining.is_zero(), "等待 {label} 超时");
            let (next, _) = self
                .signal
                .wait_timeout(events, remaining)
                .expect("event log lock");
            events = next;
        }
    }
}

/// Read one session-event payload, if the notification is one.
fn session_event<'a>(method: &str, params: &'a Value) -> Option<(&'a str, &'a Value)> {
    if method != "session.event" {
        return None;
    }
    let event = params.get("event")?;
    Some((event.get("type")?.as_str()?, event.get("data")?))
}

fn repo_root_for_tests() -> std::path::PathBuf {
    // CARGO_MANIFEST_DIR is `<repo>/src-tauri`.
    std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .expect("repo root")
        .to_path_buf()
}

/// Launch the bridge over `fixture`'s database plus a dsh child wired to it.
fn launch_dsh(
    fixture: &Fixture,
    hits: Arc<Mutex<Vec<Value>>>,
    log: Arc<EventLog>,
) -> (BridgeServer, DshSession, DshLaunchConfig) {
    let service = Arc::new(ProjectService::new(Box::new(SqliteProjectRepository::new(
        Arc::new(std::sync::Mutex::new(
            crate::db::init_db(&fixture._db_path).expect("reopen db"),
        )),
    ))));
    let bridge = BridgeServer::start(Arc::new(move |path, body| {
        if path != CALCULATE_ROUTE {
            return BridgeReply::error(404, path);
        }
        if let Ok(value) = serde_json::from_str::<Value>(body) {
            hits.lock().expect("hits lock").push(value);
        }
        let request: CalculateRequest = match serde_json::from_str(body) {
            Ok(request) => request,
            Err(e) => return BridgeReply::error(400, &e.to_string()),
        };
        match run_calculation(&service, &request) {
            Ok(response) => BridgeReply::ok(serde_json::to_string(&response).unwrap()),
            Err(e) => BridgeReply::error(422, &e),
        }
    }))
    .expect("bridge starts");

    let mut config = DshLaunchConfig::from_repo_root(&repo_root_for_tests());
    config.bridge_url = bridge.origin();
    config.bridge_token = bridge.token().to_string();

    let sink_log = Arc::clone(&log);
    let session = DshSession::spawn(
        &config,
        Arc::new(move |method, params| sink_log.record(method, params)),
    )
    .expect("dsh spawns");

    (bridge, session, config)
}

/// Checkpoint that needs no API key and no dsh boot: run the plugin's own tool
/// body against a live bridge. It isolates the plugin↔Rust transport contract
/// (URL env var, token header, request/response JSON) from the agent loop, so a
/// failure here is unambiguously a bridge fault rather than a model fault.
#[test]
#[ignore = "needs agent-bridge provisioned: npm install && npm run provision"]
fn plugin_tool_body_reaches_the_calculator_over_the_bridge() {
    let fx = build_fixture("plugin-body");
    let service = Arc::new(fx.service);
    let server = BridgeServer::start(Arc::new(move |path, body| {
        if path != CALCULATE_ROUTE {
            return BridgeReply::error(404, path);
        }
        let request: CalculateRequest = match serde_json::from_str(body) {
            Ok(request) => request,
            Err(e) => return BridgeReply::error(400, &e.to_string()),
        };
        match run_calculation(&service, &request) {
            Ok(response) => BridgeReply::ok(serde_json::to_string(&response).unwrap()),
            Err(e) => BridgeReply::error(422, &e),
        }
    }))
    .expect("bridge starts");

    let script = repo_root_for_tests().join("agent-bridge/scripts/check-bridge.mjs");
    let output = std::process::Command::new("node")
        .arg(&script)
        .arg(&fx.project_id)
        .arg("post_selection")
        .env("LAMBER_BRIDGE_URL", server.origin())
        .env("LAMBER_BRIDGE_TOKEN", server.token())
        .env("LAMBER_BRIDGE_TOKEN_HEADER", BRIDGE_TOKEN_HEADER)
        .output()
        .expect("node runs check-bridge.mjs");
    assert!(
        output.status.success(),
        "check-bridge.mjs 失败: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let value: Value =
        serde_json::from_slice(&output.stdout).expect("tool value is json");
    assert_eq!(value["schemeId"], fx.post_scheme_id);
    assert_eq!(value["stage"], "post_selection");
    assert_eq!(value["metrics"]["npv"], expected("1272000").npv);
    assert_eq!(
        value["metrics"]["marginRate"],
        expected("1272000").margin_rate
    );
    assert!(
        !value["cashflow"].as_array().expect("cashflow rows").is_empty(),
        "现金流不应为空"
    );
}

/// Checkpoint that needs no API key: dsh boots, the patch loads `dsh-tool-lamber`,
/// and `run_benefit_calculation` reaches the model's tool catalog.
#[test]
#[ignore = "needs agent-bridge provisioned: npm install && npm run provision"]
fn dsh_advertises_the_lamber_tool_in_its_request_header() {
    let fx = build_fixture("dsh-header");
    let log = Arc::new(EventLog::default());
    let (_bridge, session, config) =
        launch_dsh(&fx, Arc::new(Mutex::new(Vec::new())), Arc::clone(&log));

    session.initialize(&config).expect("initialize");
    session
        .prompt(
            &format!("lamber-test-{}", uuid::Uuid::new_v4().simple()),
            &format!("帮我算一下项目 {} 的效益", fx.project_id),
        )
        .expect("prompt accepted");

    let tools = log.wait_for(Duration::from_secs(60), "request/header", |method, params| {
        let (kind, data) = session_event(method, params)?;
        if kind != "request/header" {
            return None;
        }
        Some(
            data.get("header")?
                .get("tools")?
                .as_array()?
                .iter()
                .filter_map(|t| t.get("name")?.as_str().map(str::to_string))
                .collect::<Vec<_>>(),
        )
    });

    assert!(
        tools.iter().any(|name| name == "run_benefit_calculation"),
        "run_benefit_calculation 未出现在工具目录中: {tools:?}"
    );
}

/// The full闭环: the model calls the tool, the bridge runs the real engine, and
/// the engine's numbers come back in `tool/result`.
#[test]
#[ignore = "needs DEEPSEEK_API_KEY plus a provisioned agent-bridge"]
fn dsh_tool_call_reaches_the_calculator_and_returns_real_numbers() {
    if std::env::var("DEEPSEEK_API_KEY")
        .ok()
        .filter(|k| !k.is_empty())
        .is_none()
    {
        eprintln!("跳过：未设置 DEEPSEEK_API_KEY");
        return;
    }

    let fx = build_fixture("dsh-tool-call");
    let log = Arc::new(EventLog::default());
    let hits = Arc::new(Mutex::new(Vec::new()));
    // `DshLaunchConfig::from_repo_root` reads DEEPSEEK_API_KEY, and the child
    // inherits it at spawn time — assigning it after the spawn would be a no-op.
    let (_bridge, session, config) = launch_dsh(&fx, Arc::clone(&hits), Arc::clone(&log));
    assert!(config.api_key.is_some(), "dsh 启动配置未带上 API key");

    session.initialize(&config).expect("initialize");
    session
        .prompt(
            &format!("lamber-test-{}", uuid::Uuid::new_v4().simple()),
            &format!(
                "请调用 run_benefit_calculation 工具，projectId 用 {}，scenario 用 post_selection，\
                 然后把 NPV 和利润率告诉我。",
                fx.project_id
            ),
        )
        .expect("prompt accepted");

    let call = log.wait_for(Duration::from_secs(120), "tool/call", |method, params| {
        let (kind, data) = session_event(method, params)?;
        if kind != "tool/call" || data.get("name")?.as_str()? != "run_benefit_calculation" {
            return None;
        }
        Some(data.get("arguments")?.as_str()?.to_string())
    });
    let args: Value = serde_json::from_str(&call).expect("tool arguments are json");
    assert_eq!(args["projectId"], fx.project_id);

    let result_text = log.wait_for(Duration::from_secs(120), "tool/result", |method, params| {
        let (kind, data) = session_event(method, params)?;
        if kind != "tool/result" {
            return None;
        }
        assert!(data.get("error").is_none(), "工具调用失败: {data}");
        Some(data.get("message")?.to_string())
    });

    let bridge_hits = hits.lock().expect("hits lock");
    assert_eq!(bridge_hits.len(), 1, "桥接服务应恰好被调用一次");
    assert_eq!(bridge_hits[0]["projectId"], fx.project_id);

    let npv = expected("1272000").npv;
    assert!(
        result_text.contains(&npv),
        "tool/result 中未出现引擎算出的 NPV {npv}: {result_text}"
    );
}
