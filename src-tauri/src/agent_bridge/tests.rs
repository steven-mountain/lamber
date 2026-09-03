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

use super::approval::{
    handle_request as handle_approval, ApprovalGate, ApprovalPrompt, ApprovalRecord,
    ApprovalRequest, APPROVAL_ROUTE,
};
use super::approval_log;
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

// ------------------------------------------- approval gate (always runs) --

/// Serve the approval route over a gate, so tests exercise the real socket path.
fn approval_bridge(
    gate: Arc<ApprovalGate>,
    prompts: Arc<Mutex<Vec<ApprovalPrompt>>>,
) -> BridgeServer {
    BridgeServer::start(Arc::new(move |path, body| {
        if path != APPROVAL_ROUTE {
            return BridgeReply::error(404, path);
        }
        let request: ApprovalRequest = match serde_json::from_str(body) {
            Ok(request) => request,
            Err(e) => return BridgeReply::error(400, &e.to_string()),
        };
        let prompts = Arc::clone(&prompts);
        let decision = handle_approval(&gate, request, |prompt| {
            prompts.lock().expect("prompts lock").push(prompt.clone());
        });
        BridgeReply::ok(serde_json::to_string(&decision).unwrap())
    }))
    .expect("bridge starts")
}

fn approval_payload(tool: &str) -> String {
    serde_json::json!({
        "toolName": tool,
        "callId": "call-1",
        "reason": "该工具会写入文件，需要你确认后才执行。",
        "args": { "note": "联调" },
    })
    .to_string()
}

/// No answer must produce an explicit denial, never a hang and never a grant.
#[test]
fn approval_times_out_as_an_explicit_rejection() {
    // This gate's own short wait; the production default is 90s.
    let gate = Arc::new(ApprovalGate::new(Duration::from_secs(1)));
    let prompts = Arc::new(Mutex::new(Vec::new()));
    let server = approval_bridge(Arc::clone(&gate), Arc::clone(&prompts));

    let started = Instant::now();
    let (status, body) = http_post(
        &server.origin(),
        APPROVAL_ROUTE,
        server.token(),
        &approval_payload("write_test_marker"),
    );
    let elapsed = started.elapsed();

    assert_eq!(status, 200, "body: {body}");
    let decision: Value = serde_json::from_str(&body).expect("json body");
    assert_eq!(decision["approved"], false, "超时必须判定为拒绝");
    assert!(
        decision["reason"].as_str().unwrap_or_default().contains("超时"),
        "拒绝原因应说明是超时: {decision}"
    );
    assert!(
        elapsed < Duration::from_secs(20),
        "超时未生效，请求挂起了 {elapsed:?}"
    );
    // The parked slot must be released, not leaked.
    assert_eq!(gate.pending_count(), 0);
    assert_eq!(prompts.lock().expect("prompts lock").len(), 1);
}

/// A simulated frontend confirmation must wake the parked request and grant it.
#[test]
fn approval_confirmation_wakes_the_parked_request() {
    let gate = Arc::new(ApprovalGate::default());
    let prompts = Arc::new(Mutex::new(Vec::new()));
    let server = approval_bridge(Arc::clone(&gate), Arc::clone(&prompts));

    // Stand in for the frontend: poll for the prompt, then answer it.
    let responder_gate = Arc::clone(&gate);
    let responder_prompts = Arc::clone(&prompts);
    let responder = std::thread::spawn(move || {
        let deadline = Instant::now() + Duration::from_secs(10);
        loop {
            let id = responder_prompts
                .lock()
                .expect("prompts lock")
                .first()
                .map(|p| p.request_id.clone());
            if let Some(id) = id {
                return responder_gate.resolve(&id, true);
            }
            assert!(Instant::now() < deadline, "等待审批提示超时");
            std::thread::sleep(Duration::from_millis(20));
        }
    });

    let (status, body) = http_post(
        &server.origin(),
        APPROVAL_ROUTE,
        server.token(),
        &approval_payload("write_test_marker"),
    );
    responder.join().expect("responder thread").expect("resolve succeeds");

    assert_eq!(status, 200, "body: {body}");
    let decision: Value = serde_json::from_str(&body).expect("json body");
    assert_eq!(decision["approved"], true);
    assert_eq!(decision["reason"], "用户已确认");

    // The prompt handed to the frontend must carry what the dialog needs to show.
    let prompt = prompts.lock().expect("prompts lock")[0].clone();
    assert_eq!(prompt.tool_name, "write_test_marker");
    assert_eq!(prompt.args["note"], "联调");
    assert!(prompt.reason.unwrap_or_default().contains("写入文件"));
    assert_eq!(gate.pending_count(), 0);
}

/// A rejection must come back as a denial, and must not be mistaken for a timeout.
#[test]
fn approval_rejection_is_reported_as_denied() {
    let gate = Arc::new(ApprovalGate::default());
    let prompts = Arc::new(Mutex::new(Vec::new()));
    let server = approval_bridge(Arc::clone(&gate), Arc::clone(&prompts));

    let responder_gate = Arc::clone(&gate);
    let responder_prompts = Arc::clone(&prompts);
    std::thread::spawn(move || loop {
        let id = responder_prompts
            .lock()
            .expect("prompts lock")
            .first()
            .map(|p| p.request_id.clone());
        if let Some(id) = id {
            let _ = responder_gate.resolve(&id, false);
            return;
        }
        std::thread::sleep(Duration::from_millis(20));
    });

    let (_, body) = http_post(
        &server.origin(),
        APPROVAL_ROUTE,
        server.token(),
        &approval_payload("write_test_marker"),
    );
    let decision: Value = serde_json::from_str(&body).expect("json body");
    assert_eq!(decision["approved"], false);
    assert_eq!(decision["reason"], "用户已拒绝");
}

/// Answering an unknown or already-settled request must fail loudly.
#[test]
fn resolving_an_unknown_approval_request_errors() {
    let gate = ApprovalGate::default();
    let error = gate
        .resolve("no-such-request", true)
        .expect_err("unknown request must fail");
    assert!(error.contains("不存在"), "unexpected error: {error}");
}

/// The approval route must not weaken the bridge's token check.
#[test]
fn approval_route_still_requires_the_bridge_token() {
    let gate = Arc::new(ApprovalGate::default());
    let server = approval_bridge(Arc::clone(&gate), Arc::new(Mutex::new(Vec::new())));
    let (status, _) = http_post(
        &server.origin(),
        APPROVAL_ROUTE,
        "wrong-token",
        &approval_payload("write_test_marker"),
    );
    assert_eq!(status, 401);
    assert_eq!(gate.pending_count(), 0, "未鉴权的请求不应创建审批槽位");
}

/// A parked approval must not stall an unrelated route on the same bridge.
#[test]
fn a_parked_approval_does_not_block_the_calculation_route() {
    let fx = build_fixture("parked-approval");
    let service = Arc::new(fx.service);
    let gate = Arc::new(ApprovalGate::new(Duration::from_secs(5)));
    let handler_gate = Arc::clone(&gate);
    let server = BridgeServer::start(Arc::new(move |path, body| {
        if path == APPROVAL_ROUTE {
            let request: ApprovalRequest = serde_json::from_str(body).expect("request");
            let decision = handle_approval(&handler_gate, request, |_| {});
            return BridgeReply::ok(serde_json::to_string(&decision).unwrap());
        }
        let request: CalculateRequest = serde_json::from_str(body).expect("request");
        match run_calculation(&service, &request) {
            Ok(response) => BridgeReply::ok(serde_json::to_string(&response).unwrap()),
            Err(e) => BridgeReply::error(422, &e),
        }
    }))
    .expect("bridge starts");

    let origin = server.origin();
    let token = server.token().to_string();
    let parked = {
        let origin = origin.clone();
        let token = token.clone();
        std::thread::spawn(move || {
            http_post(&origin, APPROVAL_ROUTE, &token, &approval_payload("write_test_marker"))
        })
    };

    // While that request is parked, the calculation route must answer promptly.
    let started = Instant::now();
    let payload = serde_json::json!({ "projectId": fx.project_id }).to_string();
    let (status, body) = http_post(&origin, CALCULATE_ROUTE, &token, &payload);
    assert_eq!(status, 200, "body: {body}");
    assert!(
        started.elapsed() < Duration::from_secs(3),
        "计算路由被挂起的审批请求阻塞了 {:?}",
        started.elapsed()
    );

    let (approval_status, _) = parked.join().expect("parked thread");
    assert_eq!(approval_status, 200);
}

// ------------------------------- audit persistence & shutdown (always runs) --

/// Open a throwaway workspace database for audit-log tests.
fn audit_db(name: &str) -> (std::path::PathBuf, Arc<std::sync::Mutex<rusqlite::Connection>>) {
    let path = temp_db_path(name);
    let conn = crate::db::init_db(&path).expect("init_db");
    (path, Arc::new(std::sync::Mutex::new(conn)))
}

/// Wire a gate to persist into `conn`.
fn recording_gate(
    timeout: Duration,
    conn: Arc<std::sync::Mutex<rusqlite::Connection>>,
) -> Arc<ApprovalGate> {
    let gate = Arc::new(ApprovalGate::new(timeout));
    gate.set_recorder(Arc::new(move |record: &ApprovalRecord| {
        approval_log::insert(&conn, record).expect("audit insert");
    }));
    gate
}

/// A decision must outlive the process that made it: reopen the database from
/// disk with a fresh connection and the entry is still there.
#[test]
fn approval_decisions_survive_a_process_restart() {
    let (path, conn) = audit_db("audit-restart");
    let gate = recording_gate(Duration::from_secs(20), Arc::clone(&conn));
    let prompts = Arc::new(Mutex::new(Vec::new()));
    let server = approval_bridge(Arc::clone(&gate), Arc::clone(&prompts));

    let responder_gate = Arc::clone(&gate);
    let responder_prompts = Arc::clone(&prompts);
    std::thread::spawn(move || loop {
        let id = responder_prompts
            .lock()
            .expect("prompts lock")
            .first()
            .map(|p| p.request_id.clone());
        if let Some(id) = id {
            let _ = responder_gate.resolve(&id, true);
            return;
        }
        std::thread::sleep(Duration::from_millis(20));
    });

    let (status, _) = http_post(
        &server.origin(),
        APPROVAL_ROUTE,
        server.token(),
        &approval_payload("write_test_marker"),
    );
    assert_eq!(status, 200);

    // Drop everything that held the decision in memory.
    drop(server);
    drop(gate);
    drop(conn);

    // Reopen from disk, exactly as a restarted process would.
    let reopened = std::sync::Mutex::new(crate::db::init_db(&path).expect("reopen db"));
    let entries = approval_log::recent(&reopened, 10).expect("read audit log");
    assert_eq!(entries.len(), 1, "重启后应仍能查到审批记录");
    let entry = &entries[0];
    assert_eq!(entry.tool_name, "write_test_marker");
    assert!(entry.approved, "应记录为已批准");
    assert_eq!(entry.decided_by, "user");
    assert_eq!(entry.decision_reason, "用户已确认");
    assert_eq!(entry.call_id.as_deref(), Some("call-1"));
    assert!(entry.args_json.contains("联调"), "应保留当时展示给用户的参数");
    assert!(!entry.requested_at.is_empty() && !entry.decided_at.is_empty());
}

/// Rejections and timeouts must be auditable too, and distinguishable.
#[test]
fn audit_log_distinguishes_rejection_from_timeout() {
    let (_path, conn) = audit_db("audit-kinds");

    // A refusal by a human.
    let gate = recording_gate(Duration::from_secs(20), Arc::clone(&conn));
    let prompts = Arc::new(Mutex::new(Vec::new()));
    let server = approval_bridge(Arc::clone(&gate), Arc::clone(&prompts));
    let responder_gate = Arc::clone(&gate);
    let responder_prompts = Arc::clone(&prompts);
    std::thread::spawn(move || loop {
        let id = responder_prompts
            .lock()
            .expect("prompts lock")
            .first()
            .map(|p| p.request_id.clone());
        if let Some(id) = id {
            let _ = responder_gate.resolve(&id, false);
            return;
        }
        std::thread::sleep(Duration::from_millis(20));
    });
    http_post(
        &server.origin(),
        APPROVAL_ROUTE,
        server.token(),
        &approval_payload("write_test_marker"),
    );

    // Nobody answering at all.
    let silent_gate = recording_gate(Duration::from_secs(1), Arc::clone(&conn));
    let silent = approval_bridge(Arc::clone(&silent_gate), Arc::new(Mutex::new(Vec::new())));
    http_post(
        &silent.origin(),
        APPROVAL_ROUTE,
        silent.token(),
        &approval_payload("write_test_marker"),
    );

    let entries = approval_log::recent(&conn, 10).expect("read audit log");
    assert_eq!(entries.len(), 2);
    let by: Vec<&str> = entries.iter().map(|e| e.decided_by.as_str()).collect();
    assert!(by.contains(&"user"), "应记录到人工拒绝: {by:?}");
    assert!(by.contains(&"timeout"), "应记录到超时拒绝: {by:?}");
    assert!(
        entries.iter().all(|e| !e.approved),
        "拒绝与超时都不应记为批准"
    );
}

/// Shutting the runtime down must release parked questions at once, deny them,
/// and record why — not leave a worker (and the answerer) waiting for the timeout.
#[test]
fn shutdown_denies_parked_approvals_immediately() {
    let (_path, conn) = audit_db("audit-shutdown");
    // A long wait, so only the shutdown can end this request quickly.
    let gate = recording_gate(Duration::from_secs(300), Arc::clone(&conn));
    let prompts = Arc::new(Mutex::new(Vec::new()));
    let server = approval_bridge(Arc::clone(&gate), Arc::clone(&prompts));

    let origin = server.origin();
    let token = server.token().to_string();
    let parked = std::thread::spawn(move || {
        http_post(&origin, APPROVAL_ROUTE, &token, &approval_payload("write_test_marker"))
    });

    // Wait until the question is actually parked, then tear the gate down.
    let deadline = Instant::now() + Duration::from_secs(10);
    while gate.pending_count() == 0 {
        assert!(Instant::now() < deadline, "审批请求未进入挂起状态");
        std::thread::sleep(Duration::from_millis(20));
    }
    let started = Instant::now();
    let denied = gate.shutdown();
    assert_eq!(denied, 1, "应拒绝掉 1 个未完成的审批");

    let (status, body) = parked.join().expect("parked thread");
    assert_eq!(status, 200, "body: {body}");
    assert!(
        started.elapsed() < Duration::from_secs(10),
        "关闭后请求未及时释放，用时 {:?}",
        started.elapsed()
    );
    let decision: Value = serde_json::from_str(&body).expect("json body");
    assert_eq!(decision["approved"], false);
    assert!(
        decision["reason"].as_str().unwrap_or_default().contains("关闭"),
        "拒绝原因应说明是关闭: {decision}"
    );

    let entries = approval_log::recent(&conn, 10).expect("read audit log");
    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].decided_by, "shutdown");
    assert!(!entries[0].approved);
}

/// After a shutdown, a question arriving before the gate reopens must be denied
/// on the spot rather than parking against a runtime that no longer exists.
#[test]
fn approvals_after_shutdown_are_denied_without_parking() {
    let (_path, conn) = audit_db("audit-after-shutdown");
    let gate = recording_gate(Duration::from_secs(300), Arc::clone(&conn));
    gate.shutdown();

    let server = approval_bridge(Arc::clone(&gate), Arc::new(Mutex::new(Vec::new())));
    let started = Instant::now();
    let (status, body) = http_post(
        &server.origin(),
        APPROVAL_ROUTE,
        server.token(),
        &approval_payload("write_test_marker"),
    );

    assert_eq!(status, 200, "body: {body}");
    assert!(
        started.elapsed() < Duration::from_secs(5),
        "关闭后的请求不应挂起，用时 {:?}",
        started.elapsed()
    );
    let decision: Value = serde_json::from_str(&body).expect("json body");
    assert_eq!(decision["approved"], false);
    assert_eq!(gate.pending_count(), 0, "关闭后不应再占用槽位");

    let entries = approval_log::recent(&conn, 10).expect("read audit log");
    assert_eq!(entries[0].decided_by, "shutdown");

    // Reopening restores normal service for the next launch.
    gate.reopen();
    assert_eq!(gate.pending_count(), 0);
}

/// A recorder that fails must not change the decision the user made.
#[test]
fn an_audit_write_failure_does_not_alter_the_decision() {
    let gate = Arc::new(ApprovalGate::new(Duration::from_secs(20)));
    let attempts = Arc::new(AtomicUsize::new(0));
    let counter = Arc::clone(&attempts);
    gate.set_recorder(Arc::new(move |_record: &ApprovalRecord| {
        counter.fetch_add(1, Ordering::SeqCst);
        // Stands in for a closed workspace: the real recorder logs and swallows.
    }));

    let prompts = Arc::new(Mutex::new(Vec::new()));
    let server = approval_bridge(Arc::clone(&gate), Arc::clone(&prompts));
    let responder_gate = Arc::clone(&gate);
    let responder_prompts = Arc::clone(&prompts);
    std::thread::spawn(move || loop {
        let id = responder_prompts
            .lock()
            .expect("prompts lock")
            .first()
            .map(|p| p.request_id.clone());
        if let Some(id) = id {
            let _ = responder_gate.resolve(&id, true);
            return;
        }
        std::thread::sleep(Duration::from_millis(20));
    });

    let (_, body) = http_post(
        &server.origin(),
        APPROVAL_ROUTE,
        server.token(),
        &approval_payload("write_test_marker"),
    );
    let decision: Value = serde_json::from_str(&body).expect("json body");
    assert_eq!(decision["approved"], true, "审计失败不应影响用户的决定");
    assert_eq!(attempts.load(Ordering::SeqCst), 1);
}

/// The decision path the task pins down: an approval settled while no workspace
/// is open must not vanish. It spools, and the next workspace activation
/// backfills it into `agent_approval_log`.
#[test]
fn an_approval_taken_with_no_workspace_is_backfilled_when_one_opens() {
    let spool = std::env::temp_dir().join(format!(
        "lamber-approval-spool-{}.jsonl",
        uuid::Uuid::new_v4().simple()
    ));

    // A runtime with no workspace: `require_db` fails, so the recorder spools.
    let runtime = Arc::new(crate::workspace::WorkspaceRuntime::new());
    assert!(runtime.require_db().is_err(), "前提：此时没有打开工作区");

    let gate = Arc::new(ApprovalGate::new(Duration::from_secs(20)));
    gate.set_recorder(approval_log::workspace_recorder(
        Arc::clone(&runtime),
        spool.clone(),
    ));

    let prompts = Arc::new(Mutex::new(Vec::new()));
    let server = approval_bridge(Arc::clone(&gate), Arc::clone(&prompts));
    let responder_gate = Arc::clone(&gate);
    let responder_prompts = Arc::clone(&prompts);
    std::thread::spawn(move || loop {
        let id = responder_prompts
            .lock()
            .expect("prompts lock")
            .first()
            .map(|p| p.request_id.clone());
        if let Some(id) = id {
            let _ = responder_gate.resolve(&id, true);
            return;
        }
        std::thread::sleep(Duration::from_millis(20));
    });

    let (status, body) = http_post(
        &server.origin(),
        APPROVAL_ROUTE,
        server.token(),
        &approval_payload("write_test_marker"),
    );
    assert_eq!(status, 200, "body: {body}");
    let decision: Value = serde_json::from_str(&body).expect("json body");
    assert_eq!(decision["approved"], true, "没有工作区不应影响决定本身");

    // The decision is buffered, not lost.
    assert_eq!(
        approval_log::spool_len(&spool),
        1,
        "无工作区时的审批决定应进入缓冲"
    );

    // Now a workspace opens.
    let (_db_path, conn) = audit_db("spool-backfill");
    let drained = approval_log::drain_spool(&spool, &conn).expect("drain succeeds");
    assert_eq!(drained, 1);
    assert!(!spool.exists(), "回填后缓冲文件应被清除");

    // The record is queryable exactly as a directly written one would be.
    let entries = approval_log::recent(&conn, 10).expect("read audit log");
    assert_eq!(entries.len(), 1, "回填后应能查到该审批记录");
    assert_eq!(entries[0].tool_name, "write_test_marker");
    assert!(entries[0].approved);
    assert_eq!(entries[0].decided_by, "user");
    assert!(entries[0].args_json.contains("联调"));
    assert!(!entries[0].decided_at.is_empty());
}

/// Backfilling must be replay-safe: an interrupted drain leaves the spool in
/// place, and draining the same decisions again must not duplicate rows.
#[test]
fn backfilling_the_same_decisions_twice_does_not_duplicate_rows() {
    let spool = std::env::temp_dir().join(format!(
        "lamber-approval-spool-{}.jsonl",
        uuid::Uuid::new_v4().simple()
    ));
    let record = ApprovalRecord {
        request_id: "req-replay".to_string(),
        tool_name: "write_test_marker".to_string(),
        call_id: Some("call-1".to_string()),
        reason: Some("需要确认".to_string()),
        args_json: "{\"note\":\"联调\"}".to_string(),
        approved: false,
        decided_by: super::approval::DecidedBy::Timeout,
        decision_reason: "等待用户确认超时（90 秒），按拒绝处理".to_string(),
        requested_at: "2026-01-01T00:00:00Z".to_string(),
        decided_at: "2026-01-01T00:01:30Z".to_string(),
    };
    approval_log::append_to_spool(&spool, &record).expect("spool write");
    approval_log::append_to_spool(&spool, &record).expect("spool write");
    assert_eq!(approval_log::spool_len(&spool), 2);

    let (_db_path, conn) = audit_db("spool-replay");
    approval_log::drain_spool(&spool, &conn).expect("first drain");
    // Simulate a replay of the same buffered decision after an interrupted run.
    approval_log::append_to_spool(&spool, &record).expect("spool write");
    approval_log::drain_spool(&spool, &conn).expect("second drain");

    let entries = approval_log::recent(&conn, 10).expect("read audit log");
    assert_eq!(entries.len(), 1, "同一 request_id 不应产生重复记录");
    assert_eq!(entries[0].decided_by, "timeout");
    assert!(!entries[0].approved);
}

/// A torn trailing line from a crash mid-append must not block the rest.
#[test]
fn a_corrupt_spool_line_does_not_block_backfilling_the_others() {
    let spool = std::env::temp_dir().join(format!(
        "lamber-approval-spool-{}.jsonl",
        uuid::Uuid::new_v4().simple()
    ));
    let good = ApprovalRecord {
        request_id: "req-good".to_string(),
        tool_name: "write_test_marker".to_string(),
        call_id: None,
        reason: None,
        args_json: "{}".to_string(),
        approved: true,
        decided_by: super::approval::DecidedBy::User,
        decision_reason: "用户已确认".to_string(),
        requested_at: "2026-01-01T00:00:00Z".to_string(),
        decided_at: "2026-01-01T00:00:03Z".to_string(),
    };
    approval_log::append_to_spool(&spool, &good).expect("spool write");
    // A half-written line, as a crash during append would leave.
    {
        use std::io::Write;
        let mut f = std::fs::OpenOptions::new().append(true).open(&spool).unwrap();
        write!(f, "{{\"requestId\":\"req-torn\",\"toolNa").unwrap();
    }

    let (_db_path, conn) = audit_db("spool-corrupt");
    let drained = approval_log::drain_spool(&spool, &conn).expect("drain succeeds");
    assert_eq!(drained, 1, "完整的那条应被回填");
    let entries = approval_log::recent(&conn, 10).expect("read audit log");
    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].request_id, "req-good");
}

// -------------------------------------- Rust <-> frontend contract (always) --

/// Read one frontend source file that participates in the approval contract.
fn frontend_source(relative: &str) -> String {
    let path = repo_root_for_tests().join("src-ui/src").join(relative);
    std::fs::read_to_string(&path)
        .unwrap_or_else(|e| panic!("读取 {} 失败: {e}", path.display()))
}

/// The emitted prompt and the dialog that consumes it must agree on names.
///
/// A GUI click is the only way to exercise the dialog for real, and that needs
/// macOS Accessibility permission this test environment does not have. What a
/// click would actually catch is a drifted name across the seam — an event
/// renamed in Rust, a field the dialog reads that serde no longer emits, a
/// changed command parameter. That is checkable here, deterministically.
#[test]
fn the_approval_prompt_contract_matches_the_frontend_dialog() {
    let prompt = ApprovalPrompt {
        request_id: "req-1".to_string(),
        tool_name: "write_test_marker".to_string(),
        call_id: Some("call-1".to_string()),
        reason: Some("需要确认".to_string()),
        args: serde_json::json!({ "note": "联调" }),
        timeout_seconds: 90,
    };
    let wire: Value = serde_json::to_value(&prompt).expect("serialize prompt");
    let object = wire.as_object().expect("prompt serializes to an object");

    // The exact payload the frontend receives; camelCase, no snake_case leakage.
    let mut keys: Vec<&str> = object.keys().map(String::as_str).collect();
    keys.sort_unstable();
    assert_eq!(
        keys,
        vec![
            "args",
            "callId",
            "reason",
            "requestId",
            "timeoutSeconds",
            "toolName",
        ],
        "审批事件字段变化会让前端弹窗读到 undefined"
    );

    let dialog = frontend_source("components/ai/AgentApprovalDialog.tsx");
    assert!(
        dialog.contains(&format!("\"{}\"", super::approval::APPROVAL_EVENT)),
        "弹窗订阅的事件名与后端 APPROVAL_EVENT 不一致"
    );
    for field in ["requestId", "toolName", "reason", "args", "timeoutSeconds"] {
        assert!(
            dialog.contains(field),
            "弹窗未使用审批事件字段 `{field}`"
        );
    }

    // The response command and its parameter names, as Tauri will deserialize them.
    assert!(
        dialog.contains("\"ai_resolve_approval\""),
        "弹窗未调用 ai_resolve_approval"
    );
    assert!(
        dialog.contains("requestId:") && dialog.contains("approved"),
        "弹窗未按 ai_resolve_approval(requestId, approved) 传参"
    );
}

/// The dialog must be reachable wherever the agent can be driven, or a parked
/// approval has no answerer and dies of timeout.
#[test]
fn the_approval_dialog_is_mounted_on_every_agent_reachable_route() {
    let app = frontend_source("App.tsx");
    assert!(
        app.contains("<AgentApprovalDialog />"),
        "App 未挂载审批弹窗"
    );
    // Two render paths can reach the agent: the main shell and the agent bench.
    assert!(
        app.matches("<AgentApprovalDialog />").count() >= 2,
        "审批弹窗必须同时挂在主界面和 Agent 联调台两条渲染路径上"
    );
    assert!(
        app.contains("#/agent-lab"),
        "缺少 Agent 联调台路由，审批弹窗在真实应用里无法被触发"
    );

    // Without a caller for ai_send_prompt no approval can ever be raised.
    let lab = frontend_source("components/ai/AgentLabView.tsx");
    assert!(
        lab.contains("\"ai_send_prompt\""),
        "联调台未调用 ai_send_prompt，Agent 无法在真实应用中启动"
    );
    assert!(
        lab.contains("\"ai_list_approval_log\""),
        "联调台未查询审批审计日志"
    );
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

/// How a test answers approval questions the agent raises.
#[derive(Clone, Copy, PartialEq)]
enum ApprovalStance {
    /// Never answer; the gate must time out on its own.
    Silent,
    /// Answer as a user who clicked 确认.
    Approve,
    /// Answer as a user who clicked 拒绝.
    Reject,
}

/// Launch the bridge over `fixture`'s database plus a dsh child wired to it.
fn launch_dsh(
    fixture: &Fixture,
    hits: Arc<Mutex<Vec<Value>>>,
    log: Arc<EventLog>,
) -> (BridgeServer, DshSession, DshLaunchConfig) {
    launch_dsh_with_approval(
        fixture,
        hits,
        log,
        ApprovalStance::Silent,
        Arc::new(Mutex::new(Vec::new())),
    )
}

/// Launch the bridge and dsh, answering approval questions per `stance`.
fn launch_dsh_with_approval(
    fixture: &Fixture,
    hits: Arc<Mutex<Vec<Value>>>,
    log: Arc<EventLog>,
    stance: ApprovalStance,
    prompts: Arc<Mutex<Vec<ApprovalPrompt>>>,
) -> (BridgeServer, DshSession, DshLaunchConfig) {
    let service = Arc::new(ProjectService::new(Box::new(SqliteProjectRepository::new(
        Arc::new(std::sync::Mutex::new(
            crate::db::init_db(&fixture._db_path).expect("reopen db"),
        )),
    ))));
    // Short enough that a silent stance fails closed well inside the test's own
    // wait, rather than parking for the production default.
    let gate = Arc::new(ApprovalGate::new(Duration::from_secs(20)));
    let handler_gate = Arc::clone(&gate);
    let bridge = BridgeServer::start(Arc::new(move |path, body| {
        if path == APPROVAL_ROUTE {
            let request: ApprovalRequest = match serde_json::from_str(body) {
                Ok(request) => request,
                Err(e) => return BridgeReply::error(400, &e.to_string()),
            };
            let gate = Arc::clone(&handler_gate);
            let responder_gate = Arc::clone(&handler_gate);
            let prompts = Arc::clone(&prompts);
            let decision = handle_approval(&gate, request, |prompt| {
                prompts.lock().expect("prompts lock").push(prompt.clone());
                if stance == ApprovalStance::Silent {
                    return;
                }
                // Stand in for the frontend: answer from another thread, exactly
                // as `ai_resolve_approval` would.
                let id = prompt.request_id.clone();
                let approved = stance == ApprovalStance::Approve;
                std::thread::spawn(move || {
                    let _ = responder_gate.resolve(&id, approved);
                });
            });
            return BridgeReply::ok(serde_json::to_string(&decision).unwrap());
        }
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

/// 闭环 A's read-only tool must stay ungated after the approval work landed.
///
/// The guard's policy is a pure function of the tool name, so this asserts it
/// directly rather than inferring it from an agent run. Combined with the
/// `tools/pre-execute` waterfall's terminal default of `allow`, an ungated tool
/// never reaches the approval seam at all.
#[test]
#[ignore = "needs agent-bridge provisioned: npm install && npm run provision"]
fn only_the_write_tool_is_gated_behind_approval() {
    let script = repo_root_for_tests().join("agent-bridge/scripts/check-gating.mjs");
    let output = std::process::Command::new("node")
        .arg(&script)
        .arg("run_benefit_calculation")
        .arg("write_test_marker")
        .output()
        .expect("node runs check-gating.mjs");
    assert!(
        output.status.success(),
        "check-gating.mjs 失败: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let report = String::from_utf8_lossy(&output.stdout);
    assert!(
        report.contains("run_benefit_calculation=false"),
        "只读工具不应被审批拦截: {report}"
    );
    assert!(
        report.contains("write_test_marker=true"),
        "写操作工具必须经过审批: {report}"
    );
}

/// The answerer's own path, with no dsh and no LLM: `askLamber` → bridge →
/// gate → a simulated frontend decision → the outcome back in the plugin.
#[test]
#[ignore = "needs agent-bridge provisioned: npm install && npm run provision"]
fn answerer_receives_the_users_decision_through_the_gate() {
    for (approved, expected) in [(true, "allowed-once"), (false, "rejected")] {
        let gate = Arc::new(ApprovalGate::new(Duration::from_secs(20)));
        let prompts = Arc::new(Mutex::new(Vec::new()));
        let server = approval_bridge(Arc::clone(&gate), Arc::clone(&prompts));

        let responder_gate = Arc::clone(&gate);
        let responder_prompts = Arc::clone(&prompts);
        std::thread::spawn(move || {
            let deadline = Instant::now() + Duration::from_secs(15);
            while Instant::now() < deadline {
                let id = responder_prompts
                    .lock()
                    .expect("prompts lock")
                    .first()
                    .map(|p| p.request_id.clone());
                if let Some(id) = id {
                    let _ = responder_gate.resolve(&id, approved);
                    return;
                }
                std::thread::sleep(Duration::from_millis(20));
            }
        });

        let script = repo_root_for_tests().join("agent-bridge/scripts/check-approval.mjs");
        let output = std::process::Command::new("node")
            .arg(&script)
            .arg("write_test_marker")
            .env("LAMBER_BRIDGE_URL", server.origin())
            .env("LAMBER_BRIDGE_TOKEN", server.token())
            .env("LAMBER_BRIDGE_TOKEN_HEADER", BRIDGE_TOKEN_HEADER)
            .output()
            .expect("node runs check-approval.mjs");
        assert!(
            output.status.success(),
            "check-approval.mjs 失败: {}",
            String::from_utf8_lossy(&output.stderr)
        );
        let outcome = String::from_utf8_lossy(&output.stdout).trim().to_string();
        assert_eq!(
            outcome, expected,
            "用户{}时应得到 {expected}",
            if approved { "确认" } else { "拒绝" }
        );
    }
}

/// Application crash: lamber vanishes while the answerer is waiting.
///
/// The parked HTTP request dies with the process, so the answerer's `fetch`
/// rejects. It must settle closed rather than hang until its own 180s bound —
/// otherwise a crashed lamber would freeze the agent's turn.
#[test]
#[ignore = "needs agent-bridge provisioned: npm install && npm run provision"]
fn answerer_fails_closed_when_the_bridge_dies_mid_request() {
    // A listener that accepts the request and then drops the socket, exactly as
    // a killed process would. Not the real BridgeServer: this test is about what
    // the *answerer* does when its peer disappears.
    let listener = std::net::TcpListener::bind("127.0.0.1:0").expect("bind");
    let origin = format!("http://127.0.0.1:{}", listener.local_addr().unwrap().port());
    std::thread::spawn(move || {
        if let Ok((stream, _)) = listener.accept() {
            // Read the request head so the client finishes sending, then hang up.
            let mut buf = [0u8; 1024];
            use std::io::Read;
            let mut stream = stream;
            let _ = stream.read(&mut buf);
            drop(stream);
        }
    });

    let script = repo_root_for_tests().join("agent-bridge/scripts/check-approval.mjs");
    let started = Instant::now();
    let output = std::process::Command::new("node")
        .arg(&script)
        .env("LAMBER_BRIDGE_URL", &origin)
        .env("LAMBER_BRIDGE_TOKEN", "any-token")
        .env("LAMBER_BRIDGE_TOKEN_HEADER", BRIDGE_TOKEN_HEADER)
        .output()
        .expect("node runs check-approval.mjs");

    assert!(
        output.status.success(),
        "check-approval.mjs 失败: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    assert_eq!(
        String::from_utf8_lossy(&output.stdout).trim(),
        "rejected",
        "桥接进程消失后必须失败关闭"
    );
    assert!(
        started.elapsed() < Duration::from_secs(30),
        "连接断开后答复器挂起了 {:?}",
        started.elapsed()
    );
}

/// A bridge that never answers must still let the answerer settle, closed.
#[test]
#[ignore = "needs agent-bridge provisioned: npm install && npm run provision"]
fn answerer_fails_closed_when_nobody_answers() {
    let gate = Arc::new(ApprovalGate::new(Duration::from_secs(2)));
    let server = approval_bridge(Arc::clone(&gate), Arc::new(Mutex::new(Vec::new())));

    let script = repo_root_for_tests().join("agent-bridge/scripts/check-approval.mjs");
    let started = Instant::now();
    let output = std::process::Command::new("node")
        .arg(&script)
        .env("LAMBER_BRIDGE_URL", server.origin())
        .env("LAMBER_BRIDGE_TOKEN", server.token())
        .env("LAMBER_BRIDGE_TOKEN_HEADER", BRIDGE_TOKEN_HEADER)
        .output()
        .expect("node runs check-approval.mjs");

    assert!(output.status.success(), "{}", String::from_utf8_lossy(&output.stderr));
    assert_eq!(
        String::from_utf8_lossy(&output.stdout).trim(),
        "rejected",
        "无人应答必须判定为拒绝"
    );
    assert!(started.elapsed() < Duration::from_secs(30), "答复器挂起了");
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

/// Both tools must be advertised, and the read-only one must stay ungated —
///闭环 B's guard names only `write_test_marker`, and the `tools/pre-execute`
/// waterfall's terminal default is `allow`.
#[test]
#[ignore = "needs agent-bridge provisioned: npm install && npm run provision"]
fn dsh_advertises_the_gated_tool_alongside_the_readonly_one() {
    let fx = build_fixture("dsh-gated-header");
    let log = Arc::new(EventLog::default());
    let (_bridge, session, config) =
        launch_dsh(&fx, Arc::new(Mutex::new(Vec::new())), Arc::clone(&log));

    session.initialize(&config).expect("initialize");
    session
        .prompt(
            &format!("lamber-test-{}", uuid::Uuid::new_v4().simple()),
            "你有哪些工具？",
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
        tools.iter().any(|name| name == "write_test_marker"),
        "write_test_marker 未出现在工具目录中: {tools:?}"
    );
    assert!(
        tools.iter().any(|name| name == "run_benefit_calculation"),
        "run_benefit_calculation 未出现在工具目录中: {tools:?}"
    );
}

/// The full approval闭环: the model calls the gated tool, lamber asks, a
/// simulated user confirms, and the tool actually runs.
#[test]
#[ignore = "needs DEEPSEEK_API_KEY plus a provisioned agent-bridge"]
fn dsh_gated_tool_runs_only_after_the_user_confirms() {
    if std::env::var("DEEPSEEK_API_KEY")
        .ok()
        .filter(|k| !k.is_empty())
        .is_none()
    {
        eprintln!("跳过：未设置 DEEPSEEK_API_KEY");
        return;
    }

    for stance in [ApprovalStance::Approve, ApprovalStance::Reject] {
        let fx = build_fixture("dsh-approval-loop");
        let log = Arc::new(EventLog::default());
        let prompts = Arc::new(Mutex::new(Vec::new()));
        let (_bridge, session, config) = launch_dsh_with_approval(
            &fx,
            Arc::new(Mutex::new(Vec::new())),
            Arc::clone(&log),
            stance,
            Arc::clone(&prompts),
        );

        session.initialize(&config).expect("initialize");
        session
            .prompt(
                &format!("lamber-test-{}", uuid::Uuid::new_v4().simple()),
                "请调用 write_test_marker 工具，note 参数填「闭环B联调」。",
            )
            .expect("prompt accepted");

        // dsh audits every question; `approval/decided` carries the outcome.
        let outcome = log.wait_for(
            Duration::from_secs(120),
            "approval/decided",
            |method, params| {
                let (kind, data) = session_event(method, params)?;
                if kind != "approval/decided" {
                    return None;
                }
                Some(data.get("outcome")?.as_str()?.to_string())
            },
        );

        let expected = match stance {
            ApprovalStance::Approve => "allowed-once",
            _ => "rejected",
        };
        assert_eq!(outcome, expected, "审批结果与模拟的用户操作不一致");
        assert_eq!(
            prompts.lock().expect("prompts lock").len(),
            1,
            "应恰好向用户提出一次审批"
        );

        let result = log.wait_for(Duration::from_secs(120), "tool/result", |method, params| {
            let (kind, data) = session_event(method, params)?;
            if kind != "tool/result" {
                return None;
            }
            Some(data.to_string())
        });

        match stance {
            ApprovalStance::Approve => {
                let marker = extract_marker_path(&result)
                    .unwrap_or_else(|| panic!("确认后应写出标记文件: {result}"));
                assert!(
                    std::path::Path::new(&marker).is_file(),
                    "标记文件不存在: {marker}"
                );
                assert!(
                    marker.starts_with(&std::env::temp_dir().to_string_lossy().to_string()),
                    "标记文件必须落在系统临时目录: {marker}"
                );
            }
            _ => assert!(
                result.contains("rejected") || result.contains("拒绝"),
                "拒绝后 tool/result 应说明被拒: {result}"
            ),
        }
    }
}

/// Pull the marker path out of a `tool/result` payload, if it wrote one.
///
/// The payload embeds the path in prose (`已写入测试标记文件：<path>（207 字节…`),
/// so the bounds are anchored on the filename itself rather than on whitespace:
/// the directory marker on the left, the `.txt` suffix on the right.
fn extract_marker_path(result: &str) -> Option<String> {
    const DIR_MARKER: &str = "lamber-agent-marker-";
    const SUFFIX: &str = ".txt";
    let anchor = result.find(DIR_MARKER)?;
    // Walk left over path-legal characters; the surrounding prose (Chinese text,
    // JSON quotes) is not path-legal, so it bounds the scan on its own.
    let is_path_char = |c: char| c.is_ascii_alphanumeric() || "/._-".contains(c);
    let start = result[..anchor]
        .char_indices()
        .rev()
        .take_while(|(_, c)| is_path_char(*c))
        .map(|(i, _)| i)
        .last()?;
    let end = result[anchor..].find(SUFFIX)? + anchor + SUFFIX.len();
    Some(result[start..end].to_string())
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
