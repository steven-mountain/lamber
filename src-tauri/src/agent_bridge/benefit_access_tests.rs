use super::*;
use crate::agent_bridge::{benefit_access, benefit_simulation, project_bindings::ProjectBindings};
use serde_json::json;
#[test]
fn benefit_access_all_subjects_selectors_and_simulation_match_desktop() {
    let fx = build_fixture("benefit-access");
    let runtime = scoped_workspace(&fx);
    let witness: Value =
        serde_json::from_str(include_str!("fixtures/benefit-simulation-desktop.json")).unwrap();
    let input: IctInput = serde_json::from_value(witness["cases"][0]["saved"].clone()).unwrap();
    let repo = SqliteProjectRepository::new(runtime.require_db().unwrap());
    repo.save_snapshot(&BenefitAnalysisSnapshot {
        id: "benefit-access-snapshot".into(),
        scheme_id: fx.pre_scheme_id.clone(),
        project_id: fx.project_id.clone(),
        version: 2,
        input_params: input.clone(),
        output_metrics: crate::benefit::calculator::calculate_ict_benefit(input.clone()).unwrap(),
        fingerprint: "desktop-comparison".into(),
        created_at: "2026-09-09T00:00:00Z".into(),
    })
    .unwrap();
    let bindings = Arc::new(ProjectBindings::default());
    bindings
        .register("bound", fixture_binding(&fx, Some(&fx.project_id)))
        .unwrap();
    bindings
        .register("general", fixture_binding(&fx, None))
        .unwrap();
    let cases = witness["cases"].as_array().unwrap().clone();
    let prepare: benefit_simulation::Preparer = Arc::new(move |job| {
        let overrides = serde_json::to_value(&job.overrides).unwrap();
        let found = cases
            .iter()
            .find(|c| c["overrides"] == overrides)
            .ok_or("测试没有对应桌面输入")?;
        serde_json::from_value(found["prepared"].clone()).map_err(|e| e.to_string())
    });
    let handler = benefit_simulation::handler(
        runtime.clone(),
        bindings.clone(),
        prepare,
        crate::agent_bridge::workspace_handler(runtime.clone(), bindings),
    );
    let server = BridgeServer::start(handler).unwrap();
    let post = |route: &str, body: Value| {
        http_post(&server.origin(), route, server.token(), &body.to_string())
    };
    let changes = || {
        runtime
            .require_db()
            .unwrap()
            .lock()
            .unwrap()
            .query_row("SELECT total_changes()", [], |r| r.get::<_, i64>(0))
            .unwrap()
    };
    let before = changes();
    let (status, body) = post(benefit_access::READ_ROUTE, json!({"sessionId":"bound"}));
    assert_eq!(status, 200, "{body}");
    let read: Value = serde_json::from_str(&body).unwrap();
    assert_eq!(read["basis"], "saved_snapshot");
    assert_eq!(read["snapshotVersion"], 2);
    assert_eq!(read["schemeId"], fx.pre_scheme_id);
    let subjects = read["subjects"].as_array().unwrap();
    assert_eq!(subjects.len(), 28);
    assert_eq!(
        subjects.iter().filter(|s| s["side"] == "revenue").count(),
        9
    );
    let raw = serde_json::to_value(&input).unwrap();
    for row in subjects {
        let key = row["key"].as_str().unwrap();
        for (out, field) in [
            ("inclTax", "incl_tax"),
            ("taxRate", "tax_rate"),
            ("custom_subject_name", "custom_subject_name"),
            ("billing_subject_name", "billing_subject_name"),
        ] {
            assert_eq!(row[out], raw[key][field], "{key}/{field}");
        }
        assert!(!row["name"].as_str().unwrap().is_empty());
        assert!(row["displayName"].as_str().unwrap().contains("开票名称"));
    }
    assert_eq!(read["totals"]["revenueInclTax"], "110664.00");
    assert_eq!(read["totals"]["costInclTax"], "89888.00");
    for selector in ["pre_selection", &fx.pre_scheme_id, "甄选前"] {
        let (_, body) = post(
            benefit_access::READ_ROUTE,
            json!({"sessionId":"bound","scenario":selector}),
        );
        let v: Value = serde_json::from_str(&body).unwrap();
        assert_eq!(v["schemeId"], fx.pre_scheme_id);
    }
    let (_, body) = post(
        benefit_access::READ_ROUTE,
        json!({"sessionId":"bound","scenario":"post_selection"}),
    );
    let v: Value = serde_json::from_str(&body).unwrap();
    assert_eq!(v["schemeId"], fx.post_scheme_id);
    for session in ["general", "unregistered"] {
        let (status, body) = post(benefit_access::READ_ROUTE, json!({"sessionId":session}));
        assert_ne!(status, 200);
        assert!(body.contains("会话") || body.contains("通用聊天"));
    }
    for route in [benefit_access::READ_ROUTE, benefit_simulation::ROUTE] {
        for payload in [
            json!({"sessionId":"general","overrides":[]}),
            json!({"sessionId":"bound","projectId":fx.project_id,"overrides":[]}),
            json!({"sessionId":"bound","scenario":"不存在","overrides":[]}),
        ] {
            assert_ne!(post(route, payload).0, 200);
        }
    }
    for case in witness["cases"].as_array().unwrap() {
        let (status, body) = post(
            benefit_simulation::ROUTE,
            json!({"sessionId":"bound","overrides":case["overrides"]}),
        );
        assert_eq!(status, 200, "{body}");
        let v: Value = serde_json::from_str(&body).unwrap();
        assert_eq!(v["basis"], "hypothetical");
        assert_eq!(v["result"]["basis"], "hypothetical");
        let desktop: IctInput = serde_json::from_value(case["desktop"].clone()).unwrap();
        let result = crate::benefit::calculator::calculate_ict_benefit(desktop).unwrap();
        let (p, s, snap) = crate::agent_bridge::calculation::resolve_snapshot(
            &fx.service,
            &CalculateRequest {
                project_id: fx.project_id.clone(),
                scenario: None,
            },
        )
        .unwrap();
        let expected = serde_json::to_value(crate::agent_bridge::calculation::build_response(
            &p, &s, &snap, &result,
        ))
        .unwrap();
        assert_eq!(v["result"]["metrics"], expected["metrics"]);
        assert_eq!(v["result"]["cashflow"], expected["cashflow"]);
        if case["overrides"].as_array().unwrap().is_empty() {
            let (status, body) = post(
                CALCULATE_ROUTE,
                json!({"sessionId":"bound","projectId":fx.project_id}),
            );
            assert_eq!(status, 200);
            let saved: Value = serde_json::from_str(&body).unwrap();
            assert_eq!(saved["basis"], "saved_snapshot");
            assert_eq!(v["result"]["metrics"], saved["metrics"]);
            assert_eq!(v["result"]["cashflow"], saved["cashflow"]);
        }
        if case["overrides"][0]["subject"] == "rev_it_integration" {
            let saved = crate::benefit::calculator::calculate_ict_benefit(input.clone()).unwrap();
            assert_ne!(result.npv, saved.npv);
            assert_ne!(result.margin_rate, saved.margin_rate);
        }
    }
    assert_eq!(changes(), before, "all routes must remain read-only");
    assert_eq!(
        serde_json::to_value(
            fx.service
                .get_snapshots(&fx.pre_scheme_id)
                .unwrap()
                .into_iter()
                .find(|s| s.version == 2)
                .unwrap()
                .input_params
        )
        .unwrap(),
        raw
    );
    let (status, body) = post(
        "/lamber-bridge/query-projects",
        json!({"sessionId":"general","query":{"irr":{"gt":0.08}}}),
    );
    assert_eq!(status, 422);
    assert!(body.contains("未计算 IRR"));
}

#[test]
fn selection_fee_tools_preserve_backend_results_errors_and_general_scope() {
    use crate::agent_bridge::selection_fee::{FORWARD_ROUTE, REVERSE_ROUTE};
    let fx = build_fixture("fee-tools");
    let runtime = scoped_workspace(&fx);
    let bindings = Arc::new(ProjectBindings::default());
    bindings
        .register("general", fixture_binding(&fx, None))
        .unwrap();
    bindings
        .register("bound", fixture_binding(&fx, Some(&fx.project_id)))
        .unwrap();
    let server =
        BridgeServer::start(crate::agent_bridge::workspace_handler(runtime, bindings)).unwrap();
    let post = |route: &str, body: Value| {
        http_post(&server.origin(), route, server.token(), &body.to_string())
    };
    for session in ["bound", "general"] {
        for quote in [
            "12826",
            "51410",
            "106000",
            "106000.01",
            "300000",
            "1060000",
            "2000000",
        ] {
            let expected =
                crate::benefit::calculator::calculate_selection_fee(quote.into(), "-50".into())
                    .unwrap();
            let (status, body) = post(
                FORWARD_ROUTE,
                json!({"sessionId":session,"quote":quote,"markup":"-50"}),
            );
            assert_eq!(status, 200, "{body}");
            assert_eq!(
                serde_json::from_str::<Value>(&body).unwrap(),
                serde_json::to_value(&expected).unwrap()
            );
            // Zero quotation with a negative markup produces a negative limit, which is rejected by the existing backend.
            if quote != "0" {
                let expected = crate::benefit::calculator::reverse_calculate_selection_fee(
                    expected.final_limit,
                    "-50".into(),
                )
                .unwrap();
                let (status, body) = post(
                    REVERSE_ROUTE,
                    json!({"sessionId":session,"limit":expected.final_limit,"markup":"-50"}),
                );
                assert_eq!(status, 200, "{body}");
                assert_eq!(
                    serde_json::from_str::<Value>(&body).unwrap(),
                    serde_json::to_value(expected).unwrap()
                );
            }
        }
        for invalid in [
            "",
            "十万",
            "100,000",
            "10万元",
            "NaN",
            "Infinity",
            "-1",
            "1.001",
            "10000000000000000000",
        ] {
            let expected =
                crate::benefit::calculator::calculate_selection_fee(invalid.into(), "0".into())
                    .unwrap_err();
            let (status, body) = post(
                FORWARD_ROUTE,
                json!({"sessionId":session,"quote":invalid,"markup":"0"}),
            );
            assert_eq!(status, 422);
            assert_eq!(
                serde_json::from_str::<Value>(&body).unwrap()["error"],
                expected
            );
        }
        let expected = crate::benefit::calculator::reverse_calculate_selection_fee(
            "106600".into(),
            "0".into(),
        )
        .unwrap_err();
        let (status, body) = post(
            REVERSE_ROUTE,
            json!({"sessionId":session,"limit":"106600","markup":"0"}),
        );
        assert_eq!(status, 422);
        assert_eq!(
            serde_json::from_str::<Value>(&body).unwrap()["error"],
            expected
        );
    }
    let (_, body) = post(
        FORWARD_ROUTE,
        json!({"sessionId":"general","quote":"300000","markup":"0"}),
    );
    let v: Value = serde_json::from_str(&body).unwrap();
    assert_eq!(v["selection_fee_excl"], "2292.41");
    assert_eq!(v["selection_fee_incl"], "2429.95");
    assert_eq!(v["final_limit"], "302429.95");
    for body in [
        json!({"sessionId":"absent","quote":"300000","markup":"0"}),
        json!({"sessionId":"general","projectId":fx.project_id,"quote":"300000","markup":"0"}),
    ] {
        assert_ne!(post(FORWARD_ROUTE, body).0, 200);
    }
}

fn populated_benefit_fixture(name: &str) -> Fixture {
    let fx = build_fixture(name);
    let conn = crate::db::init_db(&fx._db_path).unwrap();
    let repo = SqliteProjectRepository::new(Arc::new(Mutex::new(conn)));
    let witness: Value =
        serde_json::from_str(include_str!("fixtures/benefit-simulation-desktop.json")).unwrap();
    let mut input: IctInput = serde_json::from_value(witness["cases"][0]["saved"].clone()).unwrap();
    input.rev_it_integration.split_parts = Some(vec![
        crate::benefit::models::IctTaxSplitPart {
            incl_tax: "53000".into(),
            excl_tax: "50000".into()
        };
        2
    ]);
    repo.save_snapshot(&BenefitAnalysisSnapshot {
        id: uuid::Uuid::new_v4().to_string(),
        scheme_id: fx.pre_scheme_id.clone(),
        project_id: fx.project_id.clone(),
        version: 2,
        input_params: input.clone(),
        output_metrics: crate::benefit::calculator::calculate_ict_benefit(input).unwrap(),
        fingerprint: "desktop-28-subject-witness".into(),
        created_at: "2026-09-09T00:00:00Z".into(),
    })
    .unwrap();
    fx
}
#[test]
#[ignore = "creates a synthetic workspace solely for Computer Use acceptance"]
fn prepare_desktop_benefit_fixture() {
    let fx = populated_benefit_fixture("28科目桌面对照");
    let root = std::env::temp_dir().join(format!(
        "lamber-ai-benefit-desktop-{}",
        uuid::Uuid::new_v4()
    ));
    std::fs::create_dir(&root).unwrap();
    let conn = crate::db::init_db(&fx._db_path).unwrap();
    conn.execute(
        "VACUUM INTO ?1",
        [root.join(".lamber.sqlite").to_str().unwrap()],
    )
    .unwrap();
    let now = chrono::Utc::now().to_rfc3339();
    let manifest = crate::workspace::WorkspaceManifest {
        app: "Lamber".into(),
        workspace_version: 1,
        workspace_id: uuid::Uuid::new_v4().to_string(),
        name: "AI测算 A/B/C 合成验收".into(),
        created_at: now.clone(),
        last_opened_at: now,
    };
    std::fs::write(
        root.join(".lamber.workspace.json"),
        serde_json::to_string_pretty(&manifest).unwrap(),
    )
    .unwrap();
    std::fs::write(std::env::temp_dir().join("lamber-ai-benefit-desktop.json"),serde_json::to_string_pretty(&json!({"root":root,"projectId":fx.project_id,"preSchemeId":fx.pre_scheme_id,"postSchemeId":fx.post_scheme_id})).unwrap()).unwrap();
}

#[test]
#[ignore = "requires real DEEPSEEK_API_KEY and provisioned plugin"]
fn benefit_tools_real_model_saved_hypothetical_and_fee_errors() {
    use std::io::Write;
    assert!(has_api_key(), "real key required, never skip");
    let fx = populated_benefit_fixture("benefit-real-model");
    let runtime = scoped_workspace(&fx);
    let bindings = Arc::new(ProjectBindings::default());
    let prepare: benefit_simulation::Preparer = Arc::new(move |job| {
        let mut child = std::process::Command::new("node")
            .arg(repo_root_for_tests().join("src-ui/scripts/test_benefit_simulation.cjs"))
            .arg("--prepare")
            .stdin(std::process::Stdio::piped())
            .stdout(std::process::Stdio::piped())
            .stderr(std::process::Stdio::piped())
            .spawn()
            .map_err(|e| e.to_string())?;
        child
            .stdin
            .take()
            .unwrap()
            .write_all(&serde_json::to_vec(&job).unwrap())
            .map_err(|e| e.to_string())?;
        let output = child.wait_with_output().map_err(|e| e.to_string())?;
        if !output.status.success() {
            return Err(String::from_utf8_lossy(&output.stderr).into());
        }
        serde_json::from_slice(&output.stdout).map_err(|e| e.to_string())
    });
    let next = crate::agent_bridge::workspace_handler(runtime.clone(), bindings.clone());
    let handler = benefit_simulation::handler(runtime.clone(), bindings.clone(), prepare, next);
    let calls = Arc::new(Mutex::new(Vec::<Value>::new()));
    let records = calls.clone();
    let server=BridgeServer::start(Arc::new(move|path,body|{let reply=handler(path,body);if path!="/lamber-bridge/authorize"{records.lock().unwrap().push(json!({"route":path,"request":serde_json::from_str::<Value>(body).unwrap_or(Value::Null),"status":reply.status,"response":serde_json::from_str::<Value>(&reply.body).unwrap_or(Value::Null)}));}reply})).unwrap();
    let mut config = DshLaunchConfig::from_repo_root(&repo_root_for_tests());
    config.bridge_url = server.origin();
    config.bridge_token = server.token().into();
    let log = Arc::new(EventLog::default());
    let acp = stage2_runtime(&config, &log);
    let session = acp.new_session(&repo_root_for_tests()).unwrap();
    bindings
        .register(&session, fixture_binding(&fx, Some(&fx.project_id)))
        .unwrap();
    let mut answers = Vec::new();
    for prompt in [format!("当前绑定项目ID是{}。请查询当前已保存方案的全部收入与支出科目，并告诉我当前利润率、NPV及IRR。",fx.project_id),"如果集成收入改成800000元呢？请试算，列出联动变化、全部效益指标，以逐年完整金额列出每组年度现金流的覆盖前后对照，以及前四年每年的现金流入、流出、净现金流原始金额，保留到分。".into(),"我没有保存刚才的假设。当前项目已保存的利润率和NPV到底是多少？".into()]{
        log.state.lock().unwrap().clear();acp.prompt(&session,&prompt).unwrap();let answer=await_stage2_end(&log);answers.push(json!({"prompt":prompt,"answer":answer}));
    }
    let general = acp.new_session(&repo_root_for_tests()).unwrap();
    bindings
        .register(&general, fixture_binding(&fx, None))
        .unwrap();
    for prompt in [
        "含税报价300000元，浮动0，甄选服务费和限价是多少？",
        "含税限价106600元、浮动0，精确反算报价是多少？",
        "含税限价51834元、浮动0，报价有几个精确解？请列全。",
    ] {
        log.state.lock().unwrap().clear();
        acp.prompt(&general, prompt).unwrap();
        let answer = await_stage2_end(&log);
        answers.push(json!({"prompt":prompt,"answer":answer}));
    }
    let calls = calls.lock().unwrap();
    let simulation = calls
        .iter()
        .find(|c| c["route"] == benefit_simulation::ROUTE && c["status"] == 200)
        .expect("successful real-model simulation");
    assert_eq!(simulation["response"]["basis"], "hypothetical");
    assert!(answers[1]["answer"].as_str().unwrap().contains("假设"));
    let annual_answer = answers[1]["answer"].as_str().unwrap().replace(',', "");
    for row in simulation["response"]["result"]["cashflow"].as_array().unwrap().iter().take(4) {
        for key in ["cashIn", "cashOut", "netCash"] {
            let amount = row[key].as_str().unwrap();
            assert!(annual_answer.contains(amount), "model omitted exact {key} {amount}: {annual_answer}");
        }
    }

    assert!(answers[0]["answer"].as_str().unwrap().contains("保存"));
    assert!(answers[2]["answer"].as_str().unwrap().contains("保存"));
    assert!(answers[3]["answer"]
        .as_str()
        .unwrap()
        .replace(',', "")
        .contains("2429.95"));
    assert!(
        answers[4]["answer"].as_str().unwrap().contains("无")
            || answers[4]["answer"].as_str().unwrap().contains("没有")
    );
    let multi = calls
        .iter()
        .rev()
        .find(|c| {
            c["response"]["quote_candidates"]
                .as_array()
                .is_some_and(|a| a.len() > 1)
        })
        .expect("all precise reverse candidates");
    for q in multi["response"]["quote_candidates"].as_array().unwrap() {
        assert!(answers[5]["answer"]
            .as_str()
            .unwrap()
            .replace(',', "")
            .contains(q.as_str().unwrap()));
    }
    for change in simulation["response"]["linkedChanges"].as_array().unwrap() {
        if change["kind"] != "annual_cashflow" { continue; }
        for key in ["before", "after"] {
            let values: Value = serde_json::from_str(change[key].as_str().unwrap()).unwrap();
            for value in values.as_array().unwrap().iter().take(4) {
                assert!(annual_answer.contains(value.as_str().unwrap()), "linked annual amount missing: {value}");
            }
        }
    }
    // Any decimal figure in the model answer must be grounded in a returned field.
    // This also catches invented financial numbers in notes after a correct table.
    fn collect_numbers(value: &Value, allowed: &mut Vec<rust_decimal::Decimal>) {
        match value {
            Value::Array(values) => values.iter().for_each(|v| collect_numbers(v, allowed)),
            Value::Object(values) => values.values().for_each(|v| collect_numbers(v, allowed)),
            Value::String(text) => {
                if let Ok(number) = text.parse::<rust_decimal::Decimal>() {
                    allowed.push(number);
                } else if let Ok(parsed) = serde_json::from_str::<Value>(text) {
                    collect_numbers(&parsed, allowed);
                }
            }
            Value::Number(number) => {
                if let Ok(number) = number.to_string().parse::<rust_decimal::Decimal>() { allowed.push(number); }
            }
            _ => {}
        }
    }
    let mut allowed = Vec::new();
    collect_numbers(&simulation["response"], &mut allowed);
    for call in calls.iter().filter(|c| c["route"] == benefit_access::READ_ROUTE || c["route"] == "/lamber-bridge/calculate") {
        collect_numbers(&call["response"], &mut allowed);
    }
    // Ratios may be rendered as percentages without losing their original precision.
    for result in std::iter::once(&simulation["response"]["result"])
        .chain(calls.iter().filter(|c| c["route"] == "/lamber-bridge/calculate").map(|c| &c["response"])) {
        for key in ["npvRate", "marginRate", "itNpvRate", "itMarginRate"] {
            if let Some(raw) = result["metrics"][key].as_str() {
                allowed.push(raw.parse::<rust_decimal::Decimal>().unwrap() * rust_decimal::Decimal::from(100));
            }
        }
    }
    for token in annual_answer.split(|c: char| !c.is_ascii_digit() && c != '.' && c != '-') {
        if !token.contains('.') { continue; }
        if let Ok(number) = token.trim_matches('.').parse::<rust_decimal::Decimal>() {
            assert!(allowed.contains(&number), "ungrounded decimal in model narrative: {token}");
        }
    }
    let read = calls
        .iter()
        .find(|c| c["route"] == benefit_access::READ_ROUTE && c["status"] == 200)
        .expect("real model read");
    assert_eq!(read["response"]["subjects"][0]["split"], true);
    let snapshots = fx.service.get_snapshots(&fx.pre_scheme_id).unwrap();
    assert_eq!(snapshots.len(), 2);
    assert_eq!(
        snapshots
            .iter()
            .find(|s| s.version == 2)
            .unwrap()
            .input_params
            .rev_it_integration
            .incl_tax,
        "106000"
    );
    std::fs::write(repo_root_for_tests().join("docs/verification/ai-benefit-real-model-evidence.json"),serde_json::to_string_pretty(&json!({"method":"real DeepSeek + dsh + Rust routes; preparation executes shared desktop TypeScript in a test-only Node host; actual main-window IPC is verified separately in desktop acceptance", "answers":answers,"calls":*calls,"snapshotUnchanged":true})).unwrap()).unwrap();
}

#[test]
#[ignore = "requires real DEEPSEEK_API_KEY and provisioned plugin"]
fn structure_reverse_real_model_preserves_application_receipts() {
    assert!(has_api_key(), "real key required, never skip");
    let fx = populated_benefit_fixture("structure-reverse-real");
    let runtime = scoped_workspace(&fx);
    let bindings = Arc::new(ProjectBindings::default());
    let handler = crate::agent_bridge::workspace_handler(runtime, bindings.clone());
    let calls = Arc::new(Mutex::new(Vec::<String>::new()));
    let records = calls.clone();
    let server = BridgeServer::start(Arc::new(move |path, body| {
        if path != "/lamber-bridge/authorize" { records.lock().unwrap().push(path.to_string()); }
        handler(path, body)
    })).unwrap();
    let mut config = DshLaunchConfig::from_repo_root(&repo_root_for_tests());
    config.bridge_url = server.origin();
    config.bridge_token = server.token().into();
    let log = Arc::new(EventLog::default());
    let acp = stage2_runtime(&config, &log);
    let fixture: Value = serde_json::from_str(include_str!("fixtures/structure-reverse-receipts.json")).unwrap();
    let mut answers = Vec::new();
    for case in fixture["cases"].as_array().unwrap() {
        let session = acp.new_session(&repo_root_for_tests()).unwrap();
        bindings.register(&session, fixture_binding(&fx, Some(&fx.project_id))).unwrap();
        log.state.lock().unwrap().clear();
        let prompt = format!("{}\n本会话应用操作回执（真实应用结果，禁止重算或改写）：\n{}\n用户：刚才反算为什么没成功？请把应用的原错误完整告诉我。",
            fixture["prompt"].as_str().unwrap(), case["receipt"].as_str().unwrap());
        acp.prompt(&session, &prompt).unwrap();
        let answer = await_stage2_end(&log);
        answers.push(json!({"case":case["name"],"answer":answer}));
        assert!(answer.contains(case["message"].as_str().unwrap()), "model compressed or rewrote the original range error: {answer}");
        assert!(!answer.contains("已为你调整"));
    }
    let session = acp.new_session(&repo_root_for_tests()).unwrap();
    bindings.register(&session, fixture_binding(&fx, Some(&fx.project_id))).unwrap();
    log.state.lock().unwrap().clear();
    acp.prompt(&session, &format!("{}\n本会话应用操作回执：\n{}\n用户：刚才目标是多少、实际达到多少？逐科目告诉我金额和每年收款计划发生了什么变化，不能省略尾差。",
        fixture["prompt"].as_str().unwrap(), fixture["success"]["receipt"].as_str().unwrap())).unwrap();
    let answer = await_stage2_end(&log);
    answers.push(json!({"case":"success","answer":answer}));
    assert!(answer.contains("目标") && (answer.contains("达成") || answer.contains("实际")));
    for amount in ["12.0000%", "12.0010%", "100.00", "120.02", "25.00", "30.01", "29.99"] {
        assert!(answer.contains(amount), "model omitted receipt value {amount}: {answer}");
    }
    let session = acp.new_session(&repo_root_for_tests()).unwrap();
    bindings.register(&session, fixture_binding(&fx, Some(&fx.project_id))).unwrap();
    log.state.lock().unwrap().clear();
    acp.prompt(&session, &format!("{}\n用户：我想做结构反算，还没决定目标。现在该怎么操作？不要查询数据。", fixture["prompt"].as_str().unwrap())).unwrap();
    let answer = await_stage2_end(&log);
    answers.push(json!({"case":"no_invented_target","answer":answer}));
    assert!(answer.contains("科目") && answer.contains("范围"));
    assert!(!answer.contains('%') && !answer.contains('％'), "model invented a percentage: {answer}");
    let routes = calls.lock().unwrap().clone();
    assert!(!routes.iter().any(|route| route.contains("fill") || route.contains("write") || route.contains("reverse") || route.contains("simulate")), "card receipts must not lead to new write or reverse tool calls");
    std::fs::write(repo_root_for_tests().join("docs/verification/ai-structure-reverse-real-model-evidence.json"),
        serde_json::to_string_pretty(&json!({"kind":"real dsh runtime with production frontend prompt/receipt fixtures; financial outcomes are controlled test receipts", "answers":answers,"routes":routes})).unwrap()).unwrap();
    drop(acp);
}
