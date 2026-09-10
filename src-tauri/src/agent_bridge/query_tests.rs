use super::*;
use crate::agent_bridge::{
    project_bindings::{ProjectBindings, AUTHORIZE_ROUTE},
    project_query::{self, MAX_LIMIT, QUERY_ROUTE},
};

fn query_fixture(count: usize) -> (Fixture, Arc<crate::workspace::WorkspaceRuntime>) {
    let fx = build_fixture("aggregate-query");
    let runtime = scoped_workspace(&fx);
    let repo = SqliteProjectRepository::new(runtime.require_db().unwrap());
    for n in 0..count {
        let mut p = if n == 0 {
            fx.service.get_project(&fx.project_id).unwrap().unwrap()
        } else {
            blank_project()
        };
        if n != 0 {
            p.id = format!("query-project-{n:03}");
        }
        p.name = if n == 0 {
            "当前甲项目".into()
        } else {
            format!("其他乙项目{n}")
        };
        p.customer_name = "聚合测试客户".into();
        p.status = "实施中".into();
        p.benefit_status = if n == 2 {
            "outdated".into()
        } else {
            "normal".into()
        };
        p.created_at = "2025-06-01T00:00:00Z".into();
        p.updated_at = "2026-09-06T00:00:00Z".into();
        p.total_revenue_incl = 1000.1;
        p.total_cost_incl = 500.2;
        p.note = Some("FORBIDDEN_NOTE".into());
        p.folder_path = Some("/private/FORBIDDEN_PATH".into());
        p.main_document_path = Some("FORBIDDEN_DOCUMENT".into());
        p.main_budget_file_path = Some("FORBIDDEN_BUDGET".into());
        p.logs = vec![crate::benefit::models::ProjectLog {
            id: "private".into(),
            timestamp: "now".into(),
            description: "FORBIDDEN_LOG".into(),
        }];
        p.summary_metrics = Some(crate::benefit::models::SummaryMetrics {
            margin_rate: if n == 0 { "0.13".into() } else { "0.87".into() },
            npv: "123.45".into(),
            npv_rate: "20%".into(),
            irr: "--".into(),
            dynamic_payback: "1.5".into(),
            risk_level: "low".into(),
        });
        repo.save_project(&p).unwrap();
    }
    (fx, runtime)
}
#[test]
fn aggregate_query_permission_matrix_and_bounded_projection() {
    let (fx, runtime) = query_fixture(65);
    let bindings = Arc::new(ProjectBindings::default());
    bindings
        .register("bound", fixture_binding(&fx, Some(&fx.project_id)))
        .unwrap();
    bindings
        .register("general", fixture_binding(&fx, None))
        .unwrap();
    let server = BridgeServer::start(crate::agent_bridge::workspace_handler(
        runtime.clone(),
        bindings,
    ))
    .unwrap();
    let post = |route: &str, body: Value| {
        http_post(&server.origin(), route, server.token(), &body.to_string())
    };
    let changes_before = runtime
        .require_db()
        .unwrap()
        .lock()
        .unwrap()
        .query_row("SELECT total_changes()", [], |r| r.get::<_, i64>(0))
        .unwrap();
    for session in ["bound", "general"] {
        let (status, body) = post(
            QUERY_ROUTE,
            serde_json::json!({"sessionId":session,"query":{"limit":9999}}),
        );
        assert_eq!(status, 200, "{body}");
        let v: Value = serde_json::from_str(&body).unwrap();
        assert_eq!(v["matchedCount"], 65);
        assert_eq!(v["returnedCount"], MAX_LIMIT);
        assert_eq!(v["truncated"], true);
        assert!(v["message"]
            .as_str()
            .unwrap()
            .contains("命中 65 条，仅返回 50 条"));
        assert!(v["totals"].is_null());
        assert_eq!(v["mixedStages"], true);
        assert_eq!(v["stageTotals"][0]["matchedCount"], 1);
        assert_eq!(v["stageTotals"][0]["totals"]["totalRevenueIncl"], "1000.1");
        assert_eq!(v["stageTotals"][1]["matchedCount"], 64);
        assert_eq!(v["stageTotals"][1]["totals"]["totalRevenueIncl"], "64006.4");
        assert_eq!(v["stageTotals"][1]["totals"]["totalCostIncl"], "32012.8");
        assert!(v["notice"]
            .as_str()
            .unwrap()
            .contains("不得用作当前绑定项目"));
        assert_eq!(
            v["boundProjectId"],
            if session == "bound" {
                serde_json::json!(fx.project_id)
            } else {
                Value::Null
            }
        );
        for forbidden in [
            "FORBIDDEN",
            "folderPath",
            "folder_path",
            "note",
            "logs",
            "mainDocumentPath",
            "snapshot",
            "supplier",
            "template",
        ] {
            assert!(!body.contains(forbidden), "{forbidden} leaked");
        }
        // Allow-list schema assertion protects against accidental future Project serialization.
        let row = v["projects"][0].as_object().unwrap();
        assert_eq!(row.len(), 19);
        assert_eq!(row["summaryMetrics"]["npvRate"], 0.2);
        assert!(row["summaryMetrics"]["irr"].is_null());
        assert_eq!(
            post(
                AUTHORIZE_ROUTE,
                serde_json::json!({"sessionId":session,"tool":"query_projects"})
            )
            .0,
            200
        );
        assert_eq!(
            post(
                AUTHORIZE_ROUTE,
                serde_json::json!({"sessionId":session,"tool":"future_no_project_tool"})
            )
            .0,
            403
        );
        assert_eq!(
            post(
                CALCULATE_ROUTE,
                serde_json::json!({"sessionId":session,"projectId":"query-project-001"})
            )
            .0,
            403
        );
    }
    assert_eq!(
        post(
            CALCULATE_ROUTE,
            serde_json::json!({"sessionId":"bound","projectId":fx.project_id})
        )
        .0,
        200
    );
    assert_eq!(
        post(
            CALCULATE_ROUTE,
            serde_json::json!({"sessionId":"general","projectId":fx.project_id})
        )
        .0,
        403
    );
    assert_eq!(
        post(
            AUTHORIZE_ROUTE,
            serde_json::json!({"sessionId":"general","tool":"write_test_marker"})
        )
        .0,
        403
    );
    for body in [
        serde_json::json!({"query":{}}),
        serde_json::json!({"sessionId":"missing","query":{}}),
        serde_json::json!({"sessionId":"general","projectId":"query-project-001","query":{}}),
    ] {
        assert_eq!(post(QUERY_ROUTE, body).0, 403);
    }
    for query in [
        serde_json::json!({"sql":"SELECT * FROM projects"}),
        serde_json::json!({"limit":-1}),
        serde_json::json!({"sortBy":"note"}),
    ] {
        assert_eq!(
            post(
                QUERY_ROUTE,
                serde_json::json!({"sessionId":"bound","query":query})
            )
            .0,
            400
        );
    }
    for query in [
        serde_json::json!({"limit":0}),
        serde_json::json!({"marginRate":{"gt":0.2,"lte":0.2}}),
        serde_json::json!({"createdFrom":"bad date"}),
    ] {
        assert_eq!(
            post(
                QUERY_ROUTE,
                serde_json::json!({"sessionId":"general","query":query})
            )
            .0,
            422
        );
    }
    let result = post(
        QUERY_ROUTE,
        serde_json::json!({"sessionId":"bound","query":{"customerName":"测试","status":"实施中","benefitStatus":"normal","createdFrom":"2026-01-01","createdBefore":"2027-01-01","updatedFrom":"2026-09-06T08:00:00+08:00","totalCostIncl":{"gte":500.2,"lte":500.2},"marginRate":{"lt":0.2},"sortBy":"marginRate","sortOrder":"asc"}}),
    );
    assert_eq!(result.0, 200);
    let v: Value = serde_json::from_str(&result.1).unwrap();
    assert_eq!(v["matchedCount"], 1);
    assert_eq!(v["projects"][0]["id"], fx.project_id);
    assert_eq!(v["projects"][0]["isBoundProject"], true);
    let empty = post(
        QUERY_ROUTE,
        serde_json::json!({"sessionId":"general","query":{"irr":{"gt":0}}}),
    );
    assert_eq!(empty.0, 422);
    assert!(serde_json::from_str::<Value>(&empty.1).unwrap()["error"].as_str().unwrap().contains("未计算 IRR"));
    assert_eq!(
        changes_before,
        runtime
            .require_db()
            .unwrap()
            .lock()
            .unwrap()
            .query_row("SELECT total_changes()", [], |r| r.get::<_, i64>(0))
            .unwrap(),
        "queries must not write"
    );
}

#[test]
#[ignore = "needs DEEPSEEK_API_KEY plus provisioned plugin; saves synthetic answer for human review"]
fn aggregate_query_real_sessions_and_followup_identity() {
    assert!(has_api_key(), "real key required");
    let (fx, runtime) = query_fixture(2);
    let bindings = Arc::new(ProjectBindings::default());
    let hits = Arc::new(Mutex::new(Vec::<Value>::new()));
    let capture = hits.clone();
    let handler = crate::agent_bridge::workspace_handler(runtime, bindings.clone());
    let bridge = BridgeServer::start(Arc::new(move |path, body| {
        let reply = handler(path, body);
        if path == QUERY_ROUTE && reply.status == 200 {
            capture
                .lock()
                .unwrap()
                .push(serde_json::from_str(&reply.body).unwrap());
        }
        reply
    }))
    .unwrap();
    let mut config = DshLaunchConfig::from_repo_root(&repo_root_for_tests());
    config.bridge_url = bridge.origin();
    config.bridge_token = bridge.token().into();
    let log = Arc::new(EventLog::default());
    let acp = stage2_runtime(&config, &log);
    let mut answers = Vec::new();
    for (label, project) in [("bound", Some(fx.project_id.as_str())), ("general", None)] {
        let hits_before = hits.lock().unwrap().len();
        let session = acp.new_session(&config.cwd).unwrap();
        bindings
            .register(&session, fixture_binding(&fx, project))
            .unwrap();
        // Exercise the production frontend scope guidance, not a more permissive test prompt.
        let policy = std::process::Command::new("node")
            .current_dir(repo_root_for_tests().join("src-ui"))
            .args(["-e", "const ts=require('typescript'),fs=require('fs');const m={exports:{}};new Function('exports','module',ts.transpileModule(fs.readFileSync('src/ai/sessionScopePolicy.ts','utf8'),{compilerOptions:{module:ts.ModuleKind.CommonJS}}).outputText)(m.exports,m);process.stdout.write(m.exports.sessionScopePrompt(process.argv[1]||null));", project.unwrap_or("")])
            .output().unwrap();
        assert!(policy.status.success());
        let scope = String::from_utf8(policy.stdout).unwrap();
        acp.prompt(&session,&format!("{scope} 请实际调用query_projects查询客户名包含‘聚合测试客户’的所有项目，按毛利率降序，列出每个项目名及已保存毛利率、命中总数、含税总成本。不要重新测算或猜测数据。只需简短回答。")).unwrap();
        let answer = await_stage2_end(&log);
        assert!(answer.contains("13") && answer.contains("87"), "{answer}");
        assert!(
            hits.lock().unwrap().len() > hits_before,
            "model must really query the database"
        );
        answers.push(serde_json::json!({"sessionType":label,"queryAnswer":answer}));
        log.state.lock().unwrap().clear();
        if project.is_some() {
            acp.prompt(&session, "我这个项目毛利率多少？").unwrap();
            let followup = await_stage2_end(&log);
            assert!(
                followup.contains("13")
                    && followup.contains("当前甲项目")
                    && (followup.contains("甄选前") || followup.contains("限价口径"))
                    && !followup.contains("87"),
                "{followup}"
            );
            answers.last_mut().unwrap()["followupAnswer"] = serde_json::json!(followup);
            log.state.lock().unwrap().clear();
        }
    }
    let result = serde_json::json!({"fixture":{"boundProjectName":"当前甲项目","boundMarginRate":0.13,"otherMarginRate":0.87},"actualModelAnswers":answers,"actualToolResponses":*hits.lock().unwrap(),"humanReview":"pending"});
    let path =
        repo_root_for_tests().join("docs/verification/cross-project-query-real-evidence.json");
    std::fs::write(path, serde_json::to_string_pretty(&result).unwrap()).unwrap();
}

#[test]
fn aggregate_query_follows_saved_default_stage_without_reading_financial_details() {
    let (fx, runtime) = query_fixture(2);
    let repo = SqliteProjectRepository::new(runtime.require_db().unwrap());
    let query =
        serde_json::from_value(serde_json::json!({"sortBy":"marginRate","sortOrder":"desc"}))
            .unwrap();
    let run = || {
        serde_json::to_value(
            project_query::query_projects(&fx.service, &query, Some(&fx.project_id)).unwrap(),
        )
        .unwrap()
    };
    let first = run();
    let current = |v: &Value| {
        v["projects"]
            .as_array()
            .unwrap()
            .iter()
            .find(|p| p["id"] == fx.project_id)
            .unwrap()
            .clone()
    };
    assert_eq!(current(&first)["stage"], "pre_selection");
    assert_eq!(current(&first)["defaultSchemeId"], fx.pre_scheme_id);
    assert_eq!(current(&first)["schemeUpdatedAt"], "2026-01-01T00:00:00Z");
    let mut p = fx.service.get_project(&fx.project_id).unwrap().unwrap();
    p.default_scheme_id = Some(fx.post_scheme_id.clone());
    repo.save_project(&p).unwrap();
    let second = run();
    assert_eq!(current(&second)["stage"], "post_selection");
    assert_eq!(current(&second)["stageLabel"], "甄选后（中标口径）");
    assert_eq!(current(&second)["schemeUpdatedAt"], "2026-01-02T00:00:00Z");
    assert_eq!(
        current(&second)["summaryMetrics"],
        current(&first)["summaryMetrics"],
        "metadata must not recalculate saved metrics"
    );
    // Broken/default-less/unknown-stage pointers must never be guessed from other schemes.
    for pointer in [None, Some("missing-scheme".to_string())] {
        p.default_scheme_id = pointer;
        repo.save_project(&p).unwrap();
        let result = run();
        assert_eq!(current(&result)["stage"], "unlabeled");
        assert_eq!(current(&result)["schemeName"], Value::Null);
        assert_eq!(result["mixedStages"], false);
        assert_eq!(result["totals"]["totalRevenueIncl"], "2000.2");
        assert!(result["comparisonNotice"]
            .as_str()
            .unwrap()
            .contains("未标注不代表"));
    }
    let mut scheme = fx
        .service
        .get_schemes(&fx.project_id)
        .unwrap()
        .into_iter()
        .find(|s| s.id == fx.post_scheme_id)
        .unwrap();
    scheme.stage = None;
    repo.save_scheme(&scheme).unwrap();
    p.default_scheme_id = Some(scheme.id.clone());
    repo.save_project(&p).unwrap();
    assert_eq!(current(&run())["stage"], "unlabeled");
    scheme.stage = Some("future_stage".into());
    repo.save_scheme(&scheme).unwrap();
    assert_eq!(current(&run())["stage"], "unlabeled");
    // Two known but different bases, even with limit=1, must retain all stage totals.
    p.default_scheme_id = Some(fx.pre_scheme_id.clone());
    repo.save_project(&p).unwrap();
    let mut other = fx
        .service
        .get_project("query-project-001")
        .unwrap()
        .unwrap();
    scheme.id = "other-post".into();
    scheme.project_id = other.id.clone();
    scheme.stage = Some("post_selection".into());
    repo.save_scheme(&scheme).unwrap();
    other.default_scheme_id = Some(scheme.id);
    repo.save_project(&other).unwrap();
    let limited = serde_json::from_value(
        serde_json::json!({"limit":1,"sortBy":"marginRate","sortOrder":"desc"}),
    )
    .unwrap();
    let result =
        serde_json::to_value(project_query::query_projects(&fx.service, &limited, None).unwrap())
            .unwrap();
    assert!(result["totals"].is_null());
    assert_eq!(result["mixedStages"], true);
    assert_eq!(result["stageTotals"].as_array().unwrap().len(), 2);
    assert_eq!(result["matchedCount"], 2);
    assert_eq!(result["returnedCount"], 1);
}
