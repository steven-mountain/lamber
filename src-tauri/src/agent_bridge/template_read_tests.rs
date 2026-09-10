use super::*;
use crate::{
    agent_bridge::template_read::{ROUTE, TEXT_LIMIT},
    project_state::{self, TemplateStatePayload},
};
const TEMPLATE: &str = "ICT项目需求导入表模板.docx";
fn seed(runtime: &crate::workspace::WorkspaceRuntime, id: &str, template: &str, text: &str) {
    project_state::save_template_state_locked(&mut runtime.require_db().unwrap().lock().unwrap(), id.into(), template.into(), TemplateStatePayload {
        template_name: Some(template.into()), template_type: Some("word".into()), template_path: Some("/private/secret.docx".into()), template_path_type: Some("external".into()),
        filled_data_json: serde_json::json!({"formData":{"gen_demand_service_content":text,"gen_demand_env_require":"","gen_demand_public_url":"https://example.test/tender","gen_demand_security_detail":"已完成密评","unknown":"SECRET_UNKNOWN"},"hasSecurity":true,"hasPublicUrl":true,"techItems":[{"serviceName":"网络服务","serviceDesc":"分区实施","amount":2,"unit":"项","path":"/private/secret"}],"attach1Images":[{"path":"/private/image.png"}],"secret":"DO_NOT_RETURN"}),
        field_mapping_json: serde_json::json!({"path":"/private/mapping"}), output_config_json: serde_json::json!({"outputDir":"/private/output"}),
    }, None).unwrap();
}
#[test]
fn template_read_bound_saved_projection_and_limits() {
    let fx = build_fixture("template-read");
    let runtime = scoped_workspace(&fx);
    seed(
        &runtime,
        &fx.project_id,
        TEMPLATE,
        "甲项目原文\n按联合勘察清单验收。",
    );
    let repo = SqliteProjectRepository::new(runtime.require_db().unwrap());
    let mut other = blank_project();
    other.id = "project-B".into();
    other.name = "乙项目".into();
    repo.save_project(&other).unwrap();
    seed(&runtime, &other.id, TEMPLATE, "B_PRIVATE_TEXT");
    seed(
        &runtime,
        &other.id,
        "乙专属需求导入表.docx",
        "B_ONLY_TEMPLATE",
    );
    let bindings = Arc::new(super::super::project_bindings::ProjectBindings::default());
    bindings
        .register("A", fixture_binding(&fx, Some(&fx.project_id)))
        .unwrap();
    bindings
        .register("B", fixture_binding(&fx, Some("project-B")))
        .unwrap();
    bindings
        .register("general", fixture_binding(&fx, None))
        .unwrap();
    bindings
        .register("deleted", fixture_binding(&fx, Some("deleted-project")))
        .unwrap();
    let handler = super::super::workspace_handler(runtime.clone(), bindings.clone());
    let read = |body: Value| handler(ROUTE, &body.to_string());
    let before = runtime
        .require_db()
        .unwrap()
        .lock()
        .unwrap()
        .query_row("SELECT total_changes()", [], |r| r.get::<_, i64>(0))
        .unwrap();
    let good = read(serde_json::json!({"sessionId":"A","templateId":"需求导入表"}));
    assert_eq!(good.status, 200, "{}", good.body);
    let result: Value = serde_json::from_str(&good.body).unwrap();
    assert!(!good.body.contains("B_PRIVATE_TEXT"));
    let b = read(serde_json::json!({"sessionId":"B","templateId":TEMPLATE}));
    assert!(b.body.contains("B_PRIVATE_TEXT"));
    let b_only = read(serde_json::json!({"sessionId":"A","templateId":"乙专属需求导入表.docx"}));
    assert!(!b_only.body.contains("B_ONLY_TEMPLATE"));
    assert_eq!(
        serde_json::from_str::<Value>(&b_only.body).unwrap()["hasSavedState"],
        false
    );
    assert_eq!(result["projectId"], fx.project_id);
    assert_eq!(result["templateId"], TEMPLATE);
    assert_eq!(
        result["fields"][2]["value"],
        "甲项目原文\n按联合勘察清单验收。"
    );
    assert_eq!(result["fields"][5]["value"], "");
    assert_eq!(result["fields"][0]["valueSource"], "default");
    for forbidden in [
        "/Users/",
        "/private/",
        "SECRET_UNKNOWN",
        "DO_NOT_RETURN",
        "templatePath",
        "outputConfig",
        "fieldMapping",
    ] {
        assert!(!good.body.contains(forbidden), "{forbidden}");
    }
    assert_eq!(
        result["attachments"][0]["exists"], false,
        "legacy path cannot invent a live attachment"
    );
    assert_eq!(
        runtime
            .require_db()
            .unwrap()
            .lock()
            .unwrap()
            .query_row("SELECT total_changes()", [], |r| r.get::<_, i64>(0))
            .unwrap(),
        before,
        "read must not save or audit"
    );
    for session in ["general", "missing", "deleted", ""] {
        assert_ne!(
            read(serde_json::json!({"sessionId":session,"templateId":TEMPLATE})).status,
            200
        );
    }
    for key in ["projectId", "fields", "path"] {
        assert_eq!(
            read(serde_json::json!({"sessionId":"A","templateId":TEMPLATE,key:"other-project"}))
                .status,
            400
        );
    }
    for name in ["../需求导入表.docx", "C:\\需求导入表.docx", "其他表.docx"] {
        assert_ne!(
            read(serde_json::json!({"sessionId":"A","templateId":name})).status,
            200
        );
    }
    let absent = read(serde_json::json!({"sessionId":"A","templateId":"新需求导入表.docx"}));
    assert_eq!(
        serde_json::from_str::<Value>(&absent.body).unwrap()["hasSavedState"],
        false
    );
    for tool in ["bash", "glob", "run_code"] {
        assert_ne!(
            handler(
                super::super::project_bindings::AUTHORIZE_ROUTE,
                &serde_json::json!({"sessionId":"A","tool":tool}).to_string()
            )
            .status,
            200
        );
    }
    assert_eq!(
        handler(
            super::super::project_bindings::AUTHORIZE_ROUTE,
            &serde_json::json!({"sessionId":"A","tool":"read_template_fields","projectId":"other"})
                .to_string()
        )
        .status,
        403
    );
    seed(
        &runtime,
        &fx.project_id,
        TEMPLATE,
        &"中文😀".repeat(TEXT_LIMIT),
    );
    let large = read(serde_json::json!({"sessionId":"A","templateId":TEMPLATE}));
    let large: Value = serde_json::from_str(&large.body).unwrap();
    assert_eq!(large["truncated"], true);
    assert_eq!(large["returnedCharacters"], TEXT_LIMIT);
    assert_eq!(
        large["completionState"]["formData"]["gen_demand_security_detail"], "filled",
        "truncated display must not alter completion"
    );
    assert!(large["notice"].as_str().unwrap().contains("截断"));
    std::fs::write(
        repo_root_for_tests().join("docs/verification/template-read-projection-evidence.json"),
        serde_json::to_string_pretty(&serde_json::json!({"normal":result,"truncated":large}))
            .unwrap(),
    )
    .unwrap();
}
#[test]
#[ignore = "requires real DEEPSEEK_API_KEY and provisioned plugin"]
fn template_read_real_without_page_context() {
    assert!(has_api_key());
    let fx = build_fixture("template-read-real");
    let runtime = scoped_workspace(&fx);
    seed(
        &runtime,
        &fx.project_id,
        TEMPLATE,
        "只读联调代号：青鹭472。园区分区实施，先试点后推广。验收要求完成回退演练。",
    );
    let log = Arc::new(EventLog::default());
    let prompts = Arc::new(Mutex::new(Vec::new()));
    let (_bridge, acp, bindings) = launch_dsh_with_approval(
        &fx,
        Arc::new(Mutex::new(Vec::new())),
        log.clone(),
        ApprovalStance::Reject,
        prompts.clone(),
    );
    let session = acp.new_session(&repo_root_for_tests()).unwrap();
    bindings
        .register(&session, fixture_binding(&fx, Some(&fx.project_id)))
        .unwrap();
    // No project/template payload and no active page: the unknown marker exists only in SQLite.
    acp.prompt(&session, "需求导入表填了什么？还有哪些没填？")
        .unwrap();
    let answer = await_stage2_end(&log);
    assert!(answer.contains("青鹭472"), "{answer}");
    assert!(answer.contains("部署环境"), "{answer}");
    assert!(
        prompts.lock().unwrap().is_empty(),
        "read must never request approval"
    );
    std::fs::write(repo_root_for_tests().join("docs/verification/template-read-real-evidence.json"),serde_json::to_string_pretty(&serde_json::json!({"method":"real dsh; fresh bound session; no page context; synthetic SQLite-only marker","prompt":"需求导入表填了什么？还有哪些没填？","answer":answer,"approvalCount":0,"humanGate":"pending"})).unwrap()).unwrap();
}

#[test]
fn meeting_catalog_projection_matches_saved_predicate_inputs() {
    let fx = build_fixture("meeting-projection");
    let runtime = scoped_workspace(&fx);
    let bindings = Arc::new(super::super::project_bindings::ProjectBindings::default());
    bindings.register("meeting", fixture_binding(&fx, Some(&fx.project_id))).unwrap();
    let handler = super::super::workspace_handler(runtime.clone(), bindings);
    let fixtures: Value = serde_json::from_str(include_str!("../../../src-ui/scripts/fixtures/meeting-catalog-baseline.json")).unwrap();
    let mut evidence = Vec::new();
    for fixture in fixtures["meeting"].as_array().unwrap().iter().map(|v| ("会审纪要.docx",v["state"].clone()))
        .chain(fixtures["baseline"].as_array().unwrap().iter().map(|v|(v["name"].as_str().unwrap(),v["state"].clone()))) {
        let (name,state) = fixture;
        project_state::save_template_state_locked(&mut runtime.require_db().unwrap().lock().unwrap(), fx.project_id.clone(),name.into(),TemplateStatePayload {
            template_name:Some(name.into()),template_type:Some("word".into()),template_path:None,template_path_type:None,
            filled_data_json:state.clone(),field_mapping_json:serde_json::json!({}),output_config_json:serde_json::json!({}),
        },None).unwrap();
        let reply = handler(ROUTE,&serde_json::json!({"sessionId":"meeting","templateId":name}).to_string());
        assert_eq!(reply.status,200,"{}",reply.body);
        let result:Value=serde_json::from_str(&reply.body).unwrap();
        if let Some(scale)=state.get("projectScale") {assert_eq!(&result["completionState"]["projectScale"],scale);}
        evidence.push(serde_json::json!({"name":name,"state":state,"completionState":result["completionState"],"assets":result["attachments"]}));
    }
    // A meaningful row beyond visible truncation must still complete the list.
    let mut rows=vec![serde_json::json!({"vendorName":"","amount":0});105];
    rows[104]=serde_json::json!({"vendorName":"","amount":1});
    let state=serde_json::json!({"inqVendors":rows});
    project_state::save_template_state_locked(&mut runtime.require_db().unwrap().lock().unwrap(),fx.project_id.clone(),"会审纪要.docx".into(),TemplateStatePayload {
        template_name:None,template_type:None,template_path:None,template_path_type:None,filled_data_json:state.clone(),field_mapping_json:serde_json::json!({}),output_config_json:serde_json::json!({}),
    },None).unwrap();
    let reply=handler(ROUTE,&serde_json::json!({"sessionId":"meeting","templateId":"会审纪要.docx"}).to_string());
    let result:Value=serde_json::from_str(&reply.body).unwrap();
    assert_eq!(result["returnedListCounts"]["inqVendors"],100);
    assert!(result["completionState"]["inqVendors"].as_array().unwrap().len()<=4);
    evidence.push(serde_json::json!({"name":"会审纪要.docx","state":state,"completionState":result["completionState"],"assets":result["attachments"]}));
    if let Ok(path)=std::env::var("LAMBER_MEETING_PROJECTION") {std::fs::write(path,serde_json::to_string_pretty(&evidence).unwrap()).unwrap();}
}

#[test]
fn technical_card_shared_save_is_atomic_and_rejects_stale_other_template() {
    use crate::project_state::{save_template_with_shared_tech_locked,SharedTechTarget};
    let fx=build_fixture("tech-shared-save");let runtime=scoped_workspace(&fx);
    let db=runtime.require_db().unwrap();let mut conn=db.lock().unwrap();
    let payload=|state|TemplateStatePayload {template_name:None,template_type:None,template_path:None,template_path_type:None,filled_data_json:state,field_mapping_json:serde_json::json!({}),output_config_json:serde_json::json!({})};
    let tech=serde_json::json!([{"serviceName":"合成服务","serviceDesc":"分区实施","amount":2,"unit":"套"}]);
    project_state::save_template_state_locked(&mut conn,fx.project_id.clone(),"需求导入表.docx".into(),payload(serde_json::json!({"techItems":[],"formData":{"gen_demand_service_content":"保留需求文本"}})),Some(0)).unwrap();
    project_state::save_template_state_locked(&mut conn,fx.project_id.clone(),"会审纪要.docx".into(),payload(serde_json::json!({"techItems":[],"formData":{"gen_tech_solution":"保留会审文本"},"inqVendors":[{"vendorName":"证据不可改"}]})),Some(0)).unwrap();
    let targets=||vec![SharedTechTarget{template_name:"会审纪要.docx".into(),expected:serde_json::json!([])}];
    let saved=save_template_with_shared_tech_locked(&mut conn,fx.project_id.clone(),"需求导入表.docx".into(),payload(serde_json::json!({"techItems":tech,"formData":{"gen_demand_service_content":"保留需求文本"}})),Some(1),targets()).unwrap();
    assert_eq!(saved.template_version,2);
    let other=project_state::get_template_state_locked(&conn,&fx.project_id,"会审纪要.docx").unwrap().unwrap();
    assert_eq!(other.filled_data_json["techItems"],tech);assert_eq!(other.filled_data_json["formData"]["gen_tech_solution"],"保留会审文本");assert_eq!(other.filled_data_json["inqVendors"][0]["vendorName"],"证据不可改");
    let err=save_template_with_shared_tech_locked(&mut conn,fx.project_id.clone(),"需求导入表.docx".into(),payload(serde_json::json!({"techItems":[]})),Some(2),targets()).err().unwrap();
    assert!(err.contains("TemplateStateConflict"));
    let unchanged=project_state::get_template_state_locked(&conn,&fx.project_id,"需求导入表.docx").unwrap().unwrap();assert_eq!(unchanged.template_version,2);assert_eq!(unchanged.filled_data_json["techItems"],tech);
    // Corrupt destination causes rollback even after selected row was written inside the transaction.
    conn.execute("UPDATE project_template_states SET filled_data_json='[]' WHERE project_id=?1 AND template_id='会审纪要.docx'",[&fx.project_id]).unwrap();
    assert!(save_template_with_shared_tech_locked(&mut conn,fx.project_id.clone(),"需求导入表.docx".into(),payload(serde_json::json!({"techItems":[]})),Some(2),targets()).is_err());
    assert_eq!(project_state::get_template_state_locked(&conn,&fx.project_id,"需求导入表.docx").unwrap().unwrap().template_version,2);
}

#[test]
fn selection_page_catalog_saved_projection_regression() {
    let fx=build_fixture("selection-page-catalog");
    let runtime=scoped_workspace(&fx);
    let bindings=Arc::new(super::super::project_bindings::ProjectBindings::default());
    bindings.register("catalog",fixture_binding(&fx,Some(&fx.project_id))).unwrap();
    let handler=super::super::workspace_handler(runtime.clone(),bindings);
    let mut evidence=Vec::new();
    for name in ["需求导入表.docx","立项签批表.docx","会审纪要.docx","甄选结果签批表.docx"] {
        for mode in ["single","batch"] { for filled in [false,true] {
            let text=if filled {"已填写"} else {""};
            let template=super::super::template_catalog::resolve(name,false).unwrap();
            let mut state=serde_json::json!({"formData":{},"selectionResultMode":mode,"selectionBatchName":text,"selectionBatchNameCustomized":true,
                "projectScale":"large","hasMidThree":true,"hasSingleSource":true,"procurementMethod":"comparison",
                "midThreeCode":text,"midThreeName":text,"selfThreeValue":text,"itContent":text,"ctContent":text,
                "revCollection":text,"expPayment":text,"itBusMode":text,"itFundSrc":text,
                "hasPublicUrl":false,"hasSecurity":false,"techItems":[],"inqVendors":[]});
            for field in template.fields.iter().filter(|f|f.kind=="text") {field.write(&mut state,serde_json::json!(text)).unwrap();}
            project_state::save_template_state_locked(&mut runtime.require_db().unwrap().lock().unwrap(),fx.project_id.clone(),name.into(),TemplateStatePayload {
                template_name:Some(name.into()),template_type:None,template_path:None,template_path_type:None,
                filled_data_json:state.clone(),field_mapping_json:serde_json::json!({}),output_config_json:serde_json::json!({})
            },None).unwrap();
            let reply=handler(ROUTE,&serde_json::json!({"sessionId":"catalog","templateId":name}).to_string());
            assert_eq!(reply.status,200,"{}",reply.body);
            let projection:Value=serde_json::from_str(&reply.body).unwrap();
            assert!(projection["completionState"].get("completionValues").is_none());
            if name=="甄选结果签批表.docx" {
                assert_eq!(projection["completionState"]["selectionResultMode"],mode);
                assert_eq!(projection["completionState"]["selectionBatchName"],if filled {"filled"} else {""});
                assert!(!projection["fields"].as_array().unwrap().iter().any(|f|f["key"]=="selection_batch_name"));
            }
            evidence.push(serde_json::json!({"name":name,"mode":mode,"filled":filled,"state":state,"projection":projection}));
        }}
    }
    std::fs::write(repo_root_for_tests().join("docs/verification/selection-page-projection-evidence.json"),serde_json::to_string_pretty(&evidence).unwrap()).unwrap();
}
