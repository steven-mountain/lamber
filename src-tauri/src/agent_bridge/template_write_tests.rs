use super::*;
use crate::agent_bridge::template_write::{self, ROUTE, TOOL};
use crate::project_state::{self, TemplateStatePayload};
const TEMPLATE: &str = "ICT项目需求导入表.docx";
fn payload(value: Value) -> TemplateStatePayload {
    TemplateStatePayload {
        template_name: Some(TEMPLATE.into()),
        template_type: Some("word".into()),
        template_path: Some(TEMPLATE.into()),
        template_path_type: Some("module".into()),
        filled_data_json: value,
        field_mapping_json: serde_json::json!({"keep":"mapping"}),
        output_config_json: serde_json::json!({"outputDir":"unchanged"}),
    }
}
fn write_args(project: &str) -> Value {
    serde_json::json!({"projectId":project,"templateId":TEMPLATE,"fields":{"gen_demand_env_require":"模型原始需求：直接切换整个园区网络。"}})
}
fn grant(
    gate: &Arc<ApprovalGate>,
    runtime: &crate::workspace::WorkspaceRuntime,
    bindings: &super::super::project_bindings::ProjectBindings,
    args: Value,
    approved: bool,
    modified: Option<Value>,
) -> ApprovalDecision {
    let resolver = gate.clone();
    template_write::request_approval(
        gate,
        runtime,
        bindings,
        ApprovalQuestion {
            session_id: Some("bound".into()),
            call_id: Some("call-write".into()),
            tool_name: TOOL.into(),
            reason: None,
            intent: None,
            args,
        },
        move |prompt| {
            let id = prompt.request_id.clone();
            assert!(prompt.intent.as_ref().unwrap().state_version.is_some());
            std::thread::spawn(move || {
                resolver.resolve_with_args(&id, approved, modified).unwrap()
            });
        },
    )
}
#[test]
fn template_write_requires_scope_approval_and_preserves_other_state() {
    let fx = build_fixture("template-write");
    let runtime = scoped_workspace(&fx);
    let bindings = Arc::new(super::super::project_bindings::ProjectBindings::default());
    bindings
        .register("bound", fixture_binding(&fx, Some(&fx.project_id)))
        .unwrap();
    bindings
        .register("general", fixture_binding(&fx, None))
        .unwrap();
    let conn = runtime.require_db().unwrap();
    let initial = serde_json::json!({"formData":{"gen_demand_env_require":"","gen_financial_value":"do-not-touch"},"hasSecurity":true,"techItems":[{"keep":true}],"attach1Images":[{"assetId":"preserve"}]});
    project_state::save_template_state_locked(
        &mut conn.lock().unwrap(),
        fx.project_id.clone(),
        TEMPLATE.into(),
        payload(initial.clone()),
        Some(0),
    )
    .unwrap();
    let gate = Arc::new(ApprovalGate::new(Duration::from_secs(1)));
    let records = Arc::new(Mutex::new(Vec::<ApprovalRecord>::new()));
    let capture = records.clone();
    gate.set_recorder(Arc::new(move |r| capture.lock().unwrap().push(r.clone())));
    let events = Arc::new(Mutex::new(Vec::new()));
    let emitted = events.clone();
    let handler = template_write::handler(
        runtime.clone(),
        bindings.clone(),
        gate.clone(),
        Arc::new(move |v| emitted.lock().unwrap().push(v)),
        super::super::workspace_handler(runtime.clone(), bindings.clone()),
    );
    let args = write_args(&fx.project_id);
    let post = |session: &str, a: &Value| {
        handler(
            ROUTE,
            &serde_json::json!({"sessionId":session,"callId":"call-write","originalArgs":a})
                .to_string(),
        )
    };
    let read = || {
        project_state::get_template_state_locked(&conn.lock().unwrap(), &fx.project_id, TEMPLATE)
            .unwrap()
            .unwrap()
    };
    assert_eq!(post("bound", &args).status, 403, "no grant must reject");
    assert_eq!(post("general", &args).status, 403);
    assert_eq!(post("missing", &args).status, 403);
    let mut wrong = args.clone();
    wrong["projectId"] = serde_json::json!("other-project");
    assert_eq!(post("bound", &wrong).status, 403);
    for key in [
        "total_cost_incl",
        "tax_rate",
        "summary_metrics",
        "discount_rate",
        "project_years",
        "npv",
        "techItems",
        "attach1",
        "unknown",
    ] {
        let mut bad = args.clone();
        bad["fields"] = serde_json::json!({key:"123"});
        assert!(
            template_write::prepare(&runtime, &bindings, "bound", &bad).is_err(),
            "{key}"
        );
        assert_eq!(post("bound", &bad).status, 403);
    }
    let mut wrong_type = args.clone();
    wrong_type["fields"]["gen_demand_env_require"] = serde_json::json!(123);
    assert!(template_write::arguments(&wrong_type).is_err());
    let intent = template_write::prepare(&runtime, &bindings, "bound", &args).unwrap();
    assert_eq!(intent.fields[0].previous_value, Some("".into()));
    assert!(!grant(&gate, &runtime, &bindings, args.clone(), false, None).approved);
    assert_eq!(post("bound", &args).status, 403);
    assert_eq!(read().filled_data_json, initial);
    let mut edited = args.clone();
    edited["fields"]["gen_demand_env_require"] =
        serde_json::json!("人工修订：分区实施，保留现网与回退。\n按双方清单验收。");
    assert!(
        grant(
            &gate,
            &runtime,
            &bindings,
            args.clone(),
            true,
            Some(edited.clone())
        )
        .approved
    );
    // An edited project/field collection cannot be substituted even after approval.
    assert!(template_write::validate_edit(&args, &wrong).is_err());
    let result = post("bound", &args);
    assert_eq!(result.status, 200, "{}", result.body);
    let saved = read();
    assert_eq!(
        saved.filled_data_json["formData"]["gen_demand_env_require"],
        edited["fields"]["gen_demand_env_require"]
    );
    for key in ["hasSecurity", "techItems", "attach1Images"] {
        assert_eq!(saved.filled_data_json[key], initial[key]);
    }
    assert_eq!(
        saved.filled_data_json["formData"]["gen_financial_value"],
        "do-not-touch"
    );
    assert_eq!(saved.output_config_json["outputDir"], "unchanged");
    assert_eq!(saved.field_mapping_json["keep"], "mapping");
    assert_eq!(saved.template_version, 2);
    assert_eq!(events.lock().unwrap().len(), 1);
    assert_eq!(post("bound", &args).status, 403, "one-use permission");
    let audit: Value = serde_json::from_str(&records.lock().unwrap()[1].args_json).unwrap();
    assert_eq!(audit["modelArgs"], args);
    assert_eq!(audit["userArgs"], edited);
    assert!(
        project_state::save_template_state_locked(
            &mut conn.lock().unwrap(),
            fx.project_id.clone(),
            TEMPLATE.into(),
            payload(initial.clone()),
            Some(1)
        )
        .is_err(),
        "stale UI must not overwrite approved text"
    );
    assert!(grant(&gate, &runtime, &bindings, args.clone(), true, None).approved);
    project_state::save_template_state_locked(
        &mut conn.lock().unwrap(),
        fx.project_id.clone(),
        TEMPLATE.into(),
        payload(initial.clone()),
        Some(2),
    )
    .unwrap();
    assert_eq!(
        post("bound", &args).status,
        403,
        "editing while approval pending invalidates old approval"
    );
    assert_eq!(read().filled_data_json, initial);
    let timeout = Arc::new(ApprovalGate::new(Duration::from_millis(5)));
    let d = template_write::request_approval(
        &timeout,
        &runtime,
        &bindings,
        ApprovalQuestion {
            session_id: Some("bound".into()),
            call_id: Some("late".into()),
            tool_name: TOOL.into(),
            reason: None,
            intent: None,
            args: args.clone(),
        },
        |_| {},
    );
    assert!(!d.approved);
    assert!(timeout.reviewed.take("bound", "late", TOOL, &args).is_err());
    assert!(grant(&gate, &runtime, &bindings, args.clone(), true, None).approved);
    bindings.forget("bound").unwrap();
    assert_eq!(post("bound", &args).status, 403);
}

#[test]
#[ignore = "requires real DEEPSEEK_API_KEY and provisioned plugin"]
fn template_write_real_model_uses_edited_text_and_saved_state() {
    assert!(has_api_key());
    let fx = build_fixture("template-write-real");
    let runtime = scoped_workspace(&fx);
    let db = runtime.require_db().unwrap();
    project_state::save_template_state_locked(&mut db.lock().unwrap(),fx.project_id.clone(),TEMPLATE.into(),payload(serde_json::json!({"formData":{"gen_demand_env_require":""},"techItems":[{"text":"synthetic"}]})),Some(0)).unwrap();
    let log = Arc::new(EventLog::default());
    let prompts = Arc::new(Mutex::new(Vec::new()));
    let (_bridge, acp, bindings) = launch_dsh_with_approval(
        &fx,
        Arc::new(Mutex::new(Vec::new())),
        log.clone(),
        ApprovalStance::Edit,
        prompts.clone(),
    );
    let session = acp.new_session(&repo_root_for_tests()).unwrap();
    bindings
        .register(&session, fixture_binding(&fx, Some(&fx.project_id)))
        .unwrap();
    acp.prompt(&session,&format!("请实际调用fill_template_fields，为绑定项目{}的{}写部署环境要求，字段gen_demand_env_require。正文写‘模型原始方案：一次性切换所有网络。’只写这一个字段，不要计算。",fx.project_id,TEMPLATE)).unwrap();
    let answer = await_stage2_end(&log);
    let saved =
        project_state::get_template_state_locked(&db.lock().unwrap(), &fx.project_id, TEMPLATE)
            .unwrap()
            .unwrap();
    assert_eq!(
        saved.filled_data_json["formData"]["gen_demand_env_require"],
        "人工修订：分区实施并保留原网络。\n先完成回退演练，再按清单验收。"
    );
    let raised = prompts.lock().unwrap();
    assert_eq!(raised.len(), 1);
    assert_eq!(raised[0].tool_name, TOOL);
    assert_eq!(
        raised[0].intent.as_ref().unwrap().fields[0].previous_value,
        Some("".into())
    );
    let audit = approval_log::recent(&db, 10).unwrap();
    assert_eq!(audit.len(), 1);
    let evidence = serde_json::json!({"execution":"real dsh model; synthetic project; automatic approval edit, not human Gate B","modelAnswer":answer,"savedState":saved,"approvalPrompt":raised[0],"audit":serde_json::from_str::<Value>(&audit[0].args_json).unwrap(),"humanGateB":"pending"});
    std::fs::write(
        repo_root_for_tests().join("docs/verification/template-write-b-real-evidence.json"),
        serde_json::to_string_pretty(&evidence).unwrap(),
    )
    .unwrap();
}

/// Extension gate: this test has no template-specific field names; adding catalog data must suffice.
#[test]
fn template_catalog_all_registered_templates_use_the_same_approved_transaction() {
    use super::super::{template_catalog, template_read};
    let fx = build_fixture("template-catalog");
    let runtime = scoped_workspace(&fx);
    let bindings = Arc::new(super::super::project_bindings::ProjectBindings::default());
    bindings.register("bound", fixture_binding(&fx, Some(&fx.project_id))).unwrap();
    bindings.register("general", fixture_binding(&fx, None)).unwrap();
    let conn = runtime.require_db().unwrap();
    let gate = Arc::new(ApprovalGate::new(Duration::from_secs(1)));
    let handler = template_write::handler(runtime.clone(), bindings.clone(), gate.clone(), Arc::new(|_| {}),
        super::super::workspace_handler(runtime.clone(), bindings.clone()));
    let mut evidence = serde_json::Map::new();
    for template in template_catalog::templates() {
        let name = format!("ICT项目{}{}", template.name, template.suffix);
        let mut fields = serde_json::Map::new();
        let mut initial = serde_json::json!({"formData":{"untouched":"保留字段"},"techItems":[{"keep":true}],"asset":"preserve"});
        for rule in template.fields.iter().filter(|r| r.kind == "text") {
            fields.insert(rule.key.clone(), serde_json::json!(format!("模型建议：{}分区实施。", rule.label)));
            rule.write(&mut initial, serde_json::json!("")).unwrap();
        }
        for rule in &template.fields {
            if let Some(condition) = &rule.required_when { initial[condition.field()] = match condition { super::super::template_catalog::RequiredCondition::Flag(_) => serde_json::json!(true), super::super::template_catalog::RequiredCondition::Equals(value) => value.equals.clone() }; }
        }
        let args = serde_json::json!({"projectId":fx.project_id,"templateId":name,"fields":fields});
        if template.excluded_reason.is_some() {
            assert!(template_write::arguments(&args).is_err());
            assert!(template_catalog::resolve(&name,true).is_err());
            continue;
        }
        assert!(!fields.is_empty());
        project_state::save_template_state_locked(&mut conn.lock().unwrap(),fx.project_id.clone(),name.clone(),payload(initial.clone()),Some(0)).unwrap();
        let read = |template_id: &str| template_read::handle(&runtime,&bindings,&serde_json::json!({"sessionId":"bound","templateId":template_id}).to_string());
        let before = read(&template.name);
        assert_eq!(before.status,200,"{}",before.body);
        let before: Value = serde_json::from_str(&before.body).unwrap();
        let post = |session: &str, args: &Value| handler(ROUTE,&serde_json::json!({"sessionId":session,"callId":"call-write","originalArgs":args}).to_string());
        assert_eq!(post("bound",&args).status,403,"unapproved write");
        assert_eq!(post("general",&args).status,403);
        let intent = template_write::prepare(&runtime,&bindings,"bound",&args).unwrap();
        assert!(intent.fields.iter().all(|f| f.previous_value.as_deref() == Some("")));
        for field in template.fields.iter().filter(|r| r.kind != "text") {
            let bad = serde_json::json!({"projectId":fx.project_id,"templateId":name,"fields":{&field.key:"伪造通过"}});
            assert!(template_write::prepare(&runtime,&bindings,"bound",&bad).is_err(),"{} {}",name,field.key);
            assert_eq!(post("bound",&bad).status,403);
        }
        for key in ["total_cost_incl","tax_rate","project_years","discount_rate","npv","unknown","公共字段一致","立项金额低于50万元"] {
            let bad = serde_json::json!({"projectId":fx.project_id,"templateId":name,"fields":{key:"123"}});
            assert_eq!(post("bound",&bad).status,403,"{key}");
        }
        let mut edited = args.clone();
        for (key,value) in edited["fields"].as_object_mut().unwrap() {
            *value = serde_json::json!(format!("人工修订：{key}分区实施，保留回退。\n双方按清单验收。"));
        }
        assert!(grant(&gate,&runtime,&bindings,args.clone(),true,Some(edited.clone())).approved);
        let receipt = post("bound",&args);
        assert_eq!(receipt.status,200,"{}",receipt.body);
        assert_eq!(post("bound",&args).status,403,"one-shot approval");
        let saved = project_state::get_template_state_locked(&conn.lock().unwrap(),&fx.project_id,&name).unwrap().unwrap();
        for rule in template.fields.iter().filter(|r| r.kind == "text") {
            assert_eq!(rule.raw(&saved.filled_data_json),edited["fields"][&rule.key].as_str());
            if rule.state_key.is_some() { assert!(saved.filled_data_json["formData"].get(&rule.key).is_none(),"no root shadow copies"); }
        }
        assert_eq!(saved.filled_data_json["formData"]["untouched"],"保留字段");
        assert_eq!(saved.filled_data_json["techItems"],initial["techItems"]);
        assert_eq!(saved.field_mapping_json["keep"],"mapping");
        assert_eq!(saved.template_version,2);
        let after = read(&template.name);
        assert_eq!(after.status,200,"{}",after.body);
        let after: Value = serde_json::from_str(&after.body).unwrap();
        assert_eq!(after["fields"].as_array().unwrap().len(),fields.len());
        evidence.insert(template.id.clone(),serde_json::json!({"before":before,"after":after,"savedState":saved,"approvedFields":edited["fields"]}));
    }
    std::fs::write(repo_root_for_tests().join("docs/verification/template-catalog-transaction-evidence.json"),serde_json::to_string_pretty(&evidence).unwrap()).unwrap();
}

#[test]
#[ignore = "requires real DEEPSEEK_API_KEY and provisioned plugin"]
fn template_catalog_real_model_reads_and_writes_registered_extensions() {
    assert!(has_api_key());
    let mut evidence = serde_json::Map::new();
    for template in super::super::template_catalog::templates().iter().filter(|t| t.excluded_reason.is_none() && t.id != "demand") {
        let fx = build_fixture("catalog-real");
        let runtime = scoped_workspace(&fx);
        let name = format!("ICT项目{}{}",template.name,template.suffix);
        let rule = template.fields.iter().find(|f| f.kind == "text").unwrap();
        let mut state = serde_json::json!({"formData":{}});
        rule.write(&mut state,serde_json::json!("")).unwrap();
        project_state::save_template_state_locked(&mut runtime.require_db().unwrap().lock().unwrap(),fx.project_id.clone(),name.clone(),payload(state),Some(0)).unwrap();
        let log = Arc::new(EventLog::default());
        let prompts = Arc::new(Mutex::new(Vec::new()));
        let (_bridge,acp,bindings) = launch_dsh_with_approval(&fx,Arc::new(Mutex::new(Vec::new())),log.clone(),ApprovalStance::Edit,prompts.clone());
        let session = acp.new_session(&repo_root_for_tests()).unwrap();
        bindings.register(&session,fixture_binding(&fx,Some(&fx.project_id))).unwrap();
        acp.prompt(&session,&format!("绑定项目{}。请先实际调用read_template_fields读取{}，然后实际调用fill_template_fields，只填写其中的{}。建议正文：分区实施，按清单验收。不要修改其他字段或执行测算，按工具审批结果回复。",fx.project_id,name,rule.label)).unwrap();
        let answer = await_stage2_end(&log);
        let saved = project_state::get_template_state_locked(&runtime.require_db().unwrap().lock().unwrap(),&fx.project_id,&name).unwrap().unwrap();
        assert_eq!(rule.raw(&saved.filled_data_json),Some("人工修订：分区实施并保留原网络。\n先完成回退演练，再按清单验收。"));
        let raised = prompts.lock().unwrap();
        assert_eq!(raised.len(),1);
        assert_eq!(raised[0].intent.as_ref().unwrap().fields[0].previous_value,Some("".into()));
        evidence.insert(template.id.clone(),serde_json::json!({"modelAnswer":answer,"savedState":saved,"approval":raised[0],"execution":"real model, synthetic project, automatic modified approval; not human Gate B"}));
    }
    std::fs::write(repo_root_for_tests().join("docs/verification/template-catalog-real-evidence.json"),serde_json::to_string_pretty(&evidence).unwrap()).unwrap();
}
