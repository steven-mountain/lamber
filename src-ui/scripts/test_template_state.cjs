const assert = require("node:assert/strict")
const fs = require("node:fs")
const path = require("node:path")
const vm = require("node:vm")
const ts = require("typescript")

const moduleCache = new Map()
function loadTsFile(sourcePath) {
  if (sourcePath.endsWith(".json")) return JSON.parse(fs.readFileSync(sourcePath,"utf8"));
  const normalizedPath = path.normalize(sourcePath)
  if (moduleCache.has(normalizedPath)) return moduleCache.get(normalizedPath).exports
  const source = fs.readFileSync(normalizedPath, "utf8")
  const transpiled = ts.transpileModule(source, {
    compilerOptions: {
      esModuleInterop: true,
      module: ts.ModuleKind.CommonJS,
      target: ts.ScriptTarget.ES2020,
    },
  })
  const moduleRef = { exports: {} }
  moduleCache.set(normalizedPath, moduleRef)
  const localRequire = request => {
    if (request.startsWith(".")) {
      const resolved = path.resolve(path.dirname(normalizedPath), request)
      return loadTsFile(path.extname(resolved) ? resolved : `${resolved}.ts`)
    }
    return require(request)
  }
  vm.runInNewContext(transpiled.outputText, {
    module: moduleRef,
    exports: moduleRef.exports,
    require: localRequire,
  }, { filename: normalizedPath })
  return moduleRef.exports
}

const lib = name => loadTsFile(path.join(__dirname, `../src/lib/${name}.ts`));
const generation = lib('templateGenerationState');
const demand = lib('templateCompletion/demand');
const fieldCatalog = lib('templateCompletion/catalog');
const batch = lib('selectionResultBatch');
const catalog = lib('ictSubjectCatalog');
const source = fs.readFileSync(path.join(__dirname, '../src/views/TemplateForms.tsx'), 'utf8');
const ast = ts.createSourceFile('TemplateForms.tsx', source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
function extract(name, sourceAst = ast) {
  let result;
  function visit(node) {
    if (ts.isVariableDeclaration(node) && node.name.getText(sourceAst) === name) result = node.initializer.getText(sourceAst);
    ts.forEachChild(node, visit);
  }
  visit(sourceAst); assert.ok(result, name);
  return ts.transpileModule(`globalThis.extracted = ${result}`, { compilerOptions: { target: ts.ScriptTarget.ES2020 } }).outputText;
}
const item = (excl, tax = 6) => ({ excl, tax, incl: Number((excl * (1 + tax / 100)).toFixed(2)) });
async function generate(template, saved, overrides = {}) {
  const captured = [], alerts = [];
  const projectData = { basic: { proj_name: '非默认回归项目', customer_name: '合成客户', project_years: 2 },
    cost: { it: { integration: item(100) }, ct: {}, mix: {} },
    revenue: { it: { integration: item(200) }, ct: {}, non_it_ct: item(0) } };
  const context = {
    ...generation, ...batch, ...catalog, console,
    autoSaveFormSettings: async () => true, beforeDocumentGenerate: undefined,
    isLoadingRef: { current: false }, formDataRef: { current: saved }, selectedTemplate: template,
    // Any dependence on a mounted form is a regression; four templates must work with no DOM.
    formRef: { get current() { throw new Error('DOM accessed during generation'); } },
    projectScale: 'large', selfThreeValue: '自主方案非默认', itBusMode: '专项购销', itFundSrc: '专项资金',
    revCollection: '验收90日付款', expPayment: '回款60日支付', todayStr: '2026-09-06',
    defaultSignItContent: '默认IT', defaultSignCtContent: '默认CT',
    attach1Images: [], attach2Images: [], loadDemandImages: async () => ({ attach1: [{ assetId: 'chat-asset-1' }], attach2: [{ assetId: 'chat-asset-2' }] }),
    mergeDemandImages: (legacy, assets) => [...legacy.filter(img => !img.assetId), ...assets],
    workspaceId: 'test-workspace', projectId: 'test-project', assetTargetRef: { current: { workspaceId: 'test-workspace', projectId: 'test-project', selectedTemplate: template } },
    hasPublicUrl: true, hasSecurity: true, hasMidThree: true, hasSingleSource: true,
    projectData, projectBackground: '非默认项目背景', currentSchemeStage: 'post_selection',
    selectionResultMode: 'single', preSelectionCostIt: { integration: item(120) }, selectionFeeData: {},
    preSchemeId: 'pre', preSchemeName: '甄选前', currentSchemeId: 'post', currentSchemeLabel: '甄选后',
    selectionRenewalDecisions: {}, selectionConflictAcknowledged: true,
    metrics: { npv_rate: 0.12, margin_rate: 0.3, it_npv_rate: 0.12, dynamic_payback: 1.5 },
    inqVendors: [{ vendorName: '回归供应商甲', amount: 105, taxRate: 6, remark: '非默认询价', images: [{ assetId: 'vendor-image' }] }],
    techItems: [{ serviceName: '专项服务', serviceDesc: '非默认明细', amount: 7, unit: '项' }],
    totalRevenueIncl: 212, normalizeProjectScale: value => value, joinedBusinessNames: names => names.join('、'),
    customItBusinessNames: [], customCtBusinessNames: [], customNonItCtBusinessNames: [], customMixBusinessNames: [],
    customItCostBusinessNames: [], customCtCostBusinessNames: [], customItRevenueBusinessNames: [], customCtRevenueBusinessNames: [],
    subjectItCost: 'IT成本', subjectCtCost: 'CT成本', subjectItRev: 'IT收入', subjectCtRev: 'CT收入',
    getSubjectBusinessName: () => '', itContent: '集成测试服务', ctContent: '专项平台能力',
    midThreeCode: 'TEST-CODE', midThreeName: '回归能力', procurementMethod: '其他',
    excelSubjectVariables: {}, excelSelectionFeeVariables: {}, outputDir: '',
    invoke: async (command, payload) => { assert.equal(command, 'generate_lifecycle_docs'); captured.push(payload); return '/tmp/test'; },
    alert: message => alerts.push(message), confirm: () => false,
    ...overrides,
  };
  vm.createContext(context);
  vm.runInContext(extract('toDocImagePayload'), context); context.toDocImagePayload = context.extracted;
  vm.runInContext(extract('handleGenerate'), context); const result = await context.extracted();
  return { captured, alerts, result };
}
async function main() {
  const saved = Object.fromEntries([...source.matchAll(/(?:get|getBind|getFormValue)\(['"](gen_[^'"]+)['"]/g)].map(match => [match[1], `自定义-${match[1]}`]));
  Object.assign(saved, { gen_meet_start: '2026-08-11', gen_meet_end: '2026-08-12', gen_is_advance: 'true', gen_after_approval_selection: 'true' });
  const fixtures = {};
  for (const template of ['售前方案会审纪要.docx', 'ICT项目立项签批表.docx', 'ICT项目需求导入表.docx', '甄选结果签批表.docx']) {
    const fieldPage = await generate(template, saved);
    const confirmation = await generate(template, JSON.parse(JSON.stringify(saved)));
    assert.deepEqual(fieldPage.alerts, [], template);
    assert.equal(fieldPage.captured.length, 1, template);
    assert.equal(JSON.stringify(fieldPage.captured), JSON.stringify(confirmation.captured), `${template} tab-independent input`);
    const v = fieldPage.captured[0].variables;
    assert.equal(v.MEETING_START_DATE, '2026年08月11日');
    assert.equal(v.TECH_SOLUTION, saved.gen_tech_solution);
    assert.equal(v.RISK_OWNER, saved.gen_risk_owner);
    assert.equal(v.REV_COLLECTION, '验收90日付款', 'preset-controlled state overrides duplicated gen_ value');
    assert.equal(v.IS_ADVANCE_PAYMENT, '是', 'legacy true checkbox');
    assert.equal(JSON.parse(v.TABLE_TECH_ITEMS)[0].TECH_ITEM_QTY, '7');
    assert.equal(JSON.parse(v.TABLE_INQ_VENDORS)[0].INQ_VENDOR_NAME, '回归供应商甲');
    if (template.includes('需求')) {
      assert.equal(v.BRANCH_NAME, saved.gen_demand_branch_name);
      assert.equal(v.DEMAND_SERVICE_CONTENT, saved.gen_demand_service_content);
      assert.equal(JSON.parse(v.ATTACH1_IMAGE)[0].assetId, 'chat-asset-1');
    }
    if (template.includes('立项')) { assert.equal(v.IT_CONTENT, saved.gen_sign_it_content); assert.ok(v.PROJECT_INVESTMENT_SITUATION.includes('申请立项后甄选')); }
    if (template.includes('甄选')) { assert.ok(v.WINNER_DESC.includes(saved.gen_zx_winner_name)); assert.equal(v.SELECTION_SCOPE, saved.gen_zx_scope); }
    fixtures[template] = v;
  }
  const lost = await generate('ICT项目需求导入表.docx', saved, {
    buildTemplateGenerationFields: (...args) => ({ ...generation.buildTemplateGenerationFields(...args), gen_demand_branch_name: '' }),
  });
  assert.equal(lost.captured.length, 0); assert.match(lost.alerts[0], /已阻止生成/);
  const fallback = await generate('ICT项目需求导入表.docx', saved, {
    buildTemplateGenerationFields: (...args) => ({ ...generation.buildTemplateGenerationFields(...args), gen_demand_branch_name: 'XXX分公司' }),
  });
  assert.equal(fallback.captured.length, 0);
  const missing = await generate('ICT项目需求导入表.docx', saved, { loadDemandImages: async () => ({ attach1: [{ assetId: 'missing', error: true }], attach2: [] }) });
  assert.equal(missing.captured.length, 0); assert.match(missing.alerts[0], /缺失/);
  const base = { formData: {}, techItems: [], attach1Images: [], attach2Images: [], hasPublicUrl: false, hasSecurity: false };
  const items = demand.getDemandCompletion('需求导入表.docx', base);
  assert.equal(items.length, 11);
  assert.equal(items.filter(x => x.filled).length, 9);
  const filled = demand.getDemandCompletion('需求导入表.docx', { ...base, techItems: [{}], hasPublicUrl: true, hasSecurity: true, formData: { gen_demand_public_url: 'https://example.test', gen_demand_security_detail: '测试密评' } }, [{ fieldKey: 'attach1', exists: true }, { fieldKey: 'attach2', exists: true }]);
  assert.equal(filled.filter(x => x.filled).length, 11);
  assert.equal(demand.getDemandCompletion('需求导入表.docx', base, [{ fieldKey: 'attach1', exists: false }])[9].filled, false);
  assert.equal(demand.getDemandCompletion('需求导入表.docx', { ...base, formData: { gen_demand_branch_name: '' } })[0].filled, false);
  if (process.env.LAMBER_TEMPLATE_FIXTURES) fs.writeFileSync(process.env.LAMBER_TEMPLATE_FIXTURES, JSON.stringify(fixtures, null, 2));
  const sync = lib('templateTextSync');
  const incoming = {gen_demand_env_require: '审批后的部署要求'};
  const untouched = {gen_demand_env_require:'旧值', localOnly:'未保存草稿'};
  assert.deepEqual(JSON.parse(JSON.stringify(sync.mergeApprovedText({gen_demand_env_require:'旧值'},untouched,incoming))), {merged:{...untouched,...incoming},conflicts:[]});
  assert.equal(sync.mergeApprovedText({gen_demand_env_require:'旧值'},{gen_demand_env_require:'本页修改'},incoming).conflicts.length,1);
  assert.equal(sync.mergeApprovedText({}, {}, incoming).conflicts.length,0);
  const conflict = await generate('ICT项目需求导入表.docx', saved, {autoSaveFormSettings: async () => {throw new Error('TemplateStateConflict')}});
  assert.equal(conflict.captured.length,0); assert.match(conflict.alerts[0],/TemplateStateConflict/);
  const evidencePath = path.join(__dirname,'../../docs/verification/template-write-b-real-evidence.json');
  if (fs.existsSync(evidencePath)) {
    const evidence=JSON.parse(fs.readFileSync(evidencePath,'utf8'));
    const state=evidence.savedState.filledDataJson;
    const before=demand.getDemandCompletion('ICT项目需求导入表.docx',{...state,formData:{...state.formData,gen_demand_env_require:''}}).filter(i=>i.filled).length;
    const after=demand.getDemandCompletion('ICT项目需求导入表.docx',state).filter(i=>i.filled).length;
    assert.equal(after,before+1);
    const generated=await generate('ICT项目需求导入表.docx',{...saved,...state.formData});
    assert.equal(generated.captured[0].variables.DEMAND_ENV_REQUIRE,state.formData.gen_demand_env_require);
    console.log('Real tool saved state → completion +1 → production generation variables matched approved edit.');
  }
  const catalogEvidencePath = path.join(__dirname,'../../docs/verification/template-catalog-transaction-evidence.json');
  if (fs.existsSync(catalogEvidencePath)) {
    const evidence = JSON.parse(fs.readFileSync(catalogEvidencePath,'utf8'));
    for (const [id, entry] of Object.entries(evidence)) {
      const state = entry.savedState.filledDataJson;
      const generated = await generate(entry.after.templateId,{...saved,...state.formData}, {
        ...(state.revCollection !== undefined ? {revCollection:state.revCollection} : {}),
        ...(state.expPayment !== undefined ? {expPayment:state.expPayment} : {}),
      });
      assert.equal(generated.captured.length,1);
      const variables = generated.captured[0].variables;
      if (id === 'approval') {
        assert.equal(variables.IT_CONTENT,state.formData.gen_sign_it_content);
        assert.equal(variables.CT_CONTENT,state.formData.gen_sign_ct_content);
        const before = fieldCatalog.getCatalogCompletion(entry.after.templateId,{...state,formData:{...state.formData,gen_sign_it_content:''}});
        const after = fieldCatalog.getCatalogCompletion(entry.after.templateId,state);
        assert.equal(after.filter(f=>f.filled).length,before.filter(f=>f.filled).length+1);
      }
      if (state.revCollection !== undefined) {
        assert.equal(variables.REV_COLLECTION,state.revCollection);
        assert.equal(variables.EXP_PAYMENT,state.expPayment);
      }
    }
    console.log('Catalog-approved saved states → production generation inputs and sign completion +1 passed.');
  }
  // Exercise the production event callback with independent form/root state owners and stale drafts.
  for (const conflict of [false,true]) {
    const target = {workspaceId:'w',projectId:'p',selectedTemplate:'立项签批表.docx'};
    const ctx = {
      ...fieldCatalog,...sync,useLatestCallback:fn=>fn,selectedTemplate:target.selectedTemplate,
      assetTargetRef:{current:target},savedTemplateVersion:{current:3},
      savedFormFields:{current:{gen_rev_collection:'旧收款',gen_sign_it_content:'旧IT'}},
      formDataRef:{current:{gen_sign_it_content:'旧IT',unrelated:'本地草稿'}},
      revCollection:conflict ? '本地改收款' : '旧收款',expPayment:'旧付款',
      setRevCollection:v=>ctx.revCollection=v,setExpPayment:v=>ctx.expPayment=v,
      setFormData:v=>ctx.formData=v,setSyncTrigger:()=>{},setTemplateConflict:v=>ctx.conflict=v,
    };
    vm.createContext(ctx);vm.runInContext(extract('applyApprovedTemplateText'),ctx);
    ctx.extracted({...target,templateId:target.selectedTemplate,templateVersion:4,fields:{gen_rev_collection:'批准收款',gen_sign_it_content:'批准IT'}});
    if(conflict) {assert.ok(ctx.conflict);assert.equal(ctx.savedTemplateVersion.current,3);assert.equal(ctx.formDataRef.current.gen_sign_it_content,'旧IT');}
    else {assert.equal(ctx.revCollection,'批准收款');assert.equal(ctx.formDataRef.current.gen_sign_it_content,'批准IT');assert.equal(ctx.formDataRef.current.unrelated,'本地草稿');assert.equal(ctx.formDataRef.current.gen_rev_collection,undefined);assert.equal(ctx.savedTemplateVersion.current,4);}
  }
  await testUploadBoundaries();
  console.log('Template generation: 4 production handlers × 2 state roundtrips; fields, tables, images, checkbox, input corruption and completion checks passed.');
}
module.exports = { generate, extract };
if (require.main === module) main().catch(error => { console.error(error); process.exitCode = 1; });

async function testUploadBoundaries() {
  const cardSource = fs.readFileSync(path.join(__dirname, '../src/components/ai/DemandImageCompletionCard.tsx'), 'utf8');
  const cardAst = ts.createSourceFile('Card.tsx', cardSource, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
  async function run(file, options = {}) {
    const errors = [], writes = [], receipts = [];
    const context = {
      disabled: false, uploading: { current: false }, active: { current: true },
      setBusy() {}, setError: error => errors.push(error), setPreview() {}, setRevision() {},
      target: { workspaceId: 'w1', projectId: 'p1', projectName: 'Synthetic', templateName: '需求导入表.docx', usage: 'attach1', label: '附件1' },
      workspaceService: { getState: async () => ({ currentWorkspace: { workspaceId: options.workspaceId || 'w1' } }) },
      FileReader: class { readAsDataURL() { this.result = 'data:image/png;base64,test'; this.onload(); } },
      Image: class { naturalWidth = 10; naturalHeight = 20; set src(value) { this.onload(); } },
      domainSaveService: { saveTemplateAsset: async (...args) => { writes.push(args); if (options.failWrite) throw new Error('database rejected'); return 'saved-image'; } },
      projectService: { getTemplateAssetPath: async () => '/workspace/.projects/p1/assets/saved-image.png' },
      assertDemandUploadBinding: async () => { if (options.workspaceId && options.workspaceId !== 'w1') throw new Error('工作区已切换'); if (options.unmounted) context.active.current = false; }, publishDemandAssetsChanged: async () => {}, onReceipt: text => receipts.push(text),
    };
    vm.createContext(context); vm.runInContext(extract('upload', cardAst), context);
    await context.extracted(file, { usage: 'attach1', label: '附件1' });
    assert.equal(context.uploading.current, false);
    return { errors, writes, receipts };
  }
  const png = { name: 'test.png', type: 'image/png', size: 100 };
  const huge = await run({ ...png, size: 20 * 1024 * 1024 + 1 });
  assert.equal(huge.writes.length, 0); assert.match(huge.errors.at(-1), /20MB/);
  for (const type of ['image/gif', 'image/bmp']) {
    const unsupported = await run({ ...png, type });
    assert.equal(unsupported.writes.length, 0); assert.match(unsupported.errors.at(-1), /仅支持/);
  }
  const switched = await run(png, { workspaceId: 'w2' });
  assert.equal(switched.writes.length, 0); assert.match(switched.errors.at(-1), /切换/);
  const unmounted = await run(png, { unmounted: true });
  assert.equal(unmounted.writes.length, 0, 'session switch during file decoding must cancel upload');
  const rejected = await run(png, { failWrite: true });
  assert.equal(rejected.receipts.length, 0); assert.match(rejected.errors.at(-1), /database rejected/);
  const success = await run(png);
  assert.equal(success.writes[0][0], 'p1'); assert.equal(success.writes[0][1], '需求导入表.docx');
  assert.equal(success.writes[0][2].usage, 'attach1');
  assert.match(success.receipts[0], /\.projects\/p1\/assets/);
  console.log('Upload boundaries: oversize, GIF/BMP, workspace switch, failed write, fixed target and path receipt passed.');
}
