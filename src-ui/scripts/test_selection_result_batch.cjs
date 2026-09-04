const assert = require("node:assert/strict")
const fs = require("node:fs")
const path = require("node:path")
const vm = require("node:vm")
const ts = require("typescript")

const moduleCache = new Map()
function loadTsFile(sourcePath) {
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

const batch = loadTsFile(path.join(__dirname, "../src/lib/selectionResultBatch.ts"))
const selectionFee = loadTsFile(path.join(__dirname, "../src/lib/selectionFee.ts"))

const item = (excl, tax) => ({ excl, tax, incl: Number((excl * (1 + tax / 100)).toFixed(2)) })
const emptyProjectData = (name, limit = "", targetSubjectCode = "") => ({
  basic: { proj_name: name, customer_name: "客户", project_years: 1 },
  cost: { it: {}, ct: {}, mix: {} },
  revenue: { it: {}, ct: {}, non_it_ct: item(0, 9) },
  selectionFee: { limit, targetSubjectCode },
})
const shared = overrides => ({
  winnerName: "供应商A",
  scope: "三级库",
  industry: "/",
  method: "竞争性甄选",
  rule: "标准方案",
  standardPlan: "竞价法",
  revenueCollection: "验收后付款",
  expenditurePayment: "回款后付款",
  ...overrides,
})
const project = (id, name, projectData, preSelectionCostIt, overrides = {}) => ({
  projectId: id,
  projectName: name,
  customerName: "客户",
  postSchemeId: `${id}-post`,
  postSchemeName: "甄选后",
  preSchemeId: `${id}-pre`,
  preSchemeName: "甄选前",
  projectData,
  preSelectionCostIt,
  metrics: { npv_rate: 0.09, margin_rate: 0.08, it_npv_rate: 0.07 },
  projectBackground: "背景",
  sharedFields: shared(overrides),
  validationErrors: [],
  missingMetrics: [],
})

const p1Data = emptyProjectData("项目一")
p1Data.cost.it.integration = item(100, 6)
p1Data.cost.ct.construction = item(50, 9)
p1Data.cost.ct.other = item(20, 6)
p1Data.cost.ct.renewal = item(10, 9)
p1Data.revenue.it.integration = item(150, 6)
p1Data.revenue.ct.line = item(30, 9)

const p2Data = emptyProjectData("项目二", "250", "cost_it_construction")
p2Data.cost.it.integration = item(200, 6)
p2Data.cost.it.construction = item(0, 9)
p2Data.cost.ct.bandwidth = item(40, 9)
p2Data.cost.ct.other = item(30, 6)
p2Data.revenue.it.integration = item(260, 6)
p2Data.revenue.ct.line = item(40, 9)

const projects = [
  project("p1", "项目一", p1Data, { integration: item(120, 6) }),
  project("p2", "项目二", p2Data, { integration: item(230, 6) }),
]

const model = batch.buildSelectionResultBatchModel(projects, { p1: "include" })
assert.equal(model.totalLimitExcl.toFixed(2), "370.00", "手填限价优先，其余取甄选前 IT")
assert.equal(model.totalWinnerExcl.toFixed(2), "300.00", "中选金额只汇总甄选后 IT")
assert.equal(model.totalCostExcl.toFixed(2), "450.00", "投入表汇总全部投入科目")
assert.equal(model.totalRevenueExcl.toFixed(2), "480.00", "收入表必须汇总 IT 与 CT，不能复制投入合计")
assert.equal(model.approvalAmountExcl.toFixed(2), "400.00", "立项金额=IT+CT专线+确认计入的续签")
assert.equal(model.tableE.length, 2, "效益表每个项目一行")
assert.equal(JSON.stringify(model.tableB.map(row => row.B_SEQ)), JSON.stringify(["1", "2"]), "跨表项目序号稳定")
assert.equal(model.tableA[1].A_FEE_TYPE, "IT-施工", "手填限价应使用所选投入科目名称")
assert.equal(model.tableA[1].A_TAX_RATE, "9%", "手填限价应使用所选投入科目税率")

const writeAmounts = selectionFee.calculateSelectionFeeWriteAmounts("89000", "400")
assert.equal(writeAmounts.valid, true)
assert.equal(writeAmounts.targetIncl, 88600, "目标科目=最高限价-供应商承担的甄选服务费")
assert.equal(writeAmounts.serviceFeeIncl, 400)
assert.equal(writeAmounts.targetIncl + writeAmounts.serviceFeeIncl, writeAmounts.limitIncl, "两科目合计必须等于最高限价")
assert.equal(selectionFee.normalizeSelectionFeeTargetSubjectCode("missing"), "cost_it_integration", "旧方案或无效科目回退集成服务")
assert.equal(selectionFee.SELECTION_FEE_TARGET_SUBJECTS.some(subject => subject.subjectCode === "cost_it_bidding"), false, "中标服务费为系统写入科目，不可作为目标科目")

const excludedRenewal = batch.buildSelectionResultBatchModel(projects, { p1: "exclude" })
assert.equal(excludedRenewal.approvalAmountExcl.toFixed(2), "390.00", "其他产品续签不得计入立项金额")

const conflicts = batch.detectSelectionSharedConflicts([
  projects[0],
  { ...projects[1], sharedFields: shared({ winnerName: "供应商B", scope: "二级库" }) },
])
assert.equal(conflicts.find(conflict => conflict.key === "winnerName").blocking, true)
assert.equal(conflicts.find(conflict => conflict.key === "scope").blocking, false)

assert.equal(batch.defaultSelectionBatchName(projects), "项目一等2个ICT项目")
assert.equal(
  batch.defaultSelectionBatchName([{ ...projects[0], projectName: "沙坪坝消防检测项目" }, projects[1]]),
  "沙坪坝消防检测等2个ICT项目",
  "默认批次名只移除项目名末尾的“项目”，不破坏名称正文",
)
assert.equal(batch.calculateSelectionApprovalAmount(p1Data, "include").toFixed(2), "160.00")
assert.equal(batch.calculateSelectionApprovalAmount(p1Data, "exclude").toFixed(2), "150.00")

console.log("selection result batch: 全部测试通过 ✅")
