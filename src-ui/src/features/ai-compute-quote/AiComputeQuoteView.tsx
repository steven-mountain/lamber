import { useEffect, useMemo, useState } from "react";
import AppIcon from "../../components/icons/AppIcon";
import { Button } from "../../components/ui/button";
import { Input } from "../../components/ui/input";
import { ICT_SUBJECT_DEFINITIONS } from "../../lib/ictSubjectCatalog";
import { useNavigationStore } from "../../store/useNavigationStore";
import { useProjectStore } from "../../store/useProjectStore";
import { projectService, type Project } from "../../utils/projectService";
import {
  buildAiComputeQuoteOutput,
  runAiComputeQuoteSensitivity,
  summarizeQuote,
} from "./calculations";
import { normalizeQuoteFormula } from "./formulaEngine";
import QuoteFormulaCalculator from "./QuoteFormulaCalculator";
import { useAiComputeQuoteStore } from "./store";
import type {
  AiComputeQuoteExpressionFormula,
  AiComputeQuoteLineItem,
  AiComputeQuoteParameterCategory,
  AiComputeQuoteSide,
} from "./types";

const CATEGORY_LABELS: Record<AiComputeQuoteParameterCategory, string> = {
  scale: "规模",
  price: "收入价格",
  cost: "成本",
  finance: "财务",
  technical: "技术",
  custom: "自定义",
};

function createId(prefix: string) {
  const suffix = typeof crypto !== "undefined" && "randomUUID" in crypto
    ? crypto.randomUUID()
    : `${Date.now()}-${Math.random().toString(16).slice(2)}`;
  return `${prefix}-${suffix}`;
}

function toNumber(value: string) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function formatWan(value: number) {
  return (value / 10000).toLocaleString("zh-CN", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function formatNumber(value: number, digits = 2) {
  return value.toLocaleString("zh-CN", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
  });
}

export default function AiComputeQuoteView() {
  const navigateTo = useNavigationStore(state => state.navigateTo);
  const activeProjectId = useNavigationStore(state => state.activeProjectId);
  const currentProject = useProjectStore(state => state.currentProject);
  const setCurrentProject = useProjectStore(state => state.setCurrentProject);
  const store = useAiComputeQuoteStore();
  const loadQuote = store.load;
  const [projects, setProjects] = useState<Project[]>([]);
  const [showOutput, setShowOutput] = useState(false);
  const [sensitivityParameterId, setSensitivityParameterId] = useState("gpu-service-price");
  const [sensitivityMin, setSensitivityMin] = useState(72000);
  const [sensitivityMax, setSensitivityMax] = useState(108000);
  const [sensitivityStep, setSensitivityStep] = useState(9000);

  const projectId = activeProjectId || currentProject?.id || null;
  const activeProject = projects.find(project => project.id === projectId)
    || (currentProject?.id === projectId ? currentProject : null);

  useEffect(() => {
    projectService.listWorkspaceProjects()
      .then(rows => setProjects(rows.map(row => row.project)))
      .catch(() => setProjects([]));
  }, []);

  useEffect(() => {
    void loadQuote(projectId);
  }, [loadQuote, projectId]);

  const summary = useMemo(() => summarizeQuote(store.blueprint), [store.blueprint]);
  const outputPackage = useMemo(() => buildAiComputeQuoteOutput(store.blueprint), [store.blueprint]);
  const sensitivityRows = useMemo(() => runAiComputeQuoteSensitivity(store.blueprint, {
    parameterId: sensitivityParameterId,
    min: sensitivityMin,
    max: sensitivityMax,
    step: sensitivityStep,
  }), [sensitivityMax, sensitivityMin, sensitivityParameterId, sensitivityStep, store.blueprint]);

  const selectProject = (nextProjectId: string) => {
    const project = projects.find(item => item.id === nextProjectId) || null;
    setCurrentProject(project);
    navigateTo("ai_compute_quote", project?.id || null);
  };

  const saveAsBlueprint = async () => {
    const name = window.prompt("请输入蓝图名称", store.blueprint.name);
    if (!name?.trim()) return;
    store.renameBlueprint(name);
    await store.save(projectId);
  };

  const requestIctOutput = () => {
    const confirmed = window.confirm(
      "本轮仅生成智算输出包预览，不会写入 ICT 正式科目。请在查看后进入 ICT 生命周期模块手工确认科目和资金计划。"
    );
    if (confirmed) setShowOutput(true);
  };

  return (
    <div className="flex h-full min-h-0 flex-1 flex-col overflow-hidden bg-background text-foreground">
      <header className="flex h-16 shrink-0 items-center gap-4 bg-card px-6 shadow-sm">
        <Button variant="ghost" onClick={() => navigateTo("hub")}>← 返回集市</Button>
        <div>
          <h1 className="text-page-title font-bold">智算报价测算</h1>
          <p className="text-caption text-secondary-foreground">智算报价蓝图与 ICT 科目输出预览</p>
        </div>
        <div className="ml-auto flex items-center gap-2">
          <Button variant="secondary" onClick={store.resetToH200}>
            <AppIcon name="reverse" />恢复 H200
          </Button>
          <Button variant="outline" disabled={!projectId || store.isSaving} onClick={() => void store.save(projectId)}>
            <AppIcon name="save" />{store.isSaving ? "保存中..." : "保存项目"}
          </Button>
          <Button onClick={() => void saveAsBlueprint()}>
            <AppIcon name="presets" />保存为蓝图
          </Button>
        </div>
      </header>

      <main className="min-h-0 flex-1 overflow-y-auto p-5">
        <section className="mb-5 rounded-xl bg-card p-4 shadow-sm">
          <div className="flex flex-wrap items-center gap-3">
            <div className="min-w-[220px]">
              <div className="mb-1 text-caption font-semibold text-secondary-foreground">当前项目</div>
              <select
                className="h-[var(--density-input-height)] w-full rounded-md border border-input bg-card px-3 text-sm"
                value={projectId || ""}
                onChange={event => selectProject(event.target.value)}
              >
                <option value="">自由预览（不持久化）</option>
                {projects.map(project => <option key={project.id} value={project.id}>{project.name}</option>)}
              </select>
            </div>
            <StatusItem label="当前蓝图" value={store.blueprint.name} />
            <StatusItem label="ICT 状态" value={activeProject?.benefit_status === "normal" ? "已有正式测算" : "待配置资金计划"} />
            <StatusItem label="资金计划来源" value="ICT 生命周期模块" />
            <StatusItem label="输出状态" value="预览中" tone="primary" />
            <div className="ml-auto flex flex-wrap gap-2">
              <Button variant="outline" onClick={() => navigateTo("ict_lifecycle", projectId)}>
                打开 ICT 测算
              </Button>
              <Button variant="secondary" onClick={() => setShowOutput(true)}>查看输出包</Button>
              <Button onClick={requestIctOutput}>输出到 ICT</Button>
            </div>
          </div>
          <div className="mt-3 rounded-lg bg-primary-soft px-3 py-2 text-caption text-primary">
            当前为智算预览测算，结果基于智算收入/成本输出包。资金计划请在 ICT 生命周期模块中维护。
          </div>
          {store.error && <div className="mt-3 rounded-lg bg-destructive-soft px-3 py-2 text-caption text-destructive">{store.error}</div>}
          <div className="mt-2 text-caption text-secondary-foreground">
            {store.isDirty ? "当前蓝图有未保存修改" : store.lastSavedAt ? `已保存：${new Date(store.lastSavedAt).toLocaleString("zh-CN")}` : "当前项目尚未保存智算蓝图"}
          </div>
        </section>

        <div className="grid items-start gap-5 xl:grid-cols-[330px_minmax(620px,1fr)_300px]">
          <ParameterPanel />
          <div className="space-y-5">
            <LineItemPanel side="revenue" title="收入计算项" />
            <LineItemPanel side="cost" title="成本计算项" />
          </div>
          <ResultPreview summary={summary} outputCount={outputPackage.length} />
        </div>

        <section className="mt-5 rounded-xl bg-card p-5 shadow-sm">
          <div className="mb-4 flex flex-wrap items-end gap-3">
            <div className="mr-auto">
              <h2 className="text-section-title">单变量敏感性分析</h2>
              <p className="text-caption text-secondary-foreground">仅创建临时计算快照，不修改当前蓝图参数。</p>
            </div>
            <Field label="分析参数">
              <select
                className="h-[var(--density-input-height)] min-w-[190px] rounded-md border border-input bg-card px-3 text-sm"
                value={sensitivityParameterId}
                onChange={event => setSensitivityParameterId(event.target.value)}
              >
                {store.blueprint.parameters.filter(parameter => parameter.sensitivityEnabled !== false).map(parameter => (
                  <option key={parameter.id} value={parameter.id}>{parameter.name}</option>
                ))}
              </select>
            </Field>
            <Field label="最小值"><Input className="w-28 numeric-value" type="number" value={sensitivityMin} onChange={event => setSensitivityMin(toNumber(event.target.value))} /></Field>
            <Field label="最大值"><Input className="w-28 numeric-value" type="number" value={sensitivityMax} onChange={event => setSensitivityMax(toNumber(event.target.value))} /></Field>
            <Field label="步长"><Input className="w-28 numeric-value" type="number" min="0" value={sensitivityStep} onChange={event => setSensitivityStep(toNumber(event.target.value))} /></Field>
          </div>
          {sensitivityRows.length > 0 ? (
            <div className="overflow-x-auto rounded-lg bg-muted/40">
              <table className="w-full text-left text-sm">
                <thead className="text-caption text-secondary-foreground">
                  <tr>
                    <th className="px-4">参数值</th><th className="px-4">总收入（万元）</th><th className="px-4">总成本（万元）</th>
                    <th className="px-4">毛利额（万元）</th><th className="px-4">毛利率</th><th className="px-4">每台每月成本（元）</th>
                  </tr>
                </thead>
                <tbody>
                  {sensitivityRows.map(row => (
                    <tr key={row.parameterValue} className="odd:bg-card/70">
                      <td className="px-4 numeric-value">{formatNumber(row.parameterValue, 4)}</td>
                      <td className="px-4 numeric-value">{formatWan(row.totalRevenue)}</td>
                      <td className="px-4 numeric-value">{formatWan(row.totalCost)}</td>
                      <td className="px-4 numeric-value">{formatWan(row.grossProfit)}</td>
                      <td className="px-4 numeric-value">{formatNumber(row.grossMarginRate)}%</td>
                      <td className="px-4 numeric-value">{formatNumber(row.costPerDeviceMonth)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <div className="rounded-lg bg-warning-soft p-4 text-sm text-warning-foreground">请检查范围和步长，最多生成 500 组结果。</div>
          )}
        </section>
      </main>

      {showOutput && <OutputPackageModal onClose={() => setShowOutput(false)} />}
    </div>
  );
}

function StatusItem({ label, value, tone }: { label: string; value: string; tone?: "primary" }) {
  return (
    <div className="rounded-lg bg-muted/60 px-3 py-2">
      <div className="text-[10px] font-semibold uppercase tracking-wide text-secondary-foreground">{label}</div>
      <div className={`text-label font-bold ${tone === "primary" ? "text-primary" : "text-foreground"}`}>{value}</div>
    </div>
  );
}

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return <label className="grid gap-1 text-caption font-semibold text-secondary-foreground">{label}{children}</label>;
}

function ParameterPanel() {
  const store = useAiComputeQuoteStore();

  const addParameter = () => {
    const index = store.blueprint.parameters.length + 1;
    store.addParameter({
      id: createId("parameter"),
      name: `新参数 ${index}`,
      key: `custom_parameter_${index}`,
      value: 0,
      unit: "",
      category: "custom",
      sensitivityEnabled: true,
    });
  };

  return (
    <section className="rounded-xl bg-card p-4 shadow-sm">
      <div className="mb-3 flex items-center justify-between">
        <div>
          <h2 className="text-section-title">参数区</h2>
          <p className="text-caption text-secondary-foreground">{store.blueprint.parameters.length} 个参数，修改后实时重算</p>
        </div>
        <Button size="sm" onClick={addParameter}>新增</Button>
      </div>
      <div className="space-y-3">
        {store.blueprint.parameters.map(parameter => (
          <div key={parameter.id} className="rounded-lg bg-muted/45 p-3">
            <div className="grid grid-cols-[1fr_92px] gap-2">
              <Input value={parameter.name} onChange={event => store.updateParameter(parameter.id, { name: event.target.value })} />
              <Input className="numeric-value" type="number" value={parameter.value} onChange={event => store.updateParameter(parameter.id, { value: toNumber(event.target.value) })} />
              <Input value={parameter.key} onChange={event => store.updateParameter(parameter.id, { key: event.target.value })} />
              <Input value={parameter.unit || ""} placeholder="单位" onChange={event => store.updateParameter(parameter.id, { unit: event.target.value })} />
              <select
                className="h-[var(--density-input-height)] rounded-md border border-input bg-card px-2 text-sm"
                value={parameter.category || "custom"}
                onChange={event => store.updateParameter(parameter.id, { category: event.target.value as AiComputeQuoteParameterCategory })}
              >
                {Object.entries(CATEGORY_LABELS).map(([value, label]) => <option key={value} value={value}>{label}</option>)}
              </select>
              <label className="flex items-center gap-2 px-1 text-caption text-secondary-foreground">
                <input type="checkbox" checked={parameter.sensitivityEnabled !== false} onChange={event => store.updateParameter(parameter.id, { sensitivityEnabled: event.target.checked })} />
                敏感性
              </label>
            </div>
            <div className="mt-2 flex justify-end gap-1">
              <Button variant="ghost" size="sm" onClick={() => store.duplicateParameter(parameter.id, createId("parameter"))}><AppIcon name="copy" />复制</Button>
              <Button variant="ghost" size="sm" disabled={parameter.locked} onClick={() => store.removeParameter(parameter.id)}><AppIcon name="delete" />删除</Button>
            </div>
          </div>
        ))}
      </div>
    </section>
  );
}

function LineItemPanel({ side, title }: { side: AiComputeQuoteSide; title: string }) {
  const store = useAiComputeQuoteStore();
  const items = side === "revenue" ? store.blueprint.revenueItems : store.blueprint.costItems;
  const total = items.reduce((sum, item) => sum + item.amountInclTax, 0);

  const addItem = () => {
    const firstParameter = store.blueprint.parameters[0];
    const item: AiComputeQuoteLineItem = {
      id: createId(side),
      side,
      name: side === "revenue" ? "新增收入项" : "新增成本项",
      formula: {
        version: 2,
        tokens: firstParameter
          ? [{ type: "parameter", id: firstParameter.id, name: firstParameter.name }]
          : [{ type: "constant", value: 0 }],
      },
      amountInclTax: 0,
      amountExclTax: 0,
      taxRate: 6,
      enabled: true,
      outputEnabled: true,
    };
    store.addLineItem(item);
  };

  return (
    <section className="rounded-xl bg-card p-4 shadow-sm">
      <div className="mb-3 flex items-center justify-between">
        <div>
          <h2 className="text-section-title">{title}</h2>
          <p className="text-caption text-secondary-foreground">含税合计 <span className="numeric-value font-bold text-foreground">{formatWan(total)} 万元</span></p>
        </div>
        <Button size="sm" onClick={addItem}>新增{side === "revenue" ? "收入" : "成本"}</Button>
      </div>
      <div className="space-y-3">
        {items.map(item => <LineItemEditor key={item.id} item={item} />)}
      </div>
    </section>
  );
}

function LineItemEditor({ item }: { item: AiComputeQuoteLineItem }) {
  const store = useAiComputeQuoteStore();
  const [expanded, setExpanded] = useState(false);
  const mapping = store.blueprint.mappings.find(candidate => candidate.lineItemId === item.id);
  const subjects = ICT_SUBJECT_DEFINITIONS.filter(subject => subject.side === item.side);
  const normalizedFormula = normalizeQuoteFormula(item.formula, store.blueprint.parameters);
  const setFormula = (formula: AiComputeQuoteExpressionFormula) => store.updateFormula(item.side, item.id, formula);
  const selectMapping = (subjectCode: string) => {
    const subject = subjects.find(candidate => candidate.subjectCode === subjectCode);
    if (!subject) {
      store.updateMapping(item.id, null);
      return;
    }
    store.updateMapping(item.id, {
      id: mapping?.id || createId("mapping"),
      lineItemId: item.id,
      side: item.side,
      ictSubjectCode: subject.subjectCode,
      ictSubjectName: subject.standardSubjectName,
      enabled: true,
    });
  };

  return (
    <article className="rounded-lg bg-muted/45 p-3">
      <div className="grid gap-2 lg:grid-cols-[auto_minmax(150px,1fr)_94px_120px_120px_auto] lg:items-center">
        <input type="checkbox" checked={item.enabled} onChange={event => store.updateLineItem(item.side, item.id, { enabled: event.target.checked })} title="启用" />
        <Input value={item.name} onChange={event => store.updateLineItem(item.side, item.id, { name: event.target.value })} />
        <label className="flex items-center gap-2 text-caption text-secondary-foreground">
          税率<Input className="w-16 numeric-value" type="number" min="0" value={item.taxRate || 0} onChange={event => store.updateLineItem(item.side, item.id, { taxRate: toNumber(event.target.value) })} />
        </label>
        <div className="text-right">
          <div className="text-[10px] text-secondary-foreground">含税（万元）</div>
          <div className="numeric-value text-body-strong">{formatWan(item.amountInclTax)}</div>
        </div>
        <div className="text-right">
          <div className="text-[10px] text-secondary-foreground">不含税（万元）</div>
          <div className="numeric-value text-body-strong">{formatWan(item.amountExclTax)}</div>
        </div>
        <Button variant="ghost" size="sm" onClick={() => store.removeLineItem(item.side, item.id)}><AppIcon name="delete" /></Button>
      </div>

      <div className="mt-3 grid gap-2 md:grid-cols-[auto_minmax(180px,1fr)_auto] md:items-center">
        <label className="flex items-center gap-2 text-caption font-semibold text-secondary-foreground">
          <input type="checkbox" checked={item.outputEnabled !== false} onChange={event => store.updateLineItem(item.side, item.id, { outputEnabled: event.target.checked })} />
          输出到 ICT 科目包
        </label>
        <select
          className="h-[var(--density-input-height)] rounded-md border border-input bg-card px-3 text-sm"
          value={mapping?.ictSubjectCode || ""}
          onChange={event => selectMapping(event.target.value)}
        >
          <option value="">未映射</option>
          {subjects.map(subject => <option key={subject.subjectCode} value={subject.subjectCode}>{subject.standardSubjectName}</option>)}
        </select>
        <Button variant="ghost" size="sm" onClick={() => setExpanded(value => !value)}>
          {expanded ? "收起计算过程" : "展开计算过程"}
          <AppIcon name={expanded ? "chevronUp" : "chevronDown"} />
        </Button>
      </div>

      {item.calculationStatus !== "valid" && !expanded && (
        <div className={`mt-2 rounded-md px-3 py-2 text-caption font-semibold ${
          item.calculationStatus === "error"
            ? "bg-destructive-soft text-destructive"
            : "bg-warning-soft text-warning-foreground"
        }`}>
          {item.calculationError || "公式不完整"}
        </div>
      )}

      {expanded && (
        <div className="mt-3">
          <QuoteFormulaCalculator
            blueprint={store.blueprint}
            item={{ ...item, formula: normalizedFormula }}
            onChange={setFormula}
          />
        </div>
      )}
    </article>
  );
}

function ResultPreview({ summary, outputCount }: { summary: ReturnType<typeof summarizeQuote>; outputCount: number }) {
  return (
    <aside className="sticky top-0 rounded-xl bg-card p-4 shadow-sm">
      <h2 className="text-section-title">效益预览</h2>
      <p className="mb-4 text-caption text-secondary-foreground">本地报价汇总，金额为含税口径</p>
      <div className="space-y-2">
        <Metric label="总收入" value={`${formatWan(summary.totalRevenue)} 万元`} />
        <Metric label="总成本" value={`${formatWan(summary.totalCost)} 万元`} />
        <Metric label="毛利额" value={`${formatWan(summary.grossProfit)} 万元`} tone={summary.grossProfit >= 0 ? "success" : "danger"} />
        <Metric label="毛利率" value={`${formatNumber(summary.grossMarginRate)}%`} tone={summary.grossMarginRate >= 0 ? "success" : "danger"} />
        <Metric label="每台每月成本" value={`${formatNumber(summary.costPerDeviceMonth)} 元`} />
        <Metric label="输出 ICT 科目" value={`${outputCount} 项`} />
      </div>
      <div className="mt-4 rounded-lg bg-warning-soft p-3 text-caption text-warning-foreground">
        NPV、净现值率、IRR 和回收期需关联 ICT 生命周期测算并配置科目资金计划后查看。当前预览不写入 ICT 正式数据。
      </div>
    </aside>
  );
}

function Metric({ label, value, tone }: { label: string; value: string; tone?: "success" | "danger" }) {
  const toneClass = tone === "success" ? "text-success-foreground" : tone === "danger" ? "text-destructive" : "text-foreground";
  return (
    <div className="rounded-lg bg-muted/55 px-3 py-3">
      <div className="text-caption text-secondary-foreground">{label}</div>
      <div className={`numeric-value text-metric ${toneClass}`}>{value}</div>
    </div>
  );
}

function OutputPackageModal({ onClose }: { onClose: () => void }) {
  const blueprint = useAiComputeQuoteStore(state => state.blueprint);
  const output = buildAiComputeQuoteOutput(blueprint);
  const lineItems = [...blueprint.revenueItems, ...blueprint.costItems];

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-foreground/25 p-6 backdrop-blur-sm">
      <div className="max-h-[86vh] w-full max-w-5xl overflow-y-auto rounded-xl bg-card p-6 shadow-xl">
        <div className="mb-4 flex items-start justify-between gap-4">
          <div>
            <h2 className="text-page-title">智算输出包</h2>
            <p className="text-body text-secondary-foreground">按 ICT 科目合并金额。此处仅预览，不写入正式测算。</p>
          </div>
          <Button variant="ghost" size="icon" onClick={onClose}><AppIcon name="close" /></Button>
        </div>
        <div className="overflow-x-auto rounded-lg bg-muted/40">
          <table className="w-full text-left text-sm">
            <thead className="text-caption text-secondary-foreground">
              <tr><th className="px-4">类型</th><th className="px-4">ICT 科目</th><th className="px-4">含税（万元）</th><th className="px-4">不含税（万元）</th><th className="px-4">来源</th></tr>
            </thead>
            <tbody>
              {output.map(item => (
                <tr key={`${item.side}-${item.ictSubjectCode}`} className="odd:bg-card/70">
                  <td className="px-4">{item.side === "revenue" ? "收入" : "成本"}</td>
                  <td className="px-4 font-semibold">{item.ictSubjectName}</td>
                  <td className="px-4 numeric-value">{formatWan(item.amountInclTax)}</td>
                  <td className="px-4 numeric-value">{formatWan(item.amountExclTax)}</td>
                  <td className="px-4 text-caption text-secondary-foreground">
                    {item.sourceLineItemIds.map(id => lineItems.find(lineItem => lineItem.id === id)?.name || id).join("、")}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        {output.length === 0 && <div className="rounded-lg bg-warning-soft p-4 text-warning-foreground">没有已启用且完成科目映射的输出项。</div>}
        <div className="mt-5 flex justify-end gap-2">
          <Button variant="outline" onClick={onClose}>关闭</Button>
          <Button onClick={() => {
            window.alert("安全边界：本阶段不自动覆盖 ICT 科目。请打开 ICT 测算后手工确认科目金额和资金计划。");
          }}>确认输出边界</Button>
        </div>
      </div>
    </div>
  );
}
