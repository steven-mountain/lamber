import { useEffect, useRef, useState } from 'react';
import { listenBusinessEvent as listen } from '../../services/businessEvents';
import { METRIC_EPSILON } from '../../hooks/useIctCalculations';
import { ICT_SUBJECT_DEFINITIONS } from '../../lib/ictSubjectCatalog';
import { reverseMoney, reverseMetricDetail, reversePercent, type ReverseMetric } from '../../lib/structureReverseResult';
import { loadReverseProject, requestStructureReverse, selectReverseScheme, structureReceipt, type ReversePreview, type ReverseReply } from '../../services/chatStructureReverse';

export default function StructureReverseCard({ sessionId, initialMetric, initialTarget, initialScenario, disabled, onReceipt }: {
  sessionId: string; initialMetric: ReverseMetric; initialTarget: string; initialScenario: string;
  disabled: boolean; onReceipt: (content: string) => void;
}) {
  const [project, setProject] = useState<Awaited<ReturnType<typeof loadReverseProject>>>(null);
  const [scenario, setScenario] = useState(initialScenario);
  const [subject, setSubject] = useState('');
  const [metric, setMetric] = useState<ReverseMetric>(initialMetric);
  const [target, setTarget] = useState(initialTarget);
  const [preview, setPreview] = useState<ReversePreview | null>(null);
  const [result, setResult] = useState<ReverseReply | null>(null);
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);
  const busyRef = useRef(false);
  const generation = useRef(0);
  useEffect(() => {
    let active = true;
    const lifecycleGeneration = generation;
    void loadReverseProject(sessionId).then(value => { if (active) setProject(value); }).catch(reason => { if (active) setError(String(reason)); });
    const invalidate = () => { generation.current++; setSubject(''); setPreview(null); setResult(null); setTarget(initialTarget); };
    const subscriptions = [listen('lamber-workspace-state-changed', () => { invalidate(); setProject(null); }),
      listen<{ view: string }>('lamber-ai-view-changed', event => { if (event.payload.view !== 'ict_lifecycle') invalidate(); })];
    return () => { active = false; lifecycleGeneration.current++; subscriptions.forEach(promise => { void promise.then(stop => stop()).catch(console.warn); }); };
  }, [sessionId, initialTarget]);
  const reset = () => { generation.current++; setPreview(null); setResult(null); setError(''); };
  const execute = async (action: 'preview' | 'apply') => {
    if (!project || !subject || busyRef.current) return;
    busyRef.current = true; setBusy(true); setError('');
    const version = generation.current;
    try {
      const scheme = selectReverseScheme(project.project, project.schemes, scenario);
      if (action === 'apply' && (!preview || !target.trim() || !Number.isFinite(Number(target)))) throw new Error('请先读取范围，再填写并确认目标值。');
      const reply = await requestStructureReverse({ sessionId, workspaceId: project.workspaceId,
        projectId: project.projectId, projectName: project.project.name, schemeId: scheme.id,
        subjectCode: subject, metricType: metric, action,
        ...(action === 'apply' ? { target: Number(target) / 100, token: preview!.token } : {}) });
      // A completed write remains a historical receipt even if its invitation was closed.
      onReceipt(structureReceipt(project.project.name, reply));
      if (version !== generation.current) return;
      if (reply.status === 'preview') { setPreview(reply); setResult(null); }
      else { setResult(reply); if (reply.status === 'error') setPreview(null); }
    } catch (reason) {
      const message = reason instanceof Error ? reason.message : String(reason);
      onReceipt(structureReceipt(project.project.name, { status: 'error', message }));
      if (version === generation.current) setError(message);
    } finally { busyRef.current = false; setBusy(false); }
  };
  const completed = result?.status === 'success' || result?.status === 'applied_warning' ? result : null;
  const reached = completed && completed.targetReached && Number.isFinite(completed.achieved)
    && Math.abs(completed.target - completed.achieved) <= METRIC_EPSILON;
  const inputClass = 'w-full rounded-md bg-card px-3 py-2 text-body disabled:opacity-50';
  return <section className="rounded-lg bg-muted/50 p-4 text-body space-y-4" aria-label="结构反算卡片">
    <div><h3 className="font-semibold">智能结构反算</h3><p className="text-caption text-secondary-foreground">{project?.project.name || '正在核对绑定项目…'} · 确认后修改当前编辑器金额，保存仍需你操作</p></div>
    <div className="space-y-2">
      <label className="block">测算方案（阶段、名称或 ID）<input aria-label="反算方案" list="structure-reverse-schemes" className={inputClass} disabled={disabled || busy || !!completed} value={scenario} placeholder="默认方案" onChange={event => { setScenario(event.target.value); setSubject(''); reset(); }} /></label>
      <datalist id="structure-reverse-schemes">{project?.schemes.map(scheme => <option key={scheme.id} value={scheme.id}>{scheme.name} · {scheme.stage || '未标注阶段'}</option>)}</datalist>
      <label className="block">1. 由你选择反算目标科目<select aria-label="反算目标科目" className={inputClass} value={subject} disabled={!project || disabled || busy || !!completed} onChange={event => { setSubject(event.target.value); reset(); }}>
        <option value="">请选择科目，不沿用上次选择</option>{ICT_SUBJECT_DEFINITIONS.map(item => <option key={item.subjectCode} value={item.subjectCode}>{item.side === 'revenue' ? '收入' : '投入'} · {item.standardSubjectName}</option>)}
      </select></label>
      <label className="block">目标指标<select aria-label="反算目标指标" className={inputClass} value={metric} disabled={disabled || busy || !!completed} onChange={event => { setMetric(event.target.value as ReverseMetric); reset(); }}><option value="margin">毛利润率</option><option value="npv_rate">净现值率</option></select></label>
      {!completed && <button type="button" disabled={!subject || disabled || busy} className="rounded-md bg-card px-4 py-2 disabled:opacity-50" onClick={() => void execute('preview')}>{busy ? '正在核对…' : '2. 读取当前前置状态与可达范围'}</button>}
    </div>
    {preview && !completed && <div className="rounded-md bg-card p-3 space-y-3">
      <p>{preview.schemeName} · {preview.stage || '未标注阶段'} · 原快照 {preview.snapshotVersion === null ? '无' : `v${preview.snapshotVersion}`}</p>
      <p>目标科目：{preview.subjectName}<br />差额承接：{preview.balancingName}<br />影响侧：{preview.side === 'revenue' ? '收入' : '投入'}；总额锁定已启用：<span className="numeric-value">{reverseMoney(preview.totalIncl)}</span> 元</p>
      <p className="numeric-value font-semibold">当前结构下可达范围约为 {reversePercent(preview.minMetric)} – {reversePercent(preview.maxMetric)}</p>
      <p className="text-caption text-secondary-foreground">范围内未必每个值都能达到，最终执行仍会验证目标。金额变动将同步科目收付款计划。</p>
      <label className="block">3. 查看范围后确认目标值（%）<input aria-label="结构反算目标百分比" type="number" step="any" className="numeric-value w-full rounded-md bg-muted px-3 py-2" value={target} disabled={disabled || busy} placeholder="由你填写目标值" onChange={event => setTarget(event.target.value)} /></label>
      <button type="button" disabled={disabled || busy || !target.trim()} className="rounded-md bg-primary px-4 py-2 text-primary-foreground disabled:opacity-50" onClick={() => void execute('apply')}>确认执行结构反算并修改金额</button>
    </div>}
    {completed && <div className="rounded-md bg-card p-3 space-y-3">
      <p className="font-semibold">{completed.status === 'success' && reached ? '反算完成 · 当前编辑器已更新，尚未保存' : '金额已变更，结果需要核对'}</p>
      <div className="grid grid-cols-2 gap-3 numeric-value"><div className="rounded-md bg-muted p-3">目标值<br /><strong>{reverseMetricDetail(completed.target)}</strong></div><div className="rounded-md bg-muted p-3">实际达成值<br /><strong>{reverseMetricDetail(completed.achieved)}</strong></div></div>
      <p className="whitespace-pre-wrap">{completed.message}</p>
      <div className="overflow-x-auto"><table className="w-full text-caption numeric-value"><thead><tr className="bg-muted"><th className="p-2 text-left">科目（含税元）</th><th>变更前</th><th>→</th><th>变更后</th></tr></thead><tbody>{completed.changes.map(change => <tr key={change.code}><td className="p-2">{change.name}</td><td className="text-right">{reverseMoney(change.before)}</td><td className="px-2">→</td><td className="text-right">{reverseMoney(change.after)}</td></tr>)}</tbody></table></div>
      {completed.changes.map(change => <div key={change.code} className="rounded-md bg-muted/50 p-3 space-y-2">
        <p>{change.name} · 科目收付款计划</p><p className="text-caption">{change.beforePlan ? `${change.beforePlan.enabled ? '启用' : '停用'} / ${change.beforePlan.mode}` : '无计划'} → {change.afterPlan ? `${change.afterPlan.enabled ? '启用' : '停用'} / ${change.afterPlan.mode}` : '无计划'}</p>
        <table className="w-full text-caption numeric-value"><thead><tr><th className="text-left">年度</th><th>变更前</th><th>→</th><th>变更后</th></tr></thead><tbody>{Array.from({ length: 10 }, (_, year) => <tr key={year}><td>第{year + 1}年</td><td className="text-right">{change.beforePlan ? reverseMoney(change.beforePlan.annualInclValues[year]) : '无计划'}</td><td className="px-2">→</td><td className="text-right">{change.afterPlan ? reverseMoney(change.afterPlan.annualInclValues[year]) : '无计划'}</td></tr>)}</tbody></table>
      </div>)}
      {!completed.changes.length && <p>科目金额和计划均未变化。</p>}
    </div>}
    {(error || result?.status === 'error') && <p role="alert" className="whitespace-pre-wrap rounded-md bg-destructive-soft p-3 text-destructive">{error || (result?.status === 'error' ? result.message : '')}</p>}
  </section>;
}
