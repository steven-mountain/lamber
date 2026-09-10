import { ICT_SUBJECT_DEFINITIONS, getSubjectExcelDisplayName } from './ictSubjectCatalog';
import type { SubjectFundingPlan } from './ictSubjectFundingPlan';

export type ReverseMetric = 'margin' | 'npv_rate';
export interface StructureReverseOptions {
  silent?: boolean;
  previewOnly?: boolean;
  metricType?: ReverseMetric;
  target?: number;
  beforeCommit?: () => Promise<void>;
}
export type StructureReverseResult =
  | { status: 'error'; message: string }
  | { status: 'range'; minMetric: number; maxMetric: number }
  | { status: 'success'; message: string; target: number; achieved: number; metricType: ReverseMetric;
      targetReached: boolean; expectedInput: Record<string, any> };

export interface ReverseSubjectChange {
  code: string; name: string; side: 'revenue' | 'cost'; before: string; after: string;
  beforePlan: SubjectFundingPlan | null; afterPlan: SubjectFundingPlan | null;
  beforeSplit: unknown; afterSplit: unknown;
}
export const reversePercent = (value: number) => `${(value * 100).toFixed(2)}%`;
export const reverseMetricDetail = (value: number) => `${(value * 100).toFixed(4)}%`;
export const reverseMoney = (value: string | number) => Number(value).toFixed(2);

/** Read all 28 subjects from the actual before/after editor payloads. No financial transformation. */
export function structureSubjectChanges(before: Record<string, any>, after: Record<string, any>): ReverseSubjectChange[] {
  return ICT_SUBJECT_DEFINITIONS.flatMap(subject => {
    const a = before[subject.subjectCode], b = after[subject.subjectCode];
    const id = `${subject.side}:${subject.groupId}:${subject.key}`;
    const beforePlan = before.subject_funding_plans?.[id] ?? null;
    const afterPlan = after.subject_funding_plans?.[id] ?? null;
    const comparable = (plan: SubjectFundingPlan | null) => plan && ({ ...plan, updatedAt: undefined, lastChangedAt: undefined });
    if (a?.incl_tax === b?.incl_tax && JSON.stringify(comparable(beforePlan)) === JSON.stringify(comparable(afterPlan))
      && JSON.stringify(a?.split_parts) === JSON.stringify(b?.split_parts)) return [];
    return [{ code: subject.subjectCode, name: getSubjectExcelDisplayName(subject, b), side: subject.side,
      before: String(a.incl_tax), after: String(b.incl_tax), beforePlan, afterPlan,
      beforeSplit: a.split_parts ?? null, afterSplit: b.split_parts ?? null }];
  });
}

export function reverseChangesText(changes: ReverseSubjectChange[]) {
  return changes.map(change => {
    const header = `${change.name}（${change.code}）：${reverseMoney(change.before)} → ${reverseMoney(change.after)} 元`;
    const plan = (value: SubjectFundingPlan | null) => value ? `${value.enabled ? '启用' : '停用'} / ${value.mode}` : '无计划';
    const years = Array.from({ length: 10 }, (_, year) => `第${year + 1}年：${change.beforePlan ? reverseMoney(change.beforePlan.annualInclValues[year]) : '无计划'} → ${change.afterPlan ? reverseMoney(change.afterPlan.annualInclValues[year]) : '无计划'} 元`).join('\n');
    const split = JSON.stringify(change.beforeSplit) !== JSON.stringify(change.afterSplit) ? '\n原金额拆分已按既有金额更新规则变更。' : '';
    return `${header}\n科目收付款计划：${plan(change.beforePlan)} → ${plan(change.afterPlan)}\n${years}${split}`;
  }).join('\n\n');
}
