import Decimal from 'decimal.js';
import { ICT_SUBJECT_DEFINITIONS } from './ictSubjectCatalog';
import { editIctTaxItem, serializeTaxItemForPayload, type IctSubjectState } from './ictTaxItemEdit';
import { buildIctFundingSubjectsFromInput, buildIctFundingCashflowFields } from './ictCalculationInput';
import { normalizeSubjectFundingPlans, type SubjectFundingPlans } from './ictSubjectFundingPlan';
import { exclFromIncl, restoreTaxSplitParts } from './taxAmount';
import type { IctInput } from '../utils/projectService';
export type BenefitOverride = {
    subject: string;
    inclTax: string;
    taxRate?: string;
};
export type SimulationChange = {
    kind: string;
    subject: string;
    before: string;
    after: string;
};
export type PreparedBenefitSimulation = {
    input: IctInput;
    explicitChanges: SimulationChange[];
    linkedChanges: SimulationChange[];
    taxInclAutoFix: boolean;
};
const cashflowKeys = ['rev_cashflow_excl', 'cost_cashflow_excl', 'it_rev_cashflow_excl', 'it_cost_cashflow_excl'] as const;
const clone = <T,>(v: T): T => JSON.parse(JSON.stringify(v));
const numeric = (v: unknown, label: string, scale: number, max: number) => {
    if (typeof v !== 'string' || !/^\d+(?:\.\d+)?$/.test(v))
        throw new Error(`${label}必须是无单位的非负十进制数字字符串`);
    const d = new Decimal(v);
    if (d.decimalPlaces() > scale || d.greaterThan(max))
        throw new Error(`${label}超出支持精度或范围（最多${scale}位小数，上限${max}）`);
    return d.toNumber();
};
const sameMoneyRows = (a: unknown, b: unknown) => Array.isArray(a) && Array.isArray(b) && a.length === 10 && b.length === 10 && a.every((v, i) => new Decimal(String(v)).eq(String(b[i])));
function assertCoverage(input: IctInput, plans: SubjectFundingPlans) {
    const result = buildIctFundingCashflowFields(buildIctFundingSubjectsFromInput(input), plans);
    if (!result.coverage.valid)
        throw new Error(`无法自洽重建现金流：${result.coverage.issues.map(i => i.message).join('；')}`);
    return result;
}
function financialPlan(plan: unknown) {
    if (!plan)
        return null;
    const { updatedAt: _updated, lastChangedAt: _changed, ...rest } = plan as Record<string, unknown>;
    return rest;
}
/** Runs only on a backend-supplied snapshot in main-window memory. Never receives editor state. */
export function prepareBenefitSimulation(saved: IctInput, overrides: BenefitOverride[], autoFix: boolean): PreparedBenefitSimulation {
    const input = clone(saved);
    const empty = { input, explicitChanges: [], linkedChanges: [], taxInclAutoFix: autoFix };
    if (!Array.isArray(overrides) || overrides.length > 28)
        throw new Error('overrides 必须是最多 28 个具名覆盖项');
    if (!overrides.length)
        return empty; // Exact replay, including legacy saved snapshots.
    if (input.cashflow_calculation_source !== 'subject_funding_plans' || !input.subject_funding_plans)
        throw new Error('该快照没有可重建的科目资金计划，不能将新科目金额配上旧年度现金流；请先在桌面核对并保存方案。');
    if (input.revenue_balance_rule?.enabled || input.investment_balance_rule?.enabled)
        throw new Error('该方案启用了总额锁定/差额承接，当前只读试算不能复用完整联动，请使用桌面编辑器。');
    const seen = new Set<string>();
    for (const o of overrides) {
        if (!o || Object.keys(o).some(k => !['subject', 'inclTax', 'taxRate'].includes(k)))
            throw new Error('覆盖项只接受 subject、inclTax、taxRate');
        if (seen.has(o.subject))
            throw new Error(`同一科目不能重复覆盖：${o.subject}`);
        if (!ICT_SUBJECT_DEFINITIONS.some(s => s.subjectCode === o.subject))
            throw new Error(`未知科目：${o.subject}`);
        seen.add(o.subject);
        numeric(o.inclTax, '含税金额', 2, 1e12);
        if (o.taxRate !== undefined)
            numeric(o.taxRate, '税率（百分数，如6）', 6, 100);
    }
    for (const [rev, cost] of [['rev_ct_product', 'cost_ct_other'], ['rev_ct_line', 'cost_ct_bandwidth']]) {
        if (seen.has(rev) && seen.has(cost))
            throw new Error(`覆盖 ${rev} 会联动 ${cost}，不能同时指定一对联动科目；请分别试算。`);
    }
    let state: IctSubjectState = { revIt: {}, revCt: {}, revNonItCt: { incl: 0, tax: 0, excl: 0 }, costIt: {}, costCt: {}, costMix: {} };
    const raw = input as unknown as Record<string, any>;
    for (const s of ICT_SUBJECT_DEFINITIONS) {
        const item = raw[s.subjectCode];
        const incl = numeric(item?.incl_tax, `${s.standardSubjectName}已保存金额`, 2, 1e12);
        const tax = numeric(item?.tax_rate, `${s.standardSubjectName}已保存税率`, 6, 100);
        const parts = restoreTaxSplitParts(item.split_parts, incl, tax);
        if (item.split_parts?.length && (!parts || parts.some((p, i) => !new Decimal(p.excl).eq(String(item.split_parts[i].excl_tax)))))
            throw new Error(`${s.standardSubjectName}的已保存拆分无效，拒绝静默退回单笔口径`);
        const value = { incl, tax, excl: parts ? parts.reduce((sum, p) => sum.plus(p.excl), new Decimal(0)).toNumber() : exclFromIncl(incl, tax), customSubjectName: item.custom_subject_name, billingSubjectName: item.billing_subject_name, splitParts: parts ?? undefined };
        if (s.groupId === 'revNonItCt')
            state.revNonItCt = value;
        else
            state[s.groupId][s.key] = value;
    }
    let plans = normalizeSubjectFundingPlans(input.subject_funding_plans);
    const baseline = assertCoverage(input, plans);
    for (const key of cashflowKeys)
        if (!sameMoneyRows(input[key], baseline.fields[key]))
            throw new Error(`快照的 ${key} 与科目资金计划不一致，不能重建该方案；请先在桌面核对并保存。`);
    for (const o of overrides) {
        const s = ICT_SUBJECT_DEFINITIONS.find(s => s.subjectCode === o.subject)!;
        let edit = editIctTaxItem(state, plans, s.groupId, s.key, 'incl', Number(o.inclTax), autoFix, 'manual_amount_sync', { normalizeIncl: true });
        state = edit.state;
        plans = edit.plans;
        if (o.taxRate !== undefined) {
            edit = editIctTaxItem(state, plans, s.groupId, s.key, 'tax', Number(o.taxRate), autoFix);
            state = edit.state;
            plans = edit.plans;
        }
    }
    for (const s of ICT_SUBJECT_DEFINITIONS) {
        const item = s.groupId === 'revNonItCt' ? state.revNonItCt : state[s.groupId][s.key];
        // Preserve untouched snapshot strings/metadata verbatim.
        const old = raw[s.subjectCode];
        if (seen.has(s.subjectCode) || Number(old.incl_tax) !== item.incl || Number(old.tax_rate) !== item.tax || Boolean(old.split_parts?.length) !== Boolean(item.splitParts?.length))
            raw[s.subjectCode] = serializeTaxItemForPayload(item);
    }
    input.subject_funding_plans = plans;
    const rebuilt = assertCoverage(input, plans);
    Object.assign(input, rebuilt.fields, { ignore_tail_difference: false, tail_difference_value: '0' });
    // Detect split-vs-single-line discrepancies as well as stale annual arrays.
    for (const [key, side, itOnly] of [['rev_cashflow_excl', 'revenue', false], ['cost_cashflow_excl', 'cost', false], ['it_rev_cashflow_excl', 'revenue', true], ['it_cost_cashflow_excl', 'cost', true]] as const) {
        let total = new Decimal(0);
        for (const s of ICT_SUBJECT_DEFINITIONS.filter(s => s.side === side && (!itOnly || s.groupId === 'revIt' || s.groupId === 'costIt'))) {
            const item = s.groupId === 'revNonItCt' ? state.revNonItCt : state[s.groupId][s.key];
            total = total.plus(item.excl);
        }
        const cash = input[key]!.reduce((sum, v) => sum.plus(v), new Decimal(0));
        if (!cash.eq(total))
            throw new Error(`${key}与科目不含税合计不一致（可能存在拆分尾差），拒绝输出矛盾指标。`);
    }
    const explicitChanges: SimulationChange[] = [], linkedChanges: SimulationChange[] = [];
    const before = saved as unknown as Record<string, any>;
    const change = (kind: string, subject: string, a: unknown, b: unknown, list = linkedChanges) => { if (JSON.stringify(a) !== JSON.stringify(b))
        list.push({ kind, subject, before: JSON.stringify(a ?? null), after: JSON.stringify(b ?? null) }); };
    for (const o of overrides) {
        explicitChanges.push({ kind: 'explicit_override', subject: o.subject, before: JSON.stringify({ inclTax: before[o.subject].incl_tax, taxRate: before[o.subject].tax_rate }), after: JSON.stringify({ inclTax: o.inclTax, taxRate: o.taxRate ?? before[o.subject].tax_rate }) });
    }
    for (const s of ICT_SUBJECT_DEFINITIONS) {
        const a = before[s.subjectCode], b = raw[s.subjectCode];
        const o = overrides.find(o => o.subject === s.subjectCode);
        if (!o)
            change('ct_linkage', s.subjectCode, { inclTax: a.incl_tax, taxRate: a.tax_rate }, { inclTax: b.incl_tax, taxRate: b.tax_rate });
        else if (Number(o.inclTax) !== Number(b.incl_tax))
            change('tax_normalization', s.subjectCode, o.inclTax, b.incl_tax);
        change('split_cleared', s.subjectCode, a.split_parts ?? [], b.split_parts ?? []);
    }
    for (const id of new Set([...Object.keys(saved.subject_funding_plans), ...Object.keys(plans)]))
        change('funding_plan', id, financialPlan(saved.subject_funding_plans[id]), financialPlan(plans[id]));
    for (const key of cashflowKeys)
        change('annual_cashflow', key, saved[key], input[key]);
    if (saved.ignore_tail_difference)
        change('reconciliation_reset', 'ignore_tail_difference', true, false);
    return { input, explicitChanges, linkedChanges, taxInclAutoFix: autoFix };
}
