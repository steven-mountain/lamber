import { exclFromIncl, inclFromExcl, serializeTaxSplitParts, type TaxSplitPart } from "./taxAmount";
import { ICT_SUBJECT_DEFINITIONS, normalizeCustomSubjectName, type IctSubjectGroupId } from "./ictSubjectCatalog";
import { initializeMissingSubjectFundingPlans, syncSubjectFundingPlansToAmounts, type SubjectFundingPlans, type SubjectFundingSubjectRef, type SubjectFundingPlanLastChangeReason } from "./ictSubjectFundingPlan";
export interface TaxItem {
    incl: number;
    tax: number;
    excl: number;
    customSubjectName?: string;
    billingSubjectName?: string;
    splitParts?: TaxSplitPart[];
}
export type IctSubjectState = {
    revIt: Record<string, TaxItem>;
    revCt: Record<string, TaxItem>;
    revNonItCt: TaxItem;
    costIt: Record<string, TaxItem>;
    costCt: Record<string, TaxItem>;
    costMix: Record<string, TaxItem>;
};
export function positiveFundingSubjects(state: IctSubjectState) {
    return ICT_SUBJECT_DEFINITIONS.flatMap(s => {
        const item = s.groupId === "revNonItCt" ? state.revNonItCt : state[s.groupId][s.key];
        return item.incl > 0 ? [{ subjectRef: { side: s.side, groupId: s.groupId, key: s.key }, amountIncl: item.incl }] : [];
    });
}
/** The desktop's single-item edit, including CT linkage and plan synchronization.
 * Pure: no React state, persistence, navigation, or preferences access.
 */
export function editIctTaxItem(state: IctSubjectState, plans: SubjectFundingPlans, groupId: IctSubjectGroupId, key: string, field: "incl" | "tax" | "excl", val: number, autoFix: boolean, reason?: SubjectFundingPlanLastChangeReason, options?: {
    normalizeIncl?: boolean;
}) {
    const { revIt, revCt, revNonItCt, costIt, costCt, costMix } = state;
    const next = { ...state };
    const setRevIt = (v: typeof revIt) => { next.revIt = v; };
    const setRevCt = (v: typeof revCt) => { next.revCt = v; };
    const setRevNonItCt = (v: typeof revNonItCt) => { next.revNonItCt = v; };
    const setCostIt = (v: typeof costIt) => { next.costIt = v; };
    const setCostCt = (v: typeof costCt) => { next.costCt = v; };
    const setCostMix = (v: typeof costMix) => { next.costMix = v; };
    // Collect effective incl amounts for funding plan sync.
    // processItem returns the resolved incl value so we can sync plans afterwards.
    // 财务口径（不含税为锚）：编辑不含税时含税取反推值；含税是否被改写
    // 取决于「财务口径自动修正」开关（normalizeIncl 为显式请求，仅在开关开启时使用），
    // 关闭时保留录入含税，由界面提示与生成前校验兜底。
    const shouldNormalizeIncl = (explicit?: boolean) => (explicit || false) && autoFix;
    const processItem = (groupState: any, setGroupState: any, targetKey: string): number => {
        // 金额/税率一经编辑，既有拆分明细即失效，回到普通单笔口径。
        const item = { ...groupState[targetKey], [field]: isNaN(val) ? 0 : val, splitParts: undefined };
        if (field === 'incl' || field === 'tax') {
            item.excl = exclFromIncl(item.incl, item.tax);
            if (shouldNormalizeIncl(field === 'tax' || options?.normalizeIncl)) {
                item.incl = inclFromExcl(item.excl, item.tax);
            }
        }
        else if (field === 'excl') {
            item.incl = inclFromExcl(item.excl, item.tax);
        }
        setGroupState({ ...groupState, [targetKey]: item });
        return Number(item.incl) || 0;
    };
    // Track incl amounts for subjects that need plan sync
    const syncUpdates: Array<{
        subjectRef: SubjectFundingSubjectRef;
        newAmountIncl: number;
        reason?: SubjectFundingPlanLastChangeReason;
    }> = [];
    const needsSync = field !== "tax";
    const sideForGroup = (gid: string) => (gid === "revIt" || gid === "revCt" || gid === "revNonItCt") ? "revenue" as const : "cost" as const;
    const trackSync = (gid: string, k: string, effectiveIncl: number, overrideReason?: SubjectFundingPlanLastChangeReason) => {
        if (!needsSync)
            return;
        syncUpdates.push({
            subjectRef: { side: sideForGroup(gid), groupId: gid as SubjectFundingSubjectRef["groupId"], key: k },
            newAmountIncl: effectiveIncl,
            reason: overrideReason || reason,
        });
    };
    if (groupId === 'revIt') {
        trackSync(groupId, key, processItem(revIt, setRevIt, key));
    }
    else if (groupId === 'revCt') {
        const effectiveIncl = processItem(revCt, setRevCt, key);
        trackSync(groupId, key, effectiveIncl);
        if (key === 'product') {
            processItem(costCt, setCostCt, 'other');
            trackSync('costCt', 'other', effectiveIncl, "ct_linkage_sync");
        }
        if (key === 'line') {
            processItem(costCt, setCostCt, 'bandwidth');
            trackSync('costCt', 'bandwidth', effectiveIncl, "ct_linkage_sync");
        }
    }
    else if (groupId === 'revNonItCt') {
        const item = { ...revNonItCt, [field]: isNaN(val) ? 0 : val, splitParts: undefined };
        if (field === 'incl' || field === 'tax') {
            item.excl = exclFromIncl(item.incl, item.tax);
            if (shouldNormalizeIncl(field === 'tax' || options?.normalizeIncl)) {
                item.incl = inclFromExcl(item.excl, item.tax);
            }
        }
        else if (field === 'excl') {
            item.incl = inclFromExcl(item.excl, item.tax);
        }
        setRevNonItCt(item);
        trackSync(groupId, key, Number(item.incl) || 0);
    }
    else if (groupId === 'costIt') {
        trackSync(groupId, key, processItem(costIt, setCostIt, key));
    }
    else if (groupId === 'costCt') {
        trackSync(groupId, key, processItem(costCt, setCostCt, key));
    }
    else if (groupId === 'costMix') {
        trackSync(groupId, key, processItem(costMix, setCostMix, key));
    }
    const zeroed = new Set(syncUpdates.filter(u => u.newAmountIncl <= 0).map(u => `${u.subjectRef.side}:${u.subjectRef.groupId}:${u.subjectRef.key}`));
    const positive = positiveFundingSubjects(state).filter(s => !zeroed.has(`${s.subjectRef.side}:${s.subjectRef.groupId}:${s.subjectRef.key}`));
    // Use a functional setter at the caller so unrelated plan edits are not lost.
    const synchronizePlans = (current: SubjectFundingPlans) => syncUpdates.length
        ? initializeMissingSubjectFundingPlans(syncSubjectFundingPlansToAmounts(current, syncUpdates), positive) : current;
    return { state: next, plans: synchronizePlans(plans), synchronizePlans, plansChanged: syncUpdates.length > 0 };
}
export const serializeTaxItemForPayload = (item: TaxItem) => {
    const customSubjectName = normalizeCustomSubjectName(item.customSubjectName);
    const billingSubjectName = normalizeCustomSubjectName(item.billingSubjectName);
    return {
        incl_tax: String(item.incl),
        tax_rate: String(item.tax),
        ...(customSubjectName ? { custom_subject_name: customSubjectName } : {}),
        ...(billingSubjectName ? { billing_subject_name: billingSubjectName } : {}),
        ...(item.splitParts?.length
            ? {
                split_parts: serializeTaxSplitParts(item.splitParts),
            }
            : {}),
    };
};
