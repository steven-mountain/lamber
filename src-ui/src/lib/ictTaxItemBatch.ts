import { normalizeTaxPairFromIncl } from './taxAmount';
import { type IctSubjectState, type TaxItem } from './ictTaxItemEdit';
import { initializeMissingSubjectFundingPlans, syncSubjectFundingPlansToAmounts,
  type SubjectFundingPlans, type SubjectFundingSubjectRef, type SubjectFundingPlanLastChangeReason } from './ictSubjectFundingPlan';

export type TaxItemInclUpdate = { groupId: string; key: string; incl: number; reason?: SubjectFundingPlanLastChangeReason };
/** Pure form of the desktop batch write. Preview and commit share normalization,
 * split invalidation, CT linkage and plan synchronization; this never sets state. */
export function prepareIctTaxItemsInclBatch(state: IctSubjectState, updates: TaxItemInclUpdate[], autoFix: boolean) {
  const { revIt, revCt, revNonItCt, costIt, costCt, costMix } = state;
  let nextRevIt = revIt;
  let nextRevCt = revCt;
  let nextRevNonItCt = revNonItCt;
  let nextCostIt = costIt;
  let nextCostCt = costCt;
  let nextCostMix = costMix;
  const itemFromIncl = (current: TaxItem | undefined, incl: number): TaxItem => {
    const tax = Number(current?.tax ?? 0);
    // 程序化写入（导入/差额承接等）：开启自动修正时归一到财务口径不动点，
    // 否则保留写入值。调用方必须在提交前校验此有效状态。
    const pair = normalizeTaxPairFromIncl(isNaN(incl) ? 0 : incl, tax);
    return {
      ...(current || { incl: 0, tax, excl: 0 }),
      incl: autoFix ? pair.incl : pair.enteredIncl,
      tax,
      excl: pair.excl,
      splitParts: undefined,
    };
  };

  const setRecordItem = <T extends Record<string, TaxItem>>(
    group: T,
    key: string,
    incl: number,
  ): T => ({
    ...group,
    [key]: itemFromIncl(group[key], incl),
  } as T);

  updates.forEach(update => {
    if (update.groupId === "revIt") {
      nextRevIt = setRecordItem(nextRevIt, update.key, update.incl);
    } else if (update.groupId === "revCt") {
      nextRevCt = setRecordItem(nextRevCt, update.key, update.incl);
      if (update.key === "product") {
        nextCostCt = setRecordItem(nextCostCt, "other", update.incl);
      }
      if (update.key === "line") {
        nextCostCt = setRecordItem(nextCostCt, "bandwidth", update.incl);
      }
    } else if (update.groupId === "revNonItCt") {
      nextRevNonItCt = itemFromIncl(nextRevNonItCt, update.incl);
    } else if (update.groupId === "costIt") {
      nextCostIt = setRecordItem(nextCostIt, update.key, update.incl);
    } else if (update.groupId === "costCt") {
      nextCostCt = setRecordItem(nextCostCt, update.key, update.incl);
    } else if (update.groupId === "costMix") {
      nextCostMix = setRecordItem(nextCostMix, update.key, update.incl);
    }
  });


  const syncUpdates: Array<{ subjectRef: SubjectFundingSubjectRef; newAmountIncl: number; reason?: SubjectFundingPlanLastChangeReason }> = [];
  // 计划同步金额取归一后的科目含税值，保证计划合计与科目金额逐分一致。
  const normalizedIncl = (groupId: string, key: string, rawIncl: number): number => {
    const groupState =
      groupId === "revIt" ? nextRevIt
      : groupId === "revCt" ? nextRevCt
      : groupId === "costIt" ? nextCostIt
      : groupId === "costCt" ? nextCostCt
      : groupId === "costMix" ? nextCostMix
      : null;
    const item = groupId === "revNonItCt" ? nextRevNonItCt : (groupState as Record<string, TaxItem> | null)?.[key];
    return Number(item?.incl ?? rawIncl) || 0;
  };
  updates.forEach(update => {
    const side = (update.groupId === "revIt" || update.groupId === "revCt" || update.groupId === "revNonItCt")
      ? "revenue" as const : "cost" as const;
    syncUpdates.push({
      subjectRef: { side, groupId: update.groupId as SubjectFundingSubjectRef["groupId"], key: update.key },
      newAmountIncl: normalizedIncl(update.groupId, update.key, update.incl),
      reason: update.reason,
    });
    // CT linkage: revCt.product → costCt.other, revCt.line → costCt.bandwidth
    if (update.groupId === "revCt" && update.key === "product") {
      syncUpdates.push({ subjectRef: { side: "cost", groupId: "costCt", key: "other" }, newAmountIncl: normalizedIncl("costCt", "other", update.incl), reason: "ct_linkage_sync" });
    }
    if (update.groupId === "revCt" && update.key === "line") {
      syncUpdates.push({ subjectRef: { side: "cost", groupId: "costCt", key: "bandwidth" }, newAmountIncl: normalizedIncl("costCt", "bandwidth", update.incl), reason: "ct_linkage_sync" });
    }
  });
  const next = { revIt: nextRevIt, revCt: nextRevCt, revNonItCt: nextRevNonItCt,
    costIt: nextCostIt, costCt: nextCostCt, costMix: nextCostMix };
  const zeroed = new Set(syncUpdates.filter(u => u.newAmountIncl <= 0)
    .map(u => `${u.subjectRef.side}:${u.subjectRef.groupId}:${u.subjectRef.key}`));
  const positiveSubjects: Array<{ subjectRef: SubjectFundingSubjectRef; amountIncl: number }> = [];
  for (const groupId of ['revIt', 'revCt', 'revNonItCt', 'costIt', 'costCt', 'costMix'] as const) {
    const items = groupId === 'revNonItCt' ? { item: next.revNonItCt } : next[groupId];
    const side = groupId.startsWith('rev') ? 'revenue' : 'cost';
    for (const [key, item] of Object.entries(items)) {
      if (item.incl > 0) positiveSubjects.push({ subjectRef: { side, groupId, key }, amountIncl: item.incl });
    }
  }
  const positive = positiveSubjects.filter(s =>
    !zeroed.has(`${s.subjectRef.side}:${s.subjectRef.groupId}:${s.subjectRef.key}`));
  return { state: next, synchronizePlans: (plans: SubjectFundingPlans) => updates.length
    ? initializeMissingSubjectFundingPlans(syncSubjectFundingPlansToAmounts(plans, syncUpdates), positive) : plans };
}
