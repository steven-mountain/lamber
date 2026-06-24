import { normalizeAnnualInclValues, normalizeSubjectFundingPlans, createSubjectFundingPlanId } from "../../lib/ictSubjectFundingPlan";
import { ICT_SUBJECT_DEFINITIONS } from "../../lib/ictSubjectCatalog";
import type { IctResult } from "../../utils/projectService";
import type { AiComputeIctExportPreview } from "./ictExport";
import type {
  AiComputeQuoteBlueprint,
  AiComputeQuoteLineItem,
  AiComputeSyncedSubjectSnapshot,
  IntelligentAmountSource,
  IntelligentComputeProjectState,
} from "./types";

type UnknownRecord = Record<string, any>;

const money = (value: unknown) => {
  const number = Number(value);
  return Number.isFinite(number) ? Math.round(number * 100) / 100 : 0;
};

const sameMoney = (left: unknown, right: unknown) => money(left) === money(right);
const sameYears = (left: unknown, right: unknown) => {
  const a = normalizeAnnualInclValues(left);
  const b = normalizeAnnualInclValues(right);
  return a.every((value, index) => sameMoney(value, b[index]));
};

export function getAiComputeSyncFingerprint(blueprint: AiComputeQuoteBlueprint) {
  const {
    syncState: _syncState,
    parameterGroups: _parameterGroups,
    parameters,
    ...businessState
  } = blueprint;
  const businessParameters = parameters
    .map(({ groupId: _groupId, isKey: _isKey, category: _category, ...parameter }) => parameter)
    .sort((left, right) => left.id.localeCompare(right.id));
  return JSON.stringify({ ...businessState, parameters: businessParameters });
}

export function buildIntelligentComputeSyncLock(
  projectState: IntelligentComputeProjectState,
  amountSources: IntelligentAmountSource[],
) {
  return {
    expectedSyncRevision: projectState.syncRevision,
    sourceVersions: Object.fromEntries(amountSources.map(source => [source.id, source.sourceVersion])),
  };
}

const ICT_RESULT_METRIC_KEYS = [
  "npv",
  "npv_rate",
  "margin_rate",
  "dynamic_payback",
  "irr",
  "it_npv",
  "it_npv_rate",
  "it_margin_rate",
] as const;

function coerceIctMetric(value: unknown) {
  if (typeof value === "string" && value.trim()) return value;
  if (typeof value === "number" && Number.isFinite(value)) return String(value);
  return null;
}

function coerceCashflowRows(value: unknown): IctResult["cashflow"] {
  if (!Array.isArray(value)) return [];
  return value.map(row => {
    const record = row && typeof row === "object" ? row as UnknownRecord : {};
    return {
      year: Number(record.year) || 0,
      cash_in: coerceIctMetric(record.cash_in) || "0",
      cash_out: coerceIctMetric(record.cash_out) || "0",
      net_cash: coerceIctMetric(record.net_cash) || "0",
      cum_net_cash: coerceIctMetric(record.cum_net_cash) || "0",
      pv: coerceIctMetric(record.pv) || "0",
      cum_pv: coerceIctMetric(record.cum_pv) || "0",
    };
  });
}

export function restoreIctResultFromProjectState(
  projectState: Pick<IntelligentComputeProjectState, "syncRevision" | "lastResult"> | null | undefined,
): IctResult | null {
  if (!projectState || projectState.syncRevision <= 0) return null;
  const raw = projectState.lastResult;
  if (!raw || typeof raw !== "object") return null;
  const record = raw as UnknownRecord;
  const metrics = Object.fromEntries(
    ICT_RESULT_METRIC_KEYS.map(key => [key, coerceIctMetric(record[key])]),
  ) as Record<typeof ICT_RESULT_METRIC_KEYS[number], string | null>;
  if (ICT_RESULT_METRIC_KEYS.some(key => metrics[key] === null)) return null;
  return {
    npv: metrics.npv!,
    npv_rate: metrics.npv_rate!,
    margin_rate: metrics.margin_rate!,
    dynamic_payback: metrics.dynamic_payback!,
    irr: metrics.irr!,
    it_npv: metrics.it_npv!,
    it_npv_rate: metrics.it_npv_rate!,
    it_margin_rate: metrics.it_margin_rate!,
    cashflow: coerceCashflowRows(record.cashflow),
  };
}

export function projectStateSyncIncludesAmountSource(
  projectState: Pick<IntelligentComputeProjectState, "controlledSubjects"> | null | undefined,
  amountSourceId: string | null | undefined,
) {
  if (!projectState || !amountSourceId) return false;
  return Object.values(projectState.controlledSubjects || {}).some(snapshot =>
    snapshot.sourceLineItemIds.some(lineItemId => lineItemId.startsWith(`${amountSourceId}:`))
  );
}

export function applySuccessfulAiComputeSync(
  blueprint: AiComputeQuoteBlueprint,
  preview: AiComputeIctExportPreview,
  revision: number,
  syncedAt: string,
): AiComputeQuoteBlueprint {
  const subjects: Record<string, AiComputeSyncedSubjectSnapshot> = {
    ...(blueprint.syncState?.subjects || {}),
  };
  preview.rows.forEach(row => {
    const key = `${row.side}:${row.ictSubjectCode}`;
    if (row.syncStatus === "released_mapping") {
      delete subjects[key];
      return;
    }
    if (row.syncStatus === "paused_override" || row.syncStatus === "merge_conflict") return;
    subjects[key] = {
      side: row.side,
      ictSubjectCode: row.ictSubjectCode,
      amountInclTax: row.writtenAmount,
      taxRate: row.taxRate,
      yearlyAmounts: normalizeAnnualInclValues(row.yearlyAmounts),
      sourceLineItemIds: [...row.sourceLineItemIds],
    };
  });
  return {
    ...blueprint,
    syncState: {
      revision,
      status: preview.rows.some(row =>
        row.syncStatus === "merge_conflict" || row.syncStatus === "paused_override"
      ) ? "conflict" : "synced",
      syncedAt,
      subjects,
    },
  };
}

const patchItems = (
  blueprint: AiComputeQuoteBlueprint,
  patches: Map<string, Partial<AiComputeQuoteLineItem>>,
): AiComputeQuoteBlueprint => ({
  ...blueprint,
  revenueItems: blueprint.revenueItems.map(item => ({ ...item, ...(patches.get(item.id) || {}) })),
  costItems: blueprint.costItems.map(item => ({ ...item, ...(patches.get(item.id) || {}) })),
});

export function reconcileAiComputeBlueprintFromIct(
  blueprint: AiComputeQuoteBlueprint,
  inputPayload: UnknownRecord,
  assumptions: UnknownRecord,
): { blueprint: AiComputeQuoteBlueprint; changed: boolean; conflicts: string[] } {
  const snapshots = blueprint.syncState?.subjects || {};
  const plans = normalizeSubjectFundingPlans(
    assumptions.subjectFundingPlans || assumptions.subject_funding_plans || inputPayload.subject_funding_plans,
  );
  const patches = new Map<string, Partial<AiComputeQuoteLineItem>>();
  const conflicts: string[] = [];

  Object.values(snapshots).forEach(snapshot => {
    const subject = ICT_SUBJECT_DEFINITIONS.find(candidate =>
      candidate.side === snapshot.side && candidate.subjectCode === snapshot.ictSubjectCode
    );
    if (!subject) return;
    const item = inputPayload[snapshot.ictSubjectCode] || {};
    const amountInclTax = money(item.incl_tax ?? item.incl);
    const taxRate = money(item.tax_rate ?? item.tax);
    const plan = plans[createSubjectFundingPlanId({
      side: subject.side,
      groupId: subject.groupId,
      key: subject.key,
    })];
    const yearlyAmounts = normalizeAnnualInclValues(plan?.annualInclValues);
    if (
      sameMoney(amountInclTax, snapshot.amountInclTax)
      && sameMoney(taxRate, snapshot.taxRate)
      && sameYears(yearlyAmounts, snapshot.yearlyAmounts)
    ) {
      return;
    }

    if (snapshot.sourceLineItemIds.length === 1) {
      const lineItemId = snapshot.sourceLineItemIds[0];
      patches.set(lineItemId, {
        formulaControlStatus: "ict_override",
        ictOverride: {
          ictSubjectCode: snapshot.ictSubjectCode,
          amountInclTax,
          taxRate,
          yearlyAmounts,
          modifiedAt: new Date().toISOString(),
        },
        ictControlMessage: "历史 ICT 差异标记（仅兼容读取）",
      });
      return;
    }

    const message = `${subject.standardSubjectName}由多个智算项合并，ICT 人工修改后无法自动拆分`;
    conflicts.push(message);
    snapshot.sourceLineItemIds.forEach(lineItemId => {
      patches.set(lineItemId, {
        formulaControlStatus: "merge_conflict",
        ictControlMessage: message,
      });
    });
  });

  if (patches.size === 0) return { blueprint, changed: false, conflicts };
  const next = patchItems(blueprint, patches);
  return {
    blueprint: {
      ...next,
      syncState: {
        revision: blueprint.syncState?.revision || 0,
        status: conflicts.length > 0 ? "conflict" : "synced",
        syncedAt: blueprint.syncState?.syncedAt,
        subjects: blueprint.syncState?.subjects || {},
      },
    },
    changed: true,
    conflicts,
  };
}

export function clearAiComputeControlForMappingChange(
  blueprint: AiComputeQuoteBlueprint,
  lineItemId: string,
): AiComputeQuoteBlueprint {
  const item = [...blueprint.revenueItems, ...blueprint.costItems]
    .find(candidate => candidate.id === lineItemId);
  if (!item || (
    item.formulaControlStatus !== "ict_override"
    && item.formulaControlStatus !== "merge_conflict"
  )) {
    return blueprint;
  }
  return restoreAiComputeFormulaControl(blueprint, lineItemId);
}

export function restoreAiComputeFormulaControl(
  blueprint: AiComputeQuoteBlueprint,
  lineItemId: string,
): AiComputeQuoteBlueprint {
  const patches = new Map<string, Partial<AiComputeQuoteLineItem>>();
  patches.set(lineItemId, {
    formulaControlStatus: "formula",
    ictOverride: undefined,
    ictControlMessage: undefined,
  });
  return patchItems(blueprint, patches);
}

export function mergePersistedAiComputeControlState(
  local: AiComputeQuoteBlueprint,
  persisted: AiComputeQuoteBlueprint,
): AiComputeQuoteBlueprint {
  const persistedItems = new Map(
    [...persisted.revenueItems, ...persisted.costItems].map(item => [item.id, item]),
  );
  const mergeItem = (item: AiComputeQuoteLineItem): AiComputeQuoteLineItem => {
    const persistedItem = persistedItems.get(item.id);
    if (!persistedItem) return item;
    return {
      ...item,
      formulaControlStatus: persistedItem.formulaControlStatus,
      ictOverride: persistedItem.ictOverride,
      ictControlMessage: persistedItem.ictControlMessage,
    };
  };
  return {
    ...local,
    revenueItems: local.revenueItems.map(mergeItem),
    costItems: local.costItems.map(mergeItem),
    syncState: persisted.syncState,
  };
}
