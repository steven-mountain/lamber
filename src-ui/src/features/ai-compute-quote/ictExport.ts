import {
  ICT_SUBJECT_DEFINITIONS,
  type IctSubjectDefinition,
} from "../../lib/ictSubjectCatalog";
import {
  createSubjectFundingPlanId,
  normalizeAnnualInclValues,
  normalizeSubjectFundingPlans,
  SUBJECT_FUNDING_PLAN_MIGRATION_VERSION,
  type SubjectFundingPlan,
  type SubjectFundingPlanImportTrace,
  type SubjectFundingPlans,
} from "../../lib/ictSubjectFundingPlan";
import {
  buildAiComputeQuoteOutputFundingPlans,
  calculateQuoteBlueprint,
} from "./calculations";
import {
  getAiComputeDiscountRateDecimal,
  getAiComputeProjectCycleYears,
  validateAiComputeFundingPlan,
} from "./fundingPlans";
import type {
  AiComputeLineItemFundingPlanMode,
  AiComputeOutputSubjectFundingPlan,
  AiComputeQuoteBlueprint,
  AiComputeSyncedSubjectSnapshot,
} from "./types";

type UnknownRecord = Record<string, any>;

export type AiComputeIctExportRow = {
  side: "revenue" | "cost";
  ictSubjectCode: string;
  ictSubjectName: string;
  subject: IctSubjectDefinition;
	  originalAmount: number;
	  quoteAmount: number;
	  writtenAmount: number;
	  amountExclTax: number;
	  originalYearlyAmounts: number[];
	  yearlyAmounts: number[];
  sourceLineItemIds: string[];
  sourceLineItemNames: string[];
  fundingPlanModes: AiComputeLineItemFundingPlanMode[];
  taxRate: number;
  syncStatus?: "ready" | "zeroed_error" | "zeroed_absent" | "paused_override" | "merge_conflict" | "released_mapping";
  syncMessages?: string[];
};

export type AiComputeIctExportPreview = {
  projectId: string;
  scenarioId: string;
  blueprintId: string;
  projectYears: number;
  discountRate: number;
  rows: AiComputeIctExportRow[];
  skippedUnmappedItems: Array<{ id: string; name: string; side: "revenue" | "cost"; reason: string }>;
  skippedItems: Array<{ id: string; name: string; side: "revenue" | "cost"; reason: string }>;
};

export type IntelligentComputeAggregateSource = {
  sourceId: string;
  sourceName: string;
  blueprint: AiComputeQuoteBlueprint;
};

const zeroYearlyAmounts = () => Array(10).fill(0);

export type AiComputeIctExportPayloads = {
  lifecycleState: {
    profileJson: UnknownRecord;
    parametersJson: UnknownRecord;
    backgroundJson: UnknownRecord;
    inputPayloadJson: UnknownRecord;
  };
  cashflowState: {
    cashflowModel?: string | null;
    paymentModelJson: UnknownRecord;
    yearlyCashflowJson: UnknownRecord;
    sectorCashflowJson: UnknownRecord;
    assumptionsJson: UnknownRecord;
    metricsJson: UnknownRecord;
  };
};

const finiteMoney = (value: unknown) => {
  const numeric = Number(value);
  return Number.isFinite(numeric) ? Math.round(numeric * 100) / 100 : 0;
};

const cloneRecord = (value: unknown): UnknownRecord =>
  value && typeof value === "object"
    ? JSON.parse(JSON.stringify(value)) as UnknownRecord
    : {};

const firstDefined = (...values: unknown[]) =>
  values.find(value => value !== undefined && value !== null);

const normalizeDistribution = (value: unknown) => {
  if (!Array.isArray(value)) return [1, 0, 0, 0, 0, 0, 0, 0, 0, 0];
  return Array.from({ length: 10 }, (_, index) => finiteMoney(value[index]));
};

const getAssumptionSubjectItem = (assumptions: UnknownRecord, subject: IctSubjectDefinition) => {
  if (subject.groupId === "revNonItCt") return assumptions.revNonItCt;
  return assumptions[subject.groupId]?.[subject.key];
};

const setAssumptionSubjectItem = (
  assumptions: UnknownRecord,
  subject: IctSubjectDefinition,
  item: UnknownRecord,
) => {
  if (subject.groupId === "revNonItCt") {
    assumptions.revNonItCt = item;
    return;
  }
  assumptions[subject.groupId] = {
    ...(assumptions[subject.groupId] || {}),
    [subject.key]: item,
  };
};

const buildCompleteIctInput = (
  fullState: UnknownRecord,
  inputPayload: UnknownRecord,
) => {
  const project = fullState.project || {};
  const lifecycle = fullState.lifecycleState || {};
  const cashflow = fullState.cashflowState || {};
  const profile = cloneRecord(lifecycle.profileJson);
  const parameters = cloneRecord(lifecycle.parametersJson);
  const background = cloneRecord(lifecycle.backgroundJson);
  const paymentModel = cloneRecord(cashflow.paymentModelJson);
  const sectorCashflow = cloneRecord(cashflow.sectorCashflowJson);
  const assumptions = cloneRecord(cashflow.assumptionsJson);
  const completeInput: UnknownRecord = {
    ...inputPayload,
    project_name: String(firstDefined(
      inputPayload.project_name,
      profile.projectName,
      profile.project_name,
      project.name,
      "",
    )),
    customer_name: String(firstDefined(
      inputPayload.customer_name,
      profile.customerName,
      profile.customer_name,
      project.customer_name,
      "",
    )),
    property_rights: String(firstDefined(
      inputPayload.property_rights,
      profile.propertyRights,
      profile.property_rights,
      assumptions.propertyRights,
      assumptions.property_rights,
      "客户",
    )),
    discount_rate: String(firstDefined(
      inputPayload.discount_rate,
      parameters.discountRate,
      parameters.discount_rate,
      assumptions.discountRate,
      assumptions.discount_rate,
      project.discount_rate,
      0.055,
    )),
    project_years: finiteMoney(firstDefined(
      inputPayload.project_years,
      parameters.projectYears,
      parameters.project_years,
      assumptions.projectYears,
      assumptions.project_years,
      project.project_years,
      1,
    )),
    cashflow_model: String(firstDefined(
      inputPayload.cashflow_model,
      cashflow.cashflowModel,
      paymentModel.cashflowModel,
      paymentModel.cashflow_model,
      project.cashflow_model,
      "model_a",
    )),
    cashflow_segment_value_mode: String(firstDefined(
      inputPayload.cashflow_segment_value_mode,
      paymentModel.segmentValueMode,
      paymentModel.cashflow_segment_value_mode,
      "ratio",
    )),
    cashflow_segments: firstDefined(
      inputPayload.cashflow_segments,
      sectorCashflow.cashflowSegments,
      sectorCashflow.cashflow_segments,
      [],
    ),
    project_background: String(firstDefined(
      inputPayload.project_background,
      background.projectBackground,
      background.project_background,
      "",
    )),
    rev_distribution: normalizeDistribution(firstDefined(
      inputPayload.rev_distribution,
      paymentModel.revDistribution,
      paymentModel.rev_distribution,
    )),
    cost_distribution: normalizeDistribution(firstDefined(
      inputPayload.cost_distribution,
      paymentModel.costDistribution,
      paymentModel.cost_distribution,
    )),
    ignore_tail_difference: Boolean(firstDefined(
      inputPayload.ignore_tail_difference,
      parameters.ignoreTailDifference,
      parameters.ignore_tail_difference,
      false,
    )),
    tail_difference_value: String(firstDefined(
      inputPayload.tail_difference_value,
      parameters.tailDifferenceValue,
      parameters.tail_difference_value,
      "0",
    )),
  };

  ICT_SUBJECT_DEFINITIONS.forEach(subject => {
    const inputItem = cloneRecord(completeInput[subject.subjectCode]);
    const assumptionItem = cloneRecord(getAssumptionSubjectItem(assumptions, subject));
    const customSubjectName = String(firstDefined(
      inputItem.custom_subject_name,
      inputItem.customSubjectName,
      assumptionItem.customSubjectName,
      assumptionItem.custom_subject_name,
      "",
    )).trim();
    const billingSubjectName = String(firstDefined(
      inputItem.billing_subject_name,
      inputItem.billingSubjectName,
      assumptionItem.billingSubjectName,
      assumptionItem.billing_subject_name,
      "",
    )).trim();
	    const amountInclTax = firstDefined(
	      inputItem.incl_tax,
	      inputItem.incl,
	      assumptionItem.incl,
	      assumptionItem.incl_tax,
	      0,
	    );
	    const rawTaxRate = firstDefined(
	      inputItem.tax_rate,
	      inputItem.tax,
	      assumptionItem.tax,
	      assumptionItem.tax_rate,
	    );
	    completeInput[subject.subjectCode] = {
	      ...inputItem,
	      incl_tax: String(amountInclTax),
	      tax_rate: String(rawTaxRate === undefined || rawTaxRate === null || rawTaxRate === ""
	        ? subject.defaultTaxRate
	        : rawTaxRate),
	      ...(customSubjectName ? { custom_subject_name: customSubjectName } : {}),
	      ...(billingSubjectName ? { billing_subject_name: billingSubjectName } : {}),
	    };
  });

  return completeInput;
};

const deriveTaxRate = (amountInclTax: number, amountExclTax: number) => {
  if (amountInclTax <= 0 || amountExclTax <= 0) return 0;
  return Math.round(((amountInclTax / amountExclTax) - 1) * 100 * 1_000_000) / 1_000_000;
};

const getBaseInput = (fullState: UnknownRecord) => {
  const legacyInput = cloneRecord(
    fullState.legacyLifecycleInput
    || fullState.latestSnapshot?.input_params
    || fullState.latestSnapshot?.inputParams,
  );
  const lifecycleInput = cloneRecord(fullState.lifecycleState?.inputPayloadJson);
  return buildCompleteIctInput(fullState, {
    ...legacyInput,
    ...lifecycleInput,
  });
};

const getExistingAmount = (
  fullState: UnknownRecord,
  subject: IctSubjectDefinition,
) => {
  const assumptions = fullState.cashflowState?.assumptionsJson || {};
  const assumptionItem = getAssumptionSubjectItem(assumptions, subject);
  if (assumptionItem) return finiteMoney(assumptionItem.incl ?? assumptionItem.incl_tax);
  const inputItem = getBaseInput(fullState)[subject.subjectCode];
  return finiteMoney(inputItem?.incl_tax ?? inputItem?.incl);
};

const getExistingTaxRate = (
  fullState: UnknownRecord,
  subject: IctSubjectDefinition,
) => {
	  const assumptions = fullState.cashflowState?.assumptionsJson || {};
	  const assumptionItem = getAssumptionSubjectItem(assumptions, subject);
	  if (assumptionItem) return finiteMoney(assumptionItem.tax ?? assumptionItem.tax_rate ?? subject.defaultTaxRate);
	  const inputItem = getBaseInput(fullState)[subject.subjectCode];
	  return finiteMoney(inputItem?.tax_rate ?? inputItem?.tax ?? subject.defaultTaxRate);
	};

const getExistingYearlyAmounts = (
  fullState: UnknownRecord,
  subject: IctSubjectDefinition,
) => normalizeAnnualInclValues(
  normalizeSubjectFundingPlans(
    fullState.cashflowState?.assumptionsJson?.subjectFundingPlans
    || fullState.cashflowState?.assumptionsJson?.subject_funding_plans,
  )[createSubjectFundingPlanId({
    side: subject.side,
    groupId: subject.groupId,
    key: subject.key,
  })]?.annualInclValues,
);

const normalizeAiComputeImportTrace = (
  value: unknown,
): SubjectFundingPlanImportTrace | null => {
	  if (!value || typeof value !== "object") return null;
	  const raw = value as UnknownRecord;
	  if (raw.source !== "ai_compute_quote" && raw.source !== "intelligent_compute") return null;
	  const projectId = String(raw.projectId ?? raw.project_id ?? "").trim();
	  const scenarioId = String(raw.scenarioId ?? raw.scenario_id ?? "").trim();
	  const blueprintId = String(raw.blueprintId ?? raw.blueprint_id ?? "").trim();
  const sourceLineItemIds = Array.isArray(raw.sourceLineItemIds ?? raw.source_line_item_ids)
    ? (raw.sourceLineItemIds ?? raw.source_line_item_ids)
      .map((id: unknown) => String(id).trim())
      .filter(Boolean)
    : [];
	  if (!projectId || (!scenarioId && !blueprintId) || sourceLineItemIds.length === 0) return null;
	  return {
	    source: raw.source,
	    sourceLabel: String(raw.sourceLabel ?? raw.source_label ?? "来自智算报价测算"),
	    projectId,
    scenarioId,
    blueprintId,
    sourceLineItemIds,
    importedAt: String(raw.importedAt ?? raw.imported_at ?? ""),
  };
};

const traceMatchesBlueprint = (
  trace: SubjectFundingPlanImportTrace,
  blueprint: AiComputeQuoteBlueprint,
  projectId: string,
) => {
  const scenarioId = blueprint.scenarioId || blueprint.id;
  return trace.projectId === projectId
    && (!trace.blueprintId || trace.blueprintId === blueprint.id)
    && (!trace.scenarioId || trace.scenarioId === scenarioId);
};

const collectPreviouslyControlledSubjects = (
  blueprint: AiComputeQuoteBlueprint,
  fullState: UnknownRecord,
  projectId: string,
) => {
  const subjects = new Map<string, AiComputeSyncedSubjectSnapshot>(
    Object.entries(blueprint.syncState?.subjects || {}),
  );
  const assumptions = fullState.cashflowState?.assumptionsJson || {};
  const plans = normalizeSubjectFundingPlans(
    assumptions.subjectFundingPlans || assumptions.subject_funding_plans,
  );

  ICT_SUBJECT_DEFINITIONS.forEach(subject => {
    const key = `${subject.side}:${subject.subjectCode}`;
    if (subjects.has(key)) return;
    const plan = plans[createSubjectFundingPlanId({
      side: subject.side,
      groupId: subject.groupId,
      key: subject.key,
    })];
    const trace = normalizeAiComputeImportTrace(plan?.importTrace);
    const isUnmodifiedAiImport = plan?.source === "ai_compute_quote"
      && (!plan.lastChangeReason || plan.lastChangeReason === "ai_compute_quote_import");
    if (!trace || !isUnmodifiedAiImport || !traceMatchesBlueprint(trace, blueprint, projectId)) {
      return;
    }
    const yearlyAmounts = normalizeAnnualInclValues(plan.annualInclValues);
    const amountInclTax = getExistingAmount(fullState, subject);
    if (
      Math.abs(amountInclTax) <= 0.004
      && yearlyAmounts.every(value => Math.abs(value) <= 0.004)
    ) {
      return;
    }

    subjects.set(key, {
      side: subject.side,
      ictSubjectCode: subject.subjectCode,
      amountInclTax,
      taxRate: getExistingTaxRate(fullState, subject),
      yearlyAmounts,
      sourceLineItemIds: [...trace.sourceLineItemIds],
    });
  });

  return subjects;
};

const collectPreviouslyControlledIntelligentSubjects = (
  fullState: UnknownRecord,
  projectId: string,
) => {
  const subjects = new Map<string, AiComputeSyncedSubjectSnapshot>();
  const assumptions = fullState.cashflowState?.assumptionsJson || {};
  const plans = normalizeSubjectFundingPlans(
    assumptions.subjectFundingPlans || assumptions.subject_funding_plans,
  );

  ICT_SUBJECT_DEFINITIONS.forEach(subject => {
    const plan = plans[createSubjectFundingPlanId({
      side: subject.side,
      groupId: subject.groupId,
      key: subject.key,
    })];
    const trace = normalizeAiComputeImportTrace(plan?.importTrace);
    const isUnmodifiedIntelligentImport = plan?.source === "intelligent_compute"
      && (!plan.lastChangeReason || plan.lastChangeReason === "intelligent_compute_import");
    if (
      !trace
      || trace.source !== "intelligent_compute"
      || trace.projectId !== projectId
      || !isUnmodifiedIntelligentImport
    ) {
      return;
    }

    const yearlyAmounts = normalizeAnnualInclValues(plan?.annualInclValues);
    const amountInclTax = getExistingAmount(fullState, subject);
    if (
      Math.abs(amountInclTax) <= 0.004
      && yearlyAmounts.every(value => Math.abs(value) <= 0.004)
    ) {
      return;
    }

    subjects.set(`${subject.side}:${subject.subjectCode}`, {
      side: subject.side,
      ictSubjectCode: subject.subjectCode,
      amountInclTax,
      taxRate: getExistingTaxRate(fullState, subject),
      yearlyAmounts,
      sourceLineItemIds: [...trace.sourceLineItemIds],
    });
  });

  return subjects;
};

const collectSkippedItems = (
  blueprint: AiComputeQuoteBlueprint,
  mappedLineItemIds: Set<string>,
) => {
  const calculated = calculateQuoteBlueprint(blueprint);
  const skippedUnmappedItems: AiComputeIctExportPreview["skippedUnmappedItems"] = [];
  const skippedItems: AiComputeIctExportPreview["skippedItems"] = [];
  [...calculated.revenueItems, ...calculated.costItems].forEach(item => {
    if (!item.enabled) {
      skippedItems.push({ id: item.id, name: item.name, side: item.side, reason: "计算项已禁用" });
    } else if (!item.outputEnabled) {
      skippedItems.push({ id: item.id, name: item.name, side: item.side, reason: "未启用输出" });
    } else if (item.calculationStatus !== "valid") {
      skippedItems.push({ id: item.id, name: item.name, side: item.side, reason: item.calculationError || "计算结果无效" });
    } else if (!item.fundingPlan?.enabled) {
      skippedItems.push({ id: item.id, name: item.name, side: item.side, reason: "资金计划未启用" });
    } else if (!mappedLineItemIds.has(item.id)) {
      skippedUnmappedItems.push({ id: item.id, name: item.name, side: item.side, reason: "未映射 ICT 科目" });
    }
  });
  return { skippedUnmappedItems, skippedItems };
};

export function buildAiComputeIctExportPreview(
  blueprint: AiComputeQuoteBlueprint,
  fullState: UnknownRecord,
  projectId: string,
): AiComputeIctExportPreview {
  const calculated = calculateQuoteBlueprint(blueprint);
  const fundingOutputs = buildAiComputeQuoteOutputFundingPlans(blueprint);
  const lineItems = [...calculated.revenueItems, ...calculated.costItems];
  const lineItemMap = new Map(lineItems.map(item => [item.id, item]));
  const mappedLineItemIds = new Set(fundingOutputs.flatMap(output => output.sourceLineItemIds));
  const existingAssumptions = fullState.cashflowState?.assumptionsJson || {};
  const existingPlans = normalizeSubjectFundingPlans(
    existingAssumptions.subjectFundingPlans || existingAssumptions.subject_funding_plans,
  );

  const rows = fundingOutputs.flatMap((fundingOutput: AiComputeOutputSubjectFundingPlan) => {
    const subject = ICT_SUBJECT_DEFINITIONS.find(candidate =>
      candidate.side === fundingOutput.side && candidate.subjectCode === fundingOutput.ictSubjectCode
    );
    if (!subject) return [];
    const sourceItems = fundingOutput.sourceLineItemIds
      .map(id => lineItemMap.get(id))
      .filter((item): item is AiComputeQuoteBlueprint["revenueItems"][number] => Boolean(item));
    const amountInclTax = sourceItems.reduce((sum, item) => sum + item.amountInclTax, 0);
    const amountExclTax = sourceItems.reduce((sum, item) => sum + item.amountExclTax, 0);
    const originalPlan = existingPlans[createSubjectFundingPlanId({
      side: subject.side,
      groupId: subject.groupId,
      key: subject.key,
    })];
    return [{
      side: fundingOutput.side,
      ictSubjectCode: fundingOutput.ictSubjectCode,
      ictSubjectName: fundingOutput.ictSubjectName,
      subject,
	      originalAmount: getExistingAmount(fullState, subject),
	      quoteAmount: finiteMoney(fundingOutput.totalAmount),
	      writtenAmount: finiteMoney(fundingOutput.totalAmount),
	      amountExclTax: finiteMoney(amountExclTax),
	      originalYearlyAmounts: normalizeAnnualInclValues(originalPlan?.annualInclValues),
      yearlyAmounts: normalizeAnnualInclValues(
        Array.from({ length: 10 }, (_, index) => fundingOutput.yearlyAmounts[String(index + 1)] || 0),
      ),
      sourceLineItemIds: [...fundingOutput.sourceLineItemIds],
      sourceLineItemNames: fundingOutput.sourceLineItemIds.map(id => lineItemMap.get(id)?.name || id),
      fundingPlanModes: Array.from(new Set(sourceItems.flatMap(item =>
        item.fundingPlan?.mode ? [item.fundingPlan.mode] : []
      ))),
      taxRate: deriveTaxRate(amountInclTax, amountExclTax),
    }];
  });

  const skipped = collectSkippedItems(blueprint, mappedLineItemIds);
  return {
    projectId,
    scenarioId: blueprint.scenarioId || blueprint.id,
    blueprintId: blueprint.id,
    projectYears: getAiComputeProjectCycleYears(calculated.parameters),
    discountRate: getAiComputeDiscountRateDecimal(calculated.parameters),
    rows,
    ...skipped,
  };
}

export function validateIntelligentComputeSources(
  sources: IntelligentComputeAggregateSource[],
) {
  const issues: string[] = [];
  const selectedSources = sources.slice(0, 1);

  selectedSources.forEach(source => {
    const calculated = calculateQuoteBlueprint(source.blueprint);
    const mappings = new Map(
      calculated.mappings
        .filter(mapping => mapping.enabled)
        .map(mapping => [mapping.lineItemId, mapping]),
    );
    [...calculated.revenueItems, ...calculated.costItems].forEach(item => {
      const mapping = mappings.get(item.id);
      if (!item.enabled || !item.outputEnabled || !mapping) return;
      if (item.calculationStatus !== "valid") {
        issues.push(`${source.sourceName} / ${item.name}：${item.calculationError || "计算结果无效"}`);
        return;
      }
      if (!item.fundingPlan?.enabled) {
        issues.push(`${source.sourceName} / ${item.name}：年度金额未启用`);
        return;
      }
      const validation = validateAiComputeFundingPlan(item.fundingPlan, item.amountInclTax);
      if (!validation.consistent) {
        issues.push(`${source.sourceName} / ${item.name}：年度金额差异 ${validation.difference} 元`);
      }
    });
  });
  return { valid: issues.length === 0, issues };
}

export function buildIntelligentComputeAggregatePreview(
  sources: IntelligentComputeAggregateSource[],
  fullState: UnknownRecord,
  projectId: string,
  controlledSubjects: Record<string, AiComputeSyncedSubjectSnapshot> = {},
): AiComputeIctExportPreview {
  const aggregate = new Map<string, AiComputeIctExportRow>();
  const skippedUnmappedItems: AiComputeIctExportPreview["skippedUnmappedItems"] = [];
  const skippedItems: AiComputeIctExportPreview["skippedItems"] = [];
  let projectYears = 1;
  let discountRate = 0.055;
  const selectedSources = sources.slice(0, 1);

  selectedSources.forEach(source => {
    const preview = buildAiComputeIctExportPreview(source.blueprint, fullState, projectId);
    projectYears = preview.projectYears;
    discountRate = preview.discountRate;
    preview.skippedUnmappedItems.forEach(item => skippedUnmappedItems.push({
      ...item,
      id: `${source.sourceId}:${item.id}`,
      name: `${source.sourceName} / ${item.name}`,
    }));
    preview.skippedItems.forEach(item => skippedItems.push({
      ...item,
      id: `${source.sourceId}:${item.id}`,
      name: `${source.sourceName} / ${item.name}`,
    }));
    preview.rows.forEach(row => {
      const key = `${row.side}:${row.ictSubjectCode}`;
      const compoundIds = row.sourceLineItemIds.map(id => `${source.sourceId}:${id}`);
      const compoundNames = row.sourceLineItemNames.map(name => `${source.sourceName} / ${name}`);
      const current = aggregate.get(key);
      if (!current) {
        aggregate.set(key, {
          ...row,
          sourceLineItemIds: compoundIds,
          sourceLineItemNames: compoundNames,
          syncStatus: "ready",
        });
        return;
      }
	      const nextAmount = finiteMoney(current.writtenAmount + row.writtenAmount);
	      const nextExcl = finiteMoney(current.amountExclTax + row.amountExclTax);
	      aggregate.set(key, {
	        ...current,
	        quoteAmount: nextAmount,
	        writtenAmount: nextAmount,
	        amountExclTax: nextExcl,
	        yearlyAmounts: current.yearlyAmounts.map((value, index) =>
	          finiteMoney(value + (row.yearlyAmounts[index] || 0))
	        ),
        sourceLineItemIds: [...current.sourceLineItemIds, ...compoundIds],
        sourceLineItemNames: [...current.sourceLineItemNames, ...compoundNames],
        fundingPlanModes: Array.from(new Set([...current.fundingPlanModes, ...row.fundingPlanModes])),
	        taxRate: deriveTaxRate(nextAmount, nextExcl),
	        syncStatus: "ready",
	      });
	    });
	  });

	  const previousControlledSubjects = new Map<string, AiComputeSyncedSubjectSnapshot>(
	    Object.entries(controlledSubjects),
	  );
	  collectPreviouslyControlledIntelligentSubjects(fullState, projectId).forEach((snapshot, key) => {
	    if (!previousControlledSubjects.has(key)) previousControlledSubjects.set(key, snapshot);
	  });

	  previousControlledSubjects.forEach((snapshot, key) => {
	    if (aggregate.has(key)) return;
	    const subject = ICT_SUBJECT_DEFINITIONS.find(candidate =>
	      candidate.side === snapshot.side && candidate.subjectCode === snapshot.ictSubjectCode
    );
    if (!subject) return;
    aggregate.set(key, {
      side: snapshot.side,
      ictSubjectCode: snapshot.ictSubjectCode,
      ictSubjectName: subject.standardSubjectName,
      subject,
	      originalAmount: getExistingAmount(fullState, subject),
	      quoteAmount: 0,
	      writtenAmount: 0,
	      amountExclTax: 0,
	      originalYearlyAmounts: normalizeAnnualInclValues(snapshot.yearlyAmounts),
      yearlyAmounts: zeroYearlyAmounts(),
      sourceLineItemIds: [...snapshot.sourceLineItemIds],
      sourceLineItemNames: [...snapshot.sourceLineItemIds],
      fundingPlanModes: [],
      taxRate: snapshot.taxRate || subject.defaultTaxRate,
      syncStatus: "released_mapping",
      syncMessages: ["该科目已不再由启用的智算金额来源控制，本次同步将清零金额和年度计划"],
    });
  });

  ICT_SUBJECT_DEFINITIONS.forEach(subject => {
    const key = `${subject.side}:${subject.subjectCode}`;
    if (aggregate.has(key)) return;
    aggregate.set(key, {
      side: subject.side,
      ictSubjectCode: subject.subjectCode,
      ictSubjectName: subject.standardSubjectName,
      subject,
      originalAmount: getExistingAmount(fullState, subject),
      quoteAmount: 0,
      writtenAmount: 0,
      amountExclTax: 0,
      originalYearlyAmounts: getExistingYearlyAmounts(fullState, subject),
      yearlyAmounts: zeroYearlyAmounts(),
      sourceLineItemIds: [],
      sourceLineItemNames: [],
      fundingPlanModes: [],
      taxRate: subject.defaultTaxRate,
      syncStatus: "zeroed_absent",
      syncMessages: ["当前同步来源未输出该科目，本次将按 0 覆盖 ICT 金额和年度计划"],
    });
  });

  return {
    projectId,
    scenarioId: "intelligent-compute-aggregate",
    blueprintId: "intelligent-compute-aggregate",
    projectYears,
    discountRate,
    rows: ICT_SUBJECT_DEFINITIONS
      .map(subject => aggregate.get(`${subject.side}:${subject.subjectCode}`))
      .filter((row): row is AiComputeIctExportRow => Boolean(row)),
    skippedUnmappedItems,
    skippedItems,
  };
}

export function buildAiComputeAutoSyncPreview(
  blueprint: AiComputeQuoteBlueprint,
  fullState: UnknownRecord,
  projectId: string,
): AiComputeIctExportPreview {
  const calculated = calculateQuoteBlueprint(blueprint);
  const lineItems = [...calculated.revenueItems, ...calculated.costItems];
  const itemMap = new Map(lineItems.map(item => [item.id, item]));
  const groups = new Map<string, typeof calculated.mappings>();

  calculated.mappings.forEach(mapping => {
    const item = itemMap.get(mapping.lineItemId);
    if (!mapping.enabled || !mapping.ictSubjectCode || !item) return;
    const key = `${mapping.side}:${mapping.ictSubjectCode}`;
    groups.set(key, [...(groups.get(key) || []), mapping]);
  });

  const rows: AiComputeIctExportRow[] = [];
  const skippedUnmappedItems: AiComputeIctExportPreview["skippedUnmappedItems"] = [];
  const skippedItems: AiComputeIctExportPreview["skippedItems"] = [];

  groups.forEach(mappings => {
    const firstMapping = mappings[0];
    const subject = ICT_SUBJECT_DEFINITIONS.find(candidate =>
      candidate.side === firstMapping.side && candidate.subjectCode === firstMapping.ictSubjectCode
    );
    if (!subject) return;
    const items = mappings
      .map(mapping => itemMap.get(mapping.lineItemId))
      .filter((item): item is AiComputeQuoteBlueprint["revenueItems"][number] => Boolean(item));
    const syncMessages: string[] = [];

    const invalidItems = items.filter(item => {
      const planValidation = item.fundingPlan
        ? validateAiComputeFundingPlan(item.fundingPlan, item.amountInclTax)
        : null;
      return !item.enabled
        || !item.outputEnabled
        || item.calculationStatus !== "valid"
        || !item.fundingPlan?.enabled
        || !planValidation?.consistent;
    });
    invalidItems.forEach(item => {
      const reason = !item.enabled
        ? "计算项已禁用"
        : !item.outputEnabled
          ? "未启用输出"
          : item.calculationError || "公式或资金计划无效";
      syncMessages.push(`${item.name}：${reason}，按 0 同步`);
    });

    const activeItems = items.filter(item => !invalidItems.some(invalid => invalid.id === item.id));
    const amountInclTax = activeItems.reduce((sum, item) => sum + item.amountInclTax, 0);
    const amountExclTax = activeItems.reduce((sum, item) => sum + item.amountExclTax, 0);
    const yearlyAmounts = activeItems.reduce((values, item) => {
      Array.from({ length: 10 }, (_, index) => {
        values[index] = finiteMoney(values[index] + (item.fundingPlan?.yearlyAmounts[String(index + 1)] || 0));
      });
      return values;
    }, zeroYearlyAmounts());

    rows.push({
      side: firstMapping.side,
      ictSubjectCode: firstMapping.ictSubjectCode,
      ictSubjectName: firstMapping.ictSubjectName,
      subject,
	      originalAmount: getExistingAmount(fullState, subject),
	      quoteAmount: finiteMoney(amountInclTax),
	      writtenAmount: finiteMoney(amountInclTax),
	      amountExclTax: finiteMoney(amountExclTax),
	      originalYearlyAmounts: normalizeAnnualInclValues(
        normalizeSubjectFundingPlans(
          fullState.cashflowState?.assumptionsJson?.subjectFundingPlans
          || fullState.cashflowState?.assumptionsJson?.subject_funding_plans,
        )[createSubjectFundingPlanId({
          side: subject.side,
          groupId: subject.groupId,
          key: subject.key,
        })]?.annualInclValues,
      ),
      yearlyAmounts,
      sourceLineItemIds: items.map(item => item.id),
      sourceLineItemNames: items.map(item => item.name),
      fundingPlanModes: Array.from(new Set(items.flatMap(item =>
        item.fundingPlan?.mode ? [item.fundingPlan.mode] : []
      ))),
	      taxRate: amountInclTax > 0
	        ? deriveTaxRate(amountInclTax, amountExclTax)
	        : Number(activeItems[0]?.taxRate ?? items[0]?.taxRate ?? subject.defaultTaxRate),
      syncStatus: invalidItems.length > 0 ? "zeroed_error" : "ready",
      syncMessages,
    });
  });

  const currentSubjectKeys = new Set(groups.keys());
  collectPreviouslyControlledSubjects(calculated, fullState, projectId)
    .forEach((snapshot, snapshotKey) => {
      if (currentSubjectKeys.has(snapshotKey)) return;
      const subject = ICT_SUBJECT_DEFINITIONS.find(candidate =>
        candidate.side === snapshot.side && candidate.subjectCode === snapshot.ictSubjectCode
      );
      if (!subject) return;
      rows.push({
        side: snapshot.side,
        ictSubjectCode: snapshot.ictSubjectCode,
        ictSubjectName: subject.standardSubjectName,
	        subject,
	        originalAmount: getExistingAmount(fullState, subject),
	        quoteAmount: 0,
	        writtenAmount: 0,
	        amountExclTax: 0,
	        originalYearlyAmounts: normalizeAnnualInclValues(
          normalizeSubjectFundingPlans(
            fullState.cashflowState?.assumptionsJson?.subjectFundingPlans
            || fullState.cashflowState?.assumptionsJson?.subject_funding_plans,
          )[createSubjectFundingPlanId({
            side: subject.side,
            groupId: subject.groupId,
            key: subject.key,
          })]?.annualInclValues,
        ),
        yearlyAmounts: zeroYearlyAmounts(),
        sourceLineItemIds: [...snapshot.sourceLineItemIds],
        sourceLineItemNames: snapshot.sourceLineItemIds.map(id => itemMap.get(id)?.name || id),
        fundingPlanModes: [],
        taxRate: getExistingTaxRate(fullState, subject) || snapshot.taxRate,
        syncStatus: "released_mapping",
        syncMessages: ["该科目已取消映射或改映射到其他科目，本次将清零原金额和年度计划"],
      });
    });

  const mappedIds = new Set(calculated.mappings.filter(mapping => mapping.enabled).map(mapping => mapping.lineItemId));
  lineItems.forEach(item => {
    if (!item.enabled) skippedItems.push({ id: item.id, name: item.name, side: item.side, reason: "计算项已禁用" });
    else if (!item.outputEnabled) skippedItems.push({ id: item.id, name: item.name, side: item.side, reason: "未启用输出" });
    else if (!mappedIds.has(item.id)) skippedUnmappedItems.push({ id: item.id, name: item.name, side: item.side, reason: "未映射 ICT 科目" });
  });

  return {
    projectId,
    scenarioId: blueprint.scenarioId || blueprint.id,
    blueprintId: blueprint.id,
    projectYears: getAiComputeProjectCycleYears(calculated.parameters),
    discountRate: getAiComputeDiscountRateDecimal(calculated.parameters),
    rows,
    skippedUnmappedItems,
    skippedItems,
  };
}

export function buildAiComputeIctExportPayloads(
  preview: AiComputeIctExportPreview,
  fullState: UnknownRecord,
): AiComputeIctExportPayloads {
  const project = fullState.project || {};
  const lifecycle = fullState.lifecycleState || {};
  const cashflow = fullState.cashflowState || {};
  const inputPayloadJson = getBaseInput(fullState);
  const assumptionsJson = cloneRecord(cashflow.assumptionsJson);
  const existingPlans: SubjectFundingPlans = normalizeSubjectFundingPlans(
    assumptionsJson.subjectFundingPlans || assumptionsJson.subject_funding_plans,
  );
  const importedAt = new Date().toISOString();

  preview.rows
    .filter(row => row.syncStatus !== "paused_override" && row.syncStatus !== "merge_conflict")
    .forEach(row => {
    const sourceLineItems = row.sourceLineItemIds.flatMap(value => {
      const separator = value.indexOf(":");
      if (separator <= 0) return [];
      return [{
        amountSourceId: value.slice(0, separator),
        lineItemId: value.slice(separator + 1),
      }];
    });
    const amountSourceIds = Array.from(new Set(sourceLineItems.map(item => item.amountSourceId)));
    const isAggregate = preview.blueprintId === "intelligent-compute-aggregate";
    const importTrace: SubjectFundingPlanImportTrace = {
      source: isAggregate ? "intelligent_compute" : "ai_compute_quote",
      sourceLabel: isAggregate ? "来自智算金额来源" : "来自智算报价测算",
      projectId: preview.projectId,
      scenarioId: preview.scenarioId,
      blueprintId: preview.blueprintId,
      sourceLineItemIds: row.sourceLineItemIds,
      ...(isAggregate ? { amountSourceIds, sourceLineItems } : {}),
      importedAt,
    };
	    const existingAssumptionItem = getAssumptionSubjectItem(assumptionsJson, row.subject) || {};
	    const excl = row.amountExclTax;
    const nextItem = {
      ...existingAssumptionItem,
      incl: row.writtenAmount,
      excl,
      tax: row.taxRate,
      importTrace,
    };
    setAssumptionSubjectItem(assumptionsJson, row.subject, nextItem);
    inputPayloadJson[row.ictSubjectCode] = {
      ...(inputPayloadJson[row.ictSubjectCode] || {}),
      incl_tax: String(row.writtenAmount),
      excl_tax: String(excl),
      tax_rate: String(row.taxRate),
      import_trace: importTrace,
    };

    const planId = createSubjectFundingPlanId({
      side: row.subject.side,
      groupId: row.subject.groupId,
      key: row.subject.key,
    });
    const plan: SubjectFundingPlan = {
      id: planId,
      subjectRef: {
        side: row.subject.side,
        groupId: row.subject.groupId,
        key: row.subject.key,
      },
      mode: "custom",
      annualInclValues: normalizeAnnualInclValues(row.yearlyAmounts),
      enabled: true,
      source: isAggregate ? "intelligent_compute" : "ai_compute_quote",
      lastChangeReason: isAggregate ? "intelligent_compute_import" : "ai_compute_quote_import",
      lastChangedAt: importedAt,
      updatedAt: importedAt,
      importTrace,
    };
    existingPlans[planId] = plan;
  });

  const isAggregate = preview.blueprintId === "intelligent-compute-aggregate";
  const importRecord = {
    source: isAggregate ? "intelligent_compute" : "ai_compute_quote",
    sourceLabel: isAggregate ? "来自智算金额来源" : "来自智算报价测算",
    projectId: preview.projectId,
    scenarioId: preview.scenarioId,
    blueprintId: preview.blueprintId,
    importedAt,
    ictSubjectCodes: Array.from(new Set(
      preview.rows
        .map(row => row.ictSubjectCode),
    )),
    releasedSubjectCodes: Array.from(new Set(
      preview.rows
        .filter(row => row.syncStatus === "released_mapping")
        .map(row => row.ictSubjectCode),
    )),
    zeroedSubjectCodes: Array.from(new Set(
      preview.rows
        .filter(row => row.syncStatus === "zeroed_absent")
        .map(row => row.ictSubjectCode),
    )),
    amountSourceIds: isAggregate
      ? Array.from(new Set(preview.rows.flatMap(row =>
          row.sourceLineItemIds.map(id => id.split(":")[0]).filter(Boolean)
        )))
      : [],
  };
  assumptionsJson.subjectFundingPlans = existingPlans;
  delete assumptionsJson.subject_funding_plans;
  assumptionsJson.projectYears = preview.projectYears;
  assumptionsJson.discountRate = preview.discountRate;
  assumptionsJson.subjectFundingPlanMigrationVersion = SUBJECT_FUNDING_PLAN_MIGRATION_VERSION;
  assumptionsJson.cashflowCalculationSource = "subject_funding_plans";
  assumptionsJson.aiComputeQuoteImport = importRecord;
  inputPayloadJson.subject_funding_plans = existingPlans;
  inputPayloadJson.project_years = preview.projectYears;
  inputPayloadJson.discount_rate = String(preview.discountRate);
  inputPayloadJson.subject_funding_plan_migration_version = SUBJECT_FUNDING_PLAN_MIGRATION_VERSION;
  inputPayloadJson.cashflow_calculation_source = "subject_funding_plans";
  inputPayloadJson.ai_compute_quote_import = importRecord;

  return {
    lifecycleState: {
      profileJson: cloneRecord(lifecycle.profileJson || {
        projectName: project.name || "",
        customerName: project.customer_name || "",
      }),
      parametersJson: {
        ...cloneRecord(lifecycle.parametersJson),
        projectYears: preview.projectYears,
        discountRate: preview.discountRate,
        cashflowCalculationSource: "subject_funding_plans",
        subjectFundingPlanMigrationVersion: SUBJECT_FUNDING_PLAN_MIGRATION_VERSION,
        aiComputeQuoteImport: importRecord,
      },
      backgroundJson: cloneRecord(lifecycle.backgroundJson),
      inputPayloadJson,
    },
    cashflowState: {
      cashflowModel: cashflow.cashflowModel ?? project.cashflow_model ?? null,
      paymentModelJson: {
        ...cloneRecord(cashflow.paymentModelJson),
        cashflowCalculationSource: "subject_funding_plans",
        subjectFundingPlanMigrationVersion: SUBJECT_FUNDING_PLAN_MIGRATION_VERSION,
      },
      yearlyCashflowJson: {},
      sectorCashflowJson: cloneRecord(cashflow.sectorCashflowJson),
      assumptionsJson,
      metricsJson: {},
    },
  };
}
