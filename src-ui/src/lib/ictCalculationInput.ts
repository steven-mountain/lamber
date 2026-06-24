import { ICT_SUBJECT_DEFINITIONS } from "./ictSubjectCatalog";
import {
  buildAnnualCashflowFromSubjectFundingPlans,
  normalizeSubjectFundingPlans,
  validateSubjectFundingPlanCoverage,
  type SubjectFundingPlanCoverageSubject,
  type SubjectFundingPlans,
} from "./ictSubjectFundingPlan";
import type { IctInput } from "../utils/projectService";

type UnknownRecord = Record<string, any>;

const finiteNumber = (value: unknown) => {
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
};

const moneyString = (value: number) => (
  Number.isFinite(value) ? value : 0
).toFixed(2);

export function buildIctFundingSubjectsFromInput(input: UnknownRecord): SubjectFundingPlanCoverageSubject[] {
	  return ICT_SUBJECT_DEFINITIONS.map(subject => {
	    const item = input[subject.subjectCode] || {};
	    const rawTaxRate = item.tax_rate ?? item.tax;
	    return {
	      subjectRef: {
	        side: subject.side,
        groupId: subject.groupId,
        key: subject.key,
	      },
	      displayName: subject.standardSubjectName,
	      subjectAmountIncl: finiteNumber(item.incl_tax ?? item.incl),
	      taxRate: rawTaxRate === undefined || rawTaxRate === null || rawTaxRate === ""
	        ? subject.defaultTaxRate
	        : finiteNumber(rawTaxRate),
	      isItScope: subject.groupId === "revIt" || subject.groupId === "costIt",
	    };
	  });
}

export function buildIctFundingCashflowFields(
  subjects: SubjectFundingPlanCoverageSubject[],
  plans: SubjectFundingPlans,
) {
  const coverage = validateSubjectFundingPlanCoverage(subjects, plans);
  // Always derive the yearly cashflow from whatever plans ARE maintained.
  // Subjects without a maintained plan fall back to a first-year (upfront)
  // payment so that one un-maintained subject no longer forces the entire
  // cashflow — including the IT breakdown — back to the legacy "all in year 1"
  // model. Maintained multi-year / proportional plans are honored as-is.
  const annualCashflow = buildAnnualCashflowFromSubjectFundingPlans(subjects, plans, {
    fallbackUnmaintainedToUpfront: true,
  });
  return {
    coverage,
    annualCashflow,
    fields: {
      rev_cashflow_excl: annualCashflow.annualRevenueExcl.map(moneyString),
      cost_cashflow_excl: annualCashflow.annualCostExcl.map(moneyString),
      it_rev_cashflow_excl: annualCashflow.annualItRevenueExcl.map(moneyString),
      it_cost_cashflow_excl: annualCashflow.annualItCostExcl.map(moneyString),
    },
  };
}

export function finalizeIctInputWithFundingPlans(
  input: UnknownRecord,
  rawPlans: unknown,
): {
  input: IctInput;
  plans: SubjectFundingPlans;
  coverage: ReturnType<typeof validateSubjectFundingPlanCoverage>;
  annualCashflow: ReturnType<typeof buildAnnualCashflowFromSubjectFundingPlans>;
} {
  const plans = normalizeSubjectFundingPlans(rawPlans);
  const subjects = buildIctFundingSubjectsFromInput(input);
  const finalized = buildIctFundingCashflowFields(subjects, plans);
  return {
    input: {
      ...input,
      cashflow_calculation_source: "subject_funding_plans",
      subject_funding_plans: plans,
      ...finalized.fields,
    } as unknown as IctInput,
    plans,
    coverage: finalized.coverage,
    annualCashflow: finalized.annualCashflow,
  };
}
