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
  const annualCashflow = buildAnnualCashflowFromSubjectFundingPlans(subjects, plans);
  return {
    coverage,
    annualCashflow,
    fields: {
      rev_cashflow_excl: coverage.valid
        ? annualCashflow.annualRevenueExcl.map(moneyString)
        : null,
      cost_cashflow_excl: coverage.valid
        ? annualCashflow.annualCostExcl.map(moneyString)
        : null,
      it_rev_cashflow_excl: coverage.valid
        ? annualCashflow.annualItRevenueExcl.map(moneyString)
        : null,
      it_cost_cashflow_excl: coverage.valid
        ? annualCashflow.annualItCostExcl.map(moneyString)
        : null,
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
