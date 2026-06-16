export type AiComputeQuoteSide = "revenue" | "cost";

export type AiComputeQuoteParameterCategory =
  | "scale"
  | "price"
  | "cost"
  | "finance"
  | "technical"
  | "custom";

export interface AiComputeQuoteParameter {
  id: string;
  name: string;
  key: string;
  value: number;
  unit?: string;
  category?: AiComputeQuoteParameterCategory;
  groupId?: string;
  isKey?: boolean;
  sensitivityEnabled?: boolean;
  locked?: boolean;
}

export interface AiComputeQuoteParameterGroup {
  id: string;
  name: string;
  description?: string;
  builtin: boolean;
}

export type AiComputeQuoteFormulaOperand =
  | { type: "parameter"; parameterId: string }
  | { type: "constant"; value: number };

export type AiComputeQuoteLegacyFormula = {
  type: "multiply" | "add" | "subtract" | "divide";
  operands: AiComputeQuoteFormulaOperand[];
};

export type AiComputeQuoteFormulaToken =
  | { type: "parameter"; id: string; name: string }
  | { type: "line_item"; id: string; name: string }
  | { type: "constant"; value: number }
  | { type: "operator"; operator: "+" | "-" | "*" | "/" }
  | { type: "left_parenthesis" }
  | { type: "right_parenthesis" }
  | { type: "function"; name: "SUM" }
  | { type: "comma" };

export type AiComputeQuoteExpressionFormula = {
  version: 2;
  tokens: AiComputeQuoteFormulaToken[];
};

export type AiComputeQuoteFormula = AiComputeQuoteLegacyFormula | AiComputeQuoteExpressionFormula;

export type AiComputeLineItemFundingPlanMode = "first_year" | "even" | "manual";

export type AiComputeLineItemFundingPlan = {
  enabled: boolean;
  mode: AiComputeLineItemFundingPlanMode;
  yearlyAmounts: Record<string, number>;
};

export type AiComputeFormulaControlStatus = "formula" | "ict_override" | "merge_conflict";

export type AiComputeIctOverride = {
  ictSubjectCode: string;
  amountInclTax: number;
  taxRate: number;
  yearlyAmounts: number[];
  modifiedAt: string;
};

export interface AiComputeQuoteLineItem {
  id: string;
  side: AiComputeQuoteSide;
  name: string;
  formula: AiComputeQuoteFormula;
  amountInclTax: number;
  amountExclTax: number;
  taxRate?: number;
  enabled: boolean;
  outputEnabled?: boolean;
  calculationStatus?: "valid" | "incomplete" | "error";
  calculationError?: string;
  calculationWarnings?: string[];
  fundingPlan?: AiComputeLineItemFundingPlan;
  formulaControlStatus?: AiComputeFormulaControlStatus;
  ictOverride?: AiComputeIctOverride;
  ictControlMessage?: string;
}

export interface AiComputeQuoteSubjectMapping {
  id: string;
  lineItemId: string;
  side: AiComputeQuoteSide;
  ictSubjectCode: string;
  ictSubjectName: string;
  enabled: boolean;
}

export interface AiComputeQuoteBlueprint {
  id: string;
  scenarioId?: string;
  name: string;
  description?: string;
  parameterGroups: AiComputeQuoteParameterGroup[];
  parameters: AiComputeQuoteParameter[];
  revenueItems: AiComputeQuoteLineItem[];
  costItems: AiComputeQuoteLineItem[];
  mappings: AiComputeQuoteSubjectMapping[];
  syncState?: AiComputeQuoteSyncState;
}

export type AiComputeSyncedSubjectSnapshot = {
  side: AiComputeQuoteSide;
  ictSubjectCode: string;
  amountInclTax: number;
  taxRate: number;
  yearlyAmounts: number[];
  sourceLineItemIds: string[];
};

export type AiComputeQuoteSyncState = {
  revision: number;
  status: "idle" | "syncing" | "synced" | "error" | "conflict";
  syncedAt?: string;
  error?: string;
  subjects: Record<string, AiComputeSyncedSubjectSnapshot>;
};

export interface AiComputeQuoteOutputSubjectAmount {
  side: AiComputeQuoteSide;
  ictSubjectCode: string;
  ictSubjectName: string;
  amountInclTax: number;
  amountExclTax: number;
  sourceLineItemIds: string[];
}

export interface AiComputeOutputSubjectFundingPlan {
  side: AiComputeQuoteSide;
  ictSubjectCode: string;
  ictSubjectName: string;
  totalAmount: number;
  yearlyAmounts: Record<string, number>;
  sourceLineItemIds: string[];
}

export interface AiComputeQuoteSummary {
  totalRevenue: number;
  totalCost: number;
  totalRevenueExclTax: number;
  totalCostExclTax: number;
  grossProfit: number;
  grossMarginRate: number;
  costPerDeviceMonth: number;
}

export interface AiComputeQuoteSensitivityConfig {
  parameterId: string;
  min: number;
  max: number;
  step: number;
}

export interface AiComputeQuoteSensitivityRow extends AiComputeQuoteSummary {
  parameterValue: number;
}

export interface FormulaEvaluationResult {
  value: number;
  errors: string[];
  warnings: string[];
  status: "valid" | "incomplete" | "error";
}

export interface AiComputeQuotePersistedState {
  version: 1 | 2 | 3 | 4;
  blueprint: AiComputeQuoteBlueprint;
  savedAt: string;
}

export interface IntelligentComputeCalculationSnapshot {
  summary?: AiComputeQuoteSummary;
  syncState?: AiComputeQuoteSyncState;
  calculatedAt?: string;
  [key: string]: unknown;
}

export interface IntelligentAmountSource {
  id: string;
  projectId: string;
  name: string;
  description?: string | null;
  enabled: boolean;
  sourceVersion: number;
  metadata: Record<string, unknown>;
  parameterGroups: AiComputeQuoteParameterGroup[];
  parameters: AiComputeQuoteParameter[];
  revenueItems: AiComputeQuoteLineItem[];
  costItems: AiComputeQuoteLineItem[];
  mappings: AiComputeQuoteSubjectMapping[];
  calculationSnapshot: IntelligentComputeCalculationSnapshot;
  createdAt: string;
  updatedAt: string;
}

export interface IntelligentComputeProjectState {
  projectId: string;
  stateVersion: number;
  activeAmountSourceId?: string | null;
  projectYears: number;
  discountRate: number;
  syncRevision: number;
  controlledSubjects: Record<string, AiComputeSyncedSubjectSnapshot>;
  lastResult: Record<string, unknown>;
  createdAt: string;
  updatedAt: string;
}

export interface IntelligentComputeProjectData {
  state: IntelligentComputeProjectState;
  amountSources: IntelligentAmountSource[];
}
