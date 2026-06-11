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
  sensitivityEnabled?: boolean;
  locked?: boolean;
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
  name: string;
  description?: string;
  parameters: AiComputeQuoteParameter[];
  revenueItems: AiComputeQuoteLineItem[];
  costItems: AiComputeQuoteLineItem[];
  mappings: AiComputeQuoteSubjectMapping[];
}

export interface AiComputeQuoteOutputSubjectAmount {
  side: AiComputeQuoteSide;
  ictSubjectCode: string;
  ictSubjectName: string;
  amountInclTax: number;
  amountExclTax: number;
  sourceLineItemIds: string[];
}

export interface AiComputeQuoteSummary {
  totalRevenue: number;
  totalCost: number;
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
  version: 1;
  blueprint: AiComputeQuoteBlueprint;
  savedAt: string;
}
