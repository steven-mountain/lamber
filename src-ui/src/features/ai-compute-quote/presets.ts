import type {
  AiComputeQuoteBlueprint,
  AiComputeQuoteFormulaToken,
  AiComputeQuoteLineItem,
  AiComputeQuoteParameter,
  AiComputeQuoteSubjectMapping,
} from "./types";
import { DEFAULT_KEY_PARAMETER_KEYS, DEFAULT_PARAMETER_GROUPS, PARAMETER_GROUP_IDS } from "./parameterLayout";

const parameter = (
  id: string,
  name: string,
  key: string,
  value: number,
  unit: string,
  category: AiComputeQuoteParameter["category"],
  groupId: string,
): AiComputeQuoteParameter => ({
  id,
  name,
  key,
  value,
  unit,
  category,
  groupId,
  isKey: DEFAULT_KEY_PARAMETER_KEYS.has(key),
  sensitivityEnabled: true,
});

const parameterToken = (parameterId: string, name: string): AiComputeQuoteFormulaToken => ({
  type: "parameter",
  id: parameterId,
  name,
});

const constantToken = (value: number): AiComputeQuoteFormulaToken => ({
  type: "constant",
  value,
});

const lineItem = (
  id: string,
  side: AiComputeQuoteLineItem["side"],
  name: string,
  parameters: Array<[id: string, name: string]>,
  constants: number[] = [],
  taxRate = 6,
): AiComputeQuoteLineItem => ({
  id,
  side,
  name,
  formula: {
    version: 2,
    tokens: [
      ...parameters.map(([id, name]) => parameterToken(id, name)),
      ...constants.map(constantToken),
    ].flatMap((token, index) => [
        ...(index > 0 ? [{ type: "operator", operator: "*" } as AiComputeQuoteFormulaToken] : []),
        token,
      ]),
  },
  amountInclTax: 0,
  amountExclTax: 0,
  taxRate,
  enabled: true,
  outputEnabled: true,
  fundingPlan: {
    enabled: true,
    mode: "first_year",
    yearlyAmounts: Object.fromEntries(Array.from({ length: 10 }, (_, index) => [String(index + 1), 0])),
  },
});

const mapping = (
  lineItemId: string,
  side: AiComputeQuoteSubjectMapping["side"],
  ictSubjectCode: string,
  ictSubjectName: string,
): AiComputeQuoteSubjectMapping => ({
  id: `mapping-${lineItemId}`,
  lineItemId,
  side,
  ictSubjectCode,
  ictSubjectName,
  enabled: true,
});

export function createH200Blueprint(): AiComputeQuoteBlueprint {
  return {
    id: "h200-standard",
    scenarioId: "h200-standard-scenario",
    name: "H200 标准智算金额来源",
    description: "64 台 H200、5 年服务期的标准报价预设。金额口径为元、含税。",
    parameterGroups: DEFAULT_PARAMETER_GROUPS.map(group => ({ ...group })),
    parameters: [
      parameter("device-count", "设备数量", "device_count", 64, "台", "scale", PARAMETER_GROUP_IDS.scale),
      {
        ...parameter("years", "项目周期", "years", 5, "年", "scale", PARAMETER_GROUP_IDS.scale),
        locked: true,
      },
      {
        ...parameter("discount-rate", "项目折现率", "discount_rate", 5.5, "%", "finance", PARAMETER_GROUP_IDS.finance),
        locked: true,
      },
      parameter("gpu-service-price", "GPU 服务单价", "gpu_service_price", 90000, "元/台/月", "price", PARAMETER_GROUP_IDS.pricing),
      parameter("bandwidth-revenue-price", "带宽收入单价", "bandwidth_revenue_price", 4000, "元/月/G", "price", PARAMETER_GROUP_IDS.pricing),
      parameter("cabinet-revenue-price", "机柜收入单价", "cabinet_revenue_price", 650, "元/kW/月", "price", PARAMETER_GROUP_IDS.pricing),
      parameter("machine-price", "单台机器价格", "machine_price", 3300000, "元/台", "cost", PARAMETER_GROUP_IDS.investment),
      parameter("network-price", "组网费用", "network_price", 300000, "元/台", "cost", PARAMETER_GROUP_IDS.investment),
      parameter("cabinet-cost-price", "机柜成本单价", "cabinet_cost_price", 440, "元/kW/月", "cost", PARAMETER_GROUP_IDS.investment),
      parameter("bandwidth-cost-price", "带宽成本单价", "bandwidth_cost_price", 4500, "元/月/G", "cost", PARAMETER_GROUP_IDS.investment),
      parameter("power-per-device", "功耗", "power_kw_per_device", 10.625, "kW/台", "technical", PARAMETER_GROUP_IDS.operations),
      parameter("maintenance-price", "维保费用", "maintenance_price", 30000, "元/台/年", "cost", PARAMETER_GROUP_IDS.operations),
      parameter("bandwidth-per-device", "单台带宽", "bandwidth_per_device", 5, "G/台", "technical", PARAMETER_GROUP_IDS.operations),
      parameter("capital-rate", "资金成本率", "capital_rate", 10, "%", "finance", PARAMETER_GROUP_IDS.finance),
    ],
    revenueItems: [
      lineItem("revenue-gpu-service", "revenue", "GPU 服务收入", [["gpu-service-price", "GPU 服务单价"], ["device-count", "设备数量"], ["years", "项目周期"]], [12]),
      lineItem("revenue-cabinet", "revenue", "机柜收入", [["cabinet-revenue-price", "机柜收入单价"], ["power-per-device", "功耗"], ["device-count", "设备数量"], ["years", "项目周期"]], [12]),
      lineItem("revenue-bandwidth", "revenue", "带宽收入", [["bandwidth-revenue-price", "带宽收入单价"], ["bandwidth-per-device", "单台带宽"], ["device-count", "设备数量"], ["years", "项目周期"]], [12]),
    ],
    costItems: [
      lineItem("cost-machine", "cost", "机器成本", [["machine-price", "单台机器价格"], ["device-count", "设备数量"]], [], 13),
      lineItem("cost-maintenance", "cost", "维保成本", [["maintenance-price", "维保费用"], ["device-count", "设备数量"], ["years", "项目周期"]]),
      lineItem("cost-network", "cost", "组网成本", [["network-price", "组网费用"], ["device-count", "设备数量"]], [], 13),
      lineItem("cost-cabinet", "cost", "机柜成本", [["cabinet-cost-price", "机柜成本单价"], ["power-per-device", "功耗"], ["device-count", "设备数量"], ["years", "项目周期"]], [12]),
      lineItem("cost-bandwidth", "cost", "带宽成本", [["bandwidth-cost-price", "带宽成本单价"], ["bandwidth-per-device", "单台带宽"], ["device-count", "设备数量"], ["years", "项目周期"]], [12]),
      {
        ...lineItem("cost-capital", "cost", "资金成本", [], []),
        formula: {
          version: 2,
          tokens: [
            { type: "function", name: "SUM" },
            { type: "line_item", id: "cost-machine", name: "机器成本" },
            { type: "comma" },
            { type: "line_item", id: "cost-maintenance", name: "维保成本" },
            { type: "comma" },
            { type: "line_item", id: "cost-network", name: "组网成本" },
            { type: "right_parenthesis" },
            { type: "operator", operator: "*" },
            { type: "parameter", id: "capital-rate", name: "资金成本率" },
            { type: "operator", operator: "*" },
            { type: "constant", value: 0.01 },
          ],
        },
      },
    ],
    mappings: [
      mapping("revenue-gpu-service", "revenue", "rev_it_cloud", "移动云-定制化收入"),
      mapping("revenue-cabinet", "revenue", "rev_it_other", "其他收入"),
      mapping("revenue-bandwidth", "revenue", "rev_ct_line", "专线收入"),
      mapping("cost-machine", "cost", "cost_it_device", "主要设备/甲供材料"),
      mapping("cost-maintenance", "cost", "cost_it_maintenance", "维护费用"),
      mapping("cost-network", "cost", "cost_it_integration", "集成服务"),
      mapping("cost-cabinet", "cost", "cost_it_running", "其他运行支出（电费等）"),
      mapping("cost-bandwidth", "cost", "cost_ct_bandwidth", "专线带宽成本"),
      mapping("cost-capital", "cost", "cost_mix_other", "其他管理费用等"),
    ],
  };
}
