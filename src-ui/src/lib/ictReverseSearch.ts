export type ReverseTargetType = "margin" | "npv_rate";

export type ReverseMetricProbe = {
  amount: number;
  metricValue: number;
};

export const buildCostReverseFeasibilityProbeAmounts = (
  targetType: ReverseTargetType,
): number[] => targetType === "npv_rate" ? [0, 0.01] : [0];

export const selectHighestMetricProbe = (
  probes: ReverseMetricProbe[],
): ReverseMetricProbe | null => probes.reduce<ReverseMetricProbe | null>((best, probe) => {
  if (!Number.isFinite(probe.amount) || !Number.isFinite(probe.metricValue)) return best;
  if (!best || probe.metricValue > best.metricValue) return probe;
  if (probe.metricValue === best.metricValue && probe.amount < best.amount) return probe;
  return best;
}, null);
