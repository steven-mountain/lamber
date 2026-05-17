export type CashflowModel = 'model_a' | 'model_b' | 'model_c' | 'model_d';

export const cashflowModelLabels: Record<CashflowModel, string> = {
  model_a: '模型 A：100% 第一年度收付',
  model_b: '模型 B：按周期等额收付',
  model_c: '模型 C：首年 95%，末年 5%',
  model_d: '模型 D：高级自定义分配',
};

export function normalizeProjectYears(projectYears: number, years = 10): number {
  return Math.max(1, Math.min(Number.isFinite(projectYears) ? Math.trunc(projectYears) : 1, years));
}

export function normalizeDistribution(dist: number[], years = 10): number[] {
  const result = Array(years).fill(0);

  for (let i = 0; i < Math.min(dist.length, years); i++) {
    const value = Number(dist[i]);
    result[i] = Number.isFinite(value) && value > 0 ? value : 0;
  }

  const sum = result.reduce((a, b) => a + b, 0);

  if (sum <= 0) {
    return Array.from({ length: years }, (_, idx) => (idx === 0 ? 1 : 0));
  }

  return result.map(v => v / sum);
}

export function buildDistributionFromModel(
  model: CashflowModel,
  projectYears: number,
  customDist?: number[]
): number[] {
  const years = 10;
  const n = normalizeProjectYears(projectYears, years);
  const dist = Array(years).fill(0);

  switch (model) {
    case 'model_a':
      dist[0] = 1;
      return dist;

    case 'model_b':
      for (let i = 0; i < n; i++) {
        dist[i] = 1 / n;
      }
      return normalizeDistribution(dist, years);

    case 'model_c':
      if (n === 1) {
        dist[0] = 1;
      } else {
        dist[0] = 0.95;
        dist[n - 1] += 0.05;
      }
      return normalizeDistribution(dist, years);

    case 'model_d':
      for (let i = 0; i < n; i++) {
        dist[i] = customDist?.[i] ?? 0;
      }
      return normalizeDistribution(dist, years);

    default:
      dist[0] = 1;
      return dist;
  }
}

export function formatDistribution(dist: number[]): string {
  const normalized = normalizeDistribution(dist, 10);

  return normalized
    .map((v, idx) => ({ year: idx + 1, ratio: v }))
    .filter(item => item.ratio > 0)
    .map(item => `第${item.year}年 ${(item.ratio * 100).toFixed(2)}%`)
    .join('，');
}
