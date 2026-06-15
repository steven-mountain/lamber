import AppIcon from "../../components/icons/AppIcon";
import { Button } from "../../components/ui/button";
import type { IctResult } from "../../utils/projectService";
import type { AiComputeQuoteSummary } from "./types";

type SyncStatus = "idle" | "syncing" | "synced" | "error" | "conflict";

type Props = {
  summary: AiComputeQuoteSummary;
  ictResult: IctResult | null;
  syncStatus: SyncStatus;
  majorRevenue: string;
  majorCost: string;
  outputCount: number;
  syncDetailsAvailable: boolean;
  onShowOutput: () => void;
  onShowSyncDetails: () => void;
};

function formatWan(value: number) {
  return (value / 10000).toLocaleString("zh-CN", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function formatNumber(value: number, digits = 2) {
  return value.toLocaleString("zh-CN", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
  });
}

function ResultMetric({
  label,
  value,
  tone,
}: {
  label: string;
  value: string;
  tone?: "success" | "danger";
}) {
  const toneClass = tone === "success"
    ? "text-success-foreground"
    : tone === "danger"
      ? "text-destructive"
      : "text-foreground";

  return (
    <div className="flex items-baseline justify-between gap-3 rounded-lg bg-muted/55 px-3 py-2.5">
      <span className="text-caption font-semibold text-secondary-foreground">{label}</span>
      <span className={`numeric-value text-right text-sm font-extrabold ${toneClass}`}>{value}</span>
    </div>
  );
}

function AmountMetric({
  label,
  value,
  tone,
}: {
  label: string;
  value: string;
  tone: "revenue" | "cost";
}) {
  return (
    <div className={`rounded-lg px-3 py-3 ${
      tone === "revenue" ? "bg-success-soft/70" : "bg-warning-soft/70"
    }`}>
      <div className="text-caption font-semibold text-secondary-foreground">{label}</div>
      <div className={`numeric-value mt-1 text-lg font-extrabold ${
        tone === "revenue" ? "text-success-foreground" : "text-warning-foreground"
      }`}>
        {value}
      </div>
    </div>
  );
}

export default function BenefitConclusionSidebar({
  summary,
  ictResult,
  syncStatus,
  majorRevenue,
  majorCost,
  outputCount,
  syncDetailsAvailable,
  onShowOutput,
  onShowSyncDetails,
}: Props) {
  const margin = ictResult ? Number(ictResult.margin_rate) : Number.NaN;
  const npv = ictResult ? Number(ictResult.npv) : Number.NaN;
  const conclusion = !ictResult
    ? {
        label: "待同步",
        description: "完成 ICT 同步后显示正式效益判断。",
        tone: "text-secondary-foreground",
        surface: "bg-muted",
        icon: "info" as const,
      }
    : npv < 0 || margin < 0
      ? {
          label: "风险较高",
          description: "存在负收益指标，请复核收入、成本与资金计划。",
          tone: "text-destructive",
          surface: "bg-destructive-soft",
          icon: "warning" as const,
        }
      : margin < 0.08
        ? {
            label: "需关注",
            description: "收益空间较窄，建议重点检查主要成本构成。",
            tone: "text-warning-foreground",
            surface: "bg-warning-soft",
            icon: "warning" as const,
          }
        : {
            label: "收益良好",
            description: "当前正式 ICT 指标处于正向区间。",
            tone: "text-success-foreground",
            surface: "bg-success-soft",
            icon: "success" as const,
          };

  return (
    <aside className="min-h-0 overflow-y-auto rounded-xl bg-card p-4 shadow-sm xl:w-[292px]">
      <div className="mb-4">
        <h2 className="text-section-title">效益结论</h2>
        <p className="text-caption text-secondary-foreground">只读展示正式 ICT 测算结果</p>
      </div>

      <section className={`rounded-lg p-3.5 ${conclusion.surface}`}>
        <div className="flex items-start gap-2.5">
          <AppIcon name={conclusion.icon} className={conclusion.tone} />
          <div>
            <div className={`text-base font-extrabold ${conclusion.tone}`}>{conclusion.label}</div>
            <p className="mt-1 text-caption leading-relaxed text-secondary-foreground">{conclusion.description}</p>
          </div>
        </div>
      </section>

      <section className="mt-4">
        <h3 className="mb-2 text-caption font-extrabold text-foreground">核心金额</h3>
        <div className="grid grid-cols-2 gap-2 xl:grid-cols-1 2xl:grid-cols-2">
          <AmountMetric label="总收入" value={`${formatWan(summary.totalRevenue)} 万元`} tone="revenue" />
          <AmountMetric label="总成本" value={`${formatWan(summary.totalCost)} 万元`} tone="cost" />
        </div>
      </section>

      <section className="mt-4">
        <h3 className="mb-2 text-caption font-extrabold text-foreground">收益指标</h3>
        <div className="space-y-2">
          <ResultMetric
            label="ICT 毛利率"
            value={ictResult ? `${formatNumber(Number(ictResult.margin_rate) * 100)}%` : "--"}
            tone={ictResult && Number(ictResult.margin_rate) >= 0 ? "success" : ictResult ? "danger" : undefined}
          />
          <ResultMetric
            label="净现值率"
            value={ictResult ? `${formatNumber(Number(ictResult.npv_rate) * 100)}%` : "--"}
            tone={ictResult && Number(ictResult.npv_rate) >= 0 ? "success" : ictResult ? "danger" : undefined}
          />
          <ResultMetric
            label="ICT NPV"
            value={ictResult ? `${formatWan(Number(ictResult.npv))} 万元` : "--"}
            tone={ictResult && Number(ictResult.npv) >= 0 ? "success" : ictResult ? "danger" : undefined}
          />
          <ResultMetric label="动态回收期" value={ictResult ? `${ictResult.dynamic_payback} 年` : "--"} />
        </div>
      </section>

      <section className="mt-4">
        <h3 className="mb-2 text-caption font-extrabold text-foreground">关键构成</h3>
        <div className="space-y-2">
          <div className="rounded-lg bg-success-soft/45 px-3 py-2.5">
            <div className="text-[11px] font-semibold text-secondary-foreground">主要收入</div>
            <div className="mt-0.5 truncate text-sm font-bold text-foreground" title={majorRevenue}>{majorRevenue}</div>
          </div>
          <div className="rounded-lg bg-warning-soft/45 px-3 py-2.5">
            <div className="text-[11px] font-semibold text-secondary-foreground">主要成本</div>
            <div className="mt-0.5 truncate text-sm font-bold text-foreground" title={majorCost}>{majorCost}</div>
          </div>
        </div>
      </section>

      <section className="mt-4 grid gap-2">
        <Button variant="secondary" onClick={onShowOutput}>
          <AppIcon name="document" />
          查看输出包
        </Button>
        <Button variant="outline" disabled={!syncDetailsAvailable} onClick={onShowSyncDetails}>
          <AppIcon name="tableProperties" />
          查看同步明细
        </Button>
      </section>

      <div className="mt-4 rounded-lg bg-muted/55 p-3 text-caption leading-relaxed text-secondary-foreground">
        输出 ICT 科目 <span className="numeric-value font-extrabold text-foreground">{outputCount}</span> 项。
        <div className="mt-1">
          {syncStatus === "syncing"
            ? "正在使用 ICT 正式参数重新计算。"
            : syncStatus === "error"
              ? "同步失败，当前指标可能不是最新结果。"
              : syncStatus === "conflict"
                ? "存在 ICT 人工修改冲突，请查看同步明细。"
                : "正式效益指标来自 ICT 测算引擎。"}
        </div>
      </div>
    </aside>
  );
}
