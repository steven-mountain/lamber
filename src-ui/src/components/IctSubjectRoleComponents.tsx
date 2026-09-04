import React, { useState, useEffect, useRef } from "react";
import AppIcon from "./icons/AppIcon";
import {
  type IctSubjectDefinition,
  type IctTaxItemLike,
  getSubjectExcelDisplayName,
} from "../lib/ictSubjectCatalog";
import {
  type BalanceAllocationState,
  type BalanceAllocationSide,
  type BalanceAllocationRule,
  isBalanceSubjectMatch,
  getBalanceSubjectRef,
} from "../lib/ictBalanceAllocation";
import {
  getReverseSubjectRef,
  getReverseSubjectRefKey,
} from "../lib/ictReverseCalculation";

interface SubjectRoleActionsProps {
  subject: IctSubjectDefinition;
  item: IctTaxItemLike | null;
  balanceAllocation: BalanceAllocationState;
  updateBalanceRule: (side: BalanceAllocationSide, patch: Partial<BalanceAllocationRule>) => void;
  updateTaxItem: (groupId: string, key: string, field: "incl" | "tax" | "excl", val: number) => void;
  revSubjectRefKey: string;
  setRevSubjectRefKey: (key: string) => void;
  setRevMode: (mode: "cost" | "revenue") => void;
  subjects: Array<{ subject: IctSubjectDefinition; item: IctTaxItemLike | null }>;
}

export const SubjectRoleActions: React.FC<SubjectRoleActionsProps> = ({
  subject,
  item,
  balanceAllocation,
  updateBalanceRule,
  updateTaxItem,
  revSubjectRefKey,
  setRevSubjectRefKey,
  setRevMode,
  subjects,
}) => {
  const [isOpen, setIsOpen] = useState(false);
  const menuRef = useRef<HTMLDivElement>(null);

  const balanceSide: BalanceAllocationSide = subject.side === "revenue" ? "revenue" : "investment";
  const rule = balanceAllocation[balanceSide];
  const isBalancing = rule.enabled && isBalanceSubjectMatch(rule.balancingSubject, subject);
  const isReverseTarget = revSubjectRefKey === getReverseSubjectRefKey(getReverseSubjectRef(subject));

  const displayName = getSubjectExcelDisplayName(subject, item);

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (menuRef.current && !menuRef.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    };
    if (isOpen) {
      document.addEventListener("mousedown", handleClickOutside);
    }
    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
    };
  }, [isOpen]);

  const onSetBalancing = () => {
    const doSet = () => {
      const existingRef = rule.balancingSubject;
      if (existingRef && !isBalanceSubjectMatch(existingRef, subject)) {
        const existingSubjectRow = subjects.find(row => isBalanceSubjectMatch(existingRef, row.subject));
        const oldDisplayName = existingSubjectRow
          ? getSubjectExcelDisplayName(existingSubjectRow.subject, existingSubjectRow.item)
          : "原科目";
        const confirmMsg = `${subject.side === 'revenue' ? '收入' : '投入'}侧已有差额承接科目“${oldDisplayName}”。\n切换后将由“${displayName}”自动承接剩余金额，并重新计算当前配置的金额分配。是否继续？`;
        if (!window.confirm(confirmMsg)) return;
      }
      updateBalanceRule(balanceSide, {
        balancingSubject: getBalanceSubjectRef(subject),
        enabled: true,
      });
    };

    if (isReverseTarget) {
      const confirmMsg = `“${displayName}”当前为智能反算目标。\n将其设置为差额承接科目后，将清除其反算目标状态。是否继续？`;
      if (window.confirm(confirmMsg)) {
        setRevSubjectRefKey("");
        doSet();
      }
    } else {
      doSet();
    }
  };

  const onClearBalancing = () => {
    updateBalanceRule(balanceSide, {
      balancingSubject: null,
      enabled: rule.totalInclAmount !== null,
    });
  };

  const onSetReverse = () => {
    if (isBalancing) {
      const confirmMsg = `“${displayName}”当前为差额承接科目。\n将其设置为智能反算目标后，将取消其差额承接角色，并将该科目金额清零。是否继续？`;
      if (!window.confirm(confirmMsg)) return;

      // 1. Cancel balancing role
      updateBalanceRule(balanceSide, {
        balancingSubject: null,
        enabled: rule.totalInclAmount !== null,
      });

      // 2. Clear amount
      updateTaxItem(subject.groupId, subject.key, "incl", 0);
    }
    setRevSubjectRefKey(getReverseSubjectRefKey(getReverseSubjectRef(subject)));
    setRevMode(subject.side);
  };

  const onClearReverse = () => {
    setRevSubjectRefKey("");
  };

  return (
    <div className="relative inline-flex items-center gap-1.5 shrink-0" ref={menuRef}>
      {isBalancing && (
        <span className="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-bold bg-amber-50 text-amber-700 border border-amber-200/30">
          差额承接
        </span>
      )}
      {isReverseTarget && (
        <span className="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-bold bg-blue-50 text-blue-700 border border-blue-200/30">
          反算目标
        </span>
      )}

      <button
        type="button"
        onClick={() => setIsOpen(!isOpen)}
        className="text-[11px] font-bold text-primary/80 hover:text-primary px-1.5 py-0.5 rounded hover:bg-primary/5 transition-colors"
      >
        {isBalancing || isReverseTarget ? "更改" : "设置角色"}
      </button>

      {isOpen && (
        <div className="absolute left-0 mt-6 w-48 bg-card border border-border rounded-lg shadow-lg z-50 p-1 flex flex-col gap-0.5 text-xs font-semibold backdrop-blur-md bg-opacity-95">
          {isBalancing ? (
            <button
              type="button"
              onClick={() => {
                onClearBalancing();
                setIsOpen(false);
              }}
              className="w-full text-left px-2.5 py-1.5 text-foreground hover:bg-muted rounded-md transition-colors flex items-center gap-1.5"
            >
              <AppIcon name="delete" size={12} className="text-amber-700" />
              取消差额承接
            </button>
          ) : (
            <button
              type="button"
              onClick={() => {
                onSetBalancing();
                setIsOpen(false);
              }}
              className="w-full text-left px-2.5 py-1.5 text-foreground hover:bg-muted rounded-md transition-colors flex items-center gap-1.5"
            >
              <AppIcon name="quickAction" size={12} className="text-amber-600" />
              设为差额承接科目
            </button>
          )}

          {isReverseTarget ? (
            <button
              type="button"
              onClick={() => {
                onClearReverse();
                setIsOpen(false);
              }}
              className="w-full text-left px-2.5 py-1.5 text-foreground hover:bg-muted rounded-md transition-colors flex items-center gap-1.5"
            >
              <AppIcon name="delete" size={12} className="text-blue-700" />
              取消反算目标
            </button>
          ) : (
            <button
              type="button"
              onClick={() => {
                onSetReverse();
                setIsOpen(false);
              }}
              className="w-full text-left px-2.5 py-1.5 text-foreground hover:bg-muted rounded-md transition-colors flex items-center gap-1.5"
            >
              <AppIcon name="reverse" size={12} className="text-blue-600" />
              设为智能反算目标
            </button>
          )}
        </div>
      )}
    </div>
  );
};

interface SelectedSubjectRoleSummaryProps {
  subject: IctSubjectDefinition;
  item: IctTaxItemLike | null;
  onLocate: () => void;
  onClear: () => void;
}

export const SelectedSubjectRoleSummary: React.FC<SelectedSubjectRoleSummaryProps> = ({
  subject,
  item,
  onLocate,
  onClear,
}) => {
  const displayName = getSubjectExcelDisplayName(subject, item);

  return (
    <div className="flex items-center justify-between bg-card border border-border/40 rounded-lg px-3 py-1.5 shadow-sm gap-2">
      <span className="text-xs font-semibold text-foreground truncate max-w-[150px] sm:max-w-[200px]" title={displayName}>
        {displayName}
      </span>
      <div className="flex items-center gap-1 shrink-0">
        <button
          type="button"
          onClick={onLocate}
          className="text-[11px] font-bold text-primary hover:bg-primary/5 px-2 py-1 rounded transition-colors"
        >
          定位
        </button>
        <button
          type="button"
          onClick={onClear}
          className="text-[11px] font-bold text-destructive hover:bg-destructive/5 px-2 py-1 rounded transition-colors"
        >
          清除
        </button>
      </div>
    </div>
  );
};
