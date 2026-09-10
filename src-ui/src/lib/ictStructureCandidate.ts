import { prepareIctTaxItemsInclBatch } from './ictTaxItemBatch';
import type { IctSubjectState } from './ictTaxItemEdit';
import type { LockedTotalStructureContext } from './ictReverseCalculation';

/** Resolve the same effective state the batch setter will write, including one
 * balancing correction. A candidate that still cannot close is never evaluated. */
export function prepareStructureCandidate(state: IctSubjectState, structure: LockedTotalStructureContext,
  target: number, balancing: number, autoFix: boolean, moneyEpsilon: number) {
  const prepare = (amount: number, targetInput = target) => prepareIctTaxItemsInclBatch(state, [
    { groupId: structure.targetSubject.groupId, key: structure.targetSubject.key, incl: targetInput, reason: 'reverse_calculation_sync' },
    { groupId: structure.balancingSubject.groupId, key: structure.balancingSubject.key, incl: amount, reason: 'balance_allocation_sync' },
  ], autoFix);
  const read = (prepared: ReturnType<typeof prepare>, subject: LockedTotalStructureContext['targetSubject']) =>
    subject.groupId === 'revNonItCt' ? prepared.state.revNonItCt.incl : prepared.state[subject.groupId][subject.key].incl;
  let prepared = prepare(balancing);
  let targetAmount = read(prepared, structure.targetSubject);
  let balancingAmount = read(prepared, structure.balancingSubject);
  const difference = () => Number((structure.totalInclAmount - structure.fixedOtherInclAmount - targetAmount - balancingAmount).toFixed(2));
  if (Math.abs(difference()) > moneyEpsilon) {
    prepared = prepare(Number((balancingAmount + difference()).toFixed(2)));
    targetAmount = read(prepared, structure.targetSubject);
    balancingAmount = read(prepared, structure.balancingSubject);
  }
  // The final setter receives effective amounts. Rebuild CT paired items from
  // those exact arguments too: their tax rates can differ from the revenue item.
  prepared = prepare(balancingAmount, targetAmount);
  targetAmount = read(prepared, structure.targetSubject);
  balancingAmount = read(prepared, structure.balancingSubject);
  const valid = Number.isFinite(targetAmount) && Number.isFinite(balancingAmount)
    && targetAmount >= 0 && balancingAmount >= 0 && Math.abs(difference()) <= moneyEpsilon;
  return { ...prepared, valid, targetAmount, balancingAmount,
    message: valid ? undefined : '结构反算候选经财务口径归一及承接补差后仍无法保持同侧含税总金额不变或出现负金额，已停止写入。' };
}
