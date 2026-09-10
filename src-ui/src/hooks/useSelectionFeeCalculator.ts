import { useCallback, useEffect, useRef, useState } from 'react';
import { invoke } from '@tauri-apps/api/core';
import { DEFAULT_SELECTION_FEE_TARGET_SUBJECT_CODE, normalizeSelectionFeeTargetSubjectCode } from '../lib/selectionFee';

type Anchor = 'quote' | 'limit';
interface Result {
  quote: string;
  final_limit: string;
  actual_cost: string;
  selection_fee_excl: string;
  selection_fee_incl: string;
  quote_candidates: string[];
}
const text = (value: unknown) => String(value ?? '').trim();
const defaults = () => ({ quote: '', limit: '', markup: '50', anchor: 'quote' as Anchor,
  target: DEFAULT_SELECTION_FEE_TARGET_SUBJECT_CODE, merge: false });

/** A result belongs to one anchor/input key. Restored derived amounts are never trusted. */
export function useSelectionFeeCalculator() {
  const [input, setInput] = useState(defaults);
  const revision = useRef(0);
  const [calculation, setCalculation] = useState<{ key: string; result?: Result; error?: string }>({ key: '' });
  const anchorValue = input[input.anchor];
  const key = JSON.stringify([input.anchor, anchorValue, input.markup, revision.current]);
  const current = calculation.key === key ? calculation : undefined;
  const result = current?.result;
  const selQuote = input.anchor === 'quote' ? input.quote : result?.quote ?? '';
  const selLimit = input.anchor === 'limit' ? input.limit : result?.final_limit ?? '';
  const selectionFeePending = Boolean(anchorValue) && !current;

  useEffect(() => {
    if (!anchorValue) return;
    let active = true;
    const atRevision = revision.current;
    const command = input.anchor === 'quote' ? 'calculate_selection_fee' : 'reverse_calculate_selection_fee';
    const args = input.anchor === 'quote' ? { quote: anchorValue, markup: input.markup || '0' } : { limit: anchorValue, markup: input.markup || '0' };
    void invoke<Result>(command, args).then(result => {
      if (!active || atRevision !== revision.current) return;
      if (!result || typeof result.selection_fee_incl !== 'string' || typeof result.selection_fee_excl !== 'string'
        || !Array.isArray(result.quote_candidates)) {
        setCalculation({ key, error: '甄选费计算接口版本不匹配，请重启更新后的应用再试。' });
        return;
      }
      setCalculation({ key, result });
    }).catch(error => {
      if (active && atRevision === revision.current) setCalculation({ key, error: String(error) });
    });
    return () => { active = false; };
  }, [anchorValue, input.anchor, input.markup, key]);

  const handleSelFeeChange = (type: Anchor | 'markup', value: string) => {
    revision.current += 1;
    setInput(previous => ({ ...previous, quote: selQuote, limit: selLimit, [type]: value,
      anchor: type === 'markup' ? previous.anchor : type }));
  };
  const restoreSelectionFeeState = useCallback((payload?: Record<string, unknown> | null) => {
    revision.current += 1;
    setCalculation({ key: '' });
    setInput({ quote: text(payload?.selection_fee_quote), limit: text(payload?.selection_fee_limit),
      markup: payload && Object.prototype.hasOwnProperty.call(payload, 'selection_fee_markup') ? text(payload.selection_fee_markup) : '50',
      anchor: payload?.selection_fee_anchor === 'limit' ? 'limit' : 'quote',
      target: normalizeSelectionFeeTargetSubjectCode(payload?.selection_fee_target_subject_code),
      merge: payload?.selection_fee_merge_service === true });
  }, []);
  const buildSelectionFeePayload = () => {
    const hasData = Boolean(anchorValue) || input.merge || input.target !== DEFAULT_SELECTION_FEE_TARGET_SUBJECT_CODE;
    return hasData ? {
      selection_fee_quote: selQuote, selection_fee_markup: input.markup,
      selection_fee_actual_cost: result?.actual_cost ?? '', selection_fee_amount: result?.selection_fee_incl ?? '',
      selection_fee_limit: selLimit, selection_fee_anchor: input.anchor,
      selection_fee_target_subject_code: input.target, selection_fee_merge_service: input.merge,
    } : {};
  };
  return {
    selQuote, selLimit, selMarkup: input.markup, selActualCost: result?.actual_cost ?? '',
    selFee: result?.selection_fee_incl ?? '', selFeeExcl: result?.selection_fee_excl ?? '',
    selectionFeeAnchor: input.anchor,
    setSelectionFeeAnchor: (anchor: Anchor) => handleSelFeeChange(anchor, anchor === 'quote' ? selQuote : selLimit),
    selectionFeeTargetSubjectCode: input.target,
    setSelectionFeeTargetSubjectCode: (target: string) => setInput(previous => ({ ...previous, target: normalizeSelectionFeeTargetSubjectCode(target) })),
    selectionFeeMergeService: input.merge,
    setSelectionFeeMergeService: (merge: boolean) => setInput(previous => ({ ...previous, merge })),
    selectionFeeError: current?.error ?? '', selectionFeePending, selectionFeeReady: Boolean(result),
    selectionFeeNotice: (result?.quote_candidates.length ?? 0) > 1
      ? `该限价有多个有效报价：${result!.quote_candidates.join('、')} 元；当前采用较低报价，可改用报价正算指定。` : '',
    handleSelFeeChange, restoreSelectionFeeState, buildSelectionFeePayload,
  };
}
