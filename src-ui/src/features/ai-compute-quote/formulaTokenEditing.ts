import type { AiComputeQuoteFormulaToken } from "./types";

export function clampFormulaCursor(cursor: number, tokenCount: number) {
  if (!Number.isFinite(cursor)) return tokenCount;
  return Math.min(Math.max(Math.trunc(cursor), 0), tokenCount);
}

export function insertFormulaTokensAt(
  tokens: AiComputeQuoteFormulaToken[],
  cursor: number,
  insertedTokens: AiComputeQuoteFormulaToken[],
) {
  const safeCursor = clampFormulaCursor(cursor, tokens.length);
  return {
    tokens: [
      ...tokens.slice(0, safeCursor),
      ...insertedTokens,
      ...tokens.slice(safeCursor),
    ],
    cursor: safeCursor + insertedTokens.length,
  };
}

export function removeFormulaTokenAt(
  tokens: AiComputeQuoteFormulaToken[],
  tokenIndex: number,
  cursor: number,
) {
  if (tokenIndex < 0 || tokenIndex >= tokens.length) {
    return { tokens, cursor: clampFormulaCursor(cursor, tokens.length) };
  }
  const nextTokens = tokens.filter((_, index) => index !== tokenIndex);
  const nextCursor = tokenIndex < cursor ? cursor - 1 : cursor;
  return {
    tokens: nextTokens,
    cursor: clampFormulaCursor(nextCursor, nextTokens.length),
  };
}

export function removeFormulaTokenBeforeCursor(
  tokens: AiComputeQuoteFormulaToken[],
  cursor: number,
) {
  const safeCursor = clampFormulaCursor(cursor, tokens.length);
  if (safeCursor === 0) return { tokens, cursor: 0 };
  return removeFormulaTokenAt(tokens, safeCursor - 1, safeCursor);
}
