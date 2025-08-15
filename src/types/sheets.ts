/*
 * Узкие интерфейсы для Google Sheets + type-guards
 */

export type SheetRange = string; // e.g. 'Sheet1!A1:D100'

export interface SheetValueRange {
  range: SheetRange;
  values: (string | number | null)[][]; // 2D массив в текстовом представлении
}

export interface BatchGetResult {
  valueRanges: SheetValueRange[];
}

export interface BatchUpdateRequest {
  valueInputOption: 'RAW' | 'USER_ENTERED';
  data: Array<{
    range: SheetRange;
    values: (string | number | null)[][];
  }>;
}

export function isSheetValueRange(v: unknown): v is SheetValueRange {
  if (!v || typeof v !== 'object') return false;
  const o = v as { range?: unknown; values?: unknown };
  if (typeof o.range !== 'string') return false;
  if (!Array.isArray(o.values)) return false;
  return o.values.every(row => Array.isArray(row));
}

export function isBatchGetResult(v: unknown): v is BatchGetResult {
  if (!v || typeof v !== 'object') return false;
  const o = v as { valueRanges?: unknown };
  if (!Array.isArray(o.valueRanges)) return false;
  return o.valueRanges.every(isSheetValueRange);
}
