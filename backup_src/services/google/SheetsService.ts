import type { MetricsService } from '../MetricsService';
import type { SheetRange, SheetValueRange, BatchGetResult, BatchUpdateRequest } from '@/types/sheets';
import type { SheetData } from '@/types';

/**
 * SheetsService: узкая логика конверсии структур Google Sheets
 * Без прямых сетевых вызовов. Парсинг/валидация/подготовка данных для фасада GoogleService.
 */
export class SheetsService {
  constructor(private readonly metrics?: MetricsService) {}

  public normalizeRange(range: string): SheetRange {
    // Минимальная нормализация: трим и замена пробелов
    return range.trim().replace(/\s+/g, ' ');
  }

  public parseBatchGet(raw: unknown): BatchGetResult {
    const start = Date.now();
    try {
      const res: BatchGetResult = { valueRanges: [] };
      if (!raw || typeof raw !== 'object') return res;
      const obj = raw as { valueRanges?: unknown };
      const arr = Array.isArray(obj.valueRanges) ? obj.valueRanges : [];
      res.valueRanges = arr.map(vr => this.toValueRangeSafe(vr)).filter((v): v is SheetValueRange => !!v);
      return res;
    } finally {
      const dur = Date.now() - start;
      try { this.metrics?.updateGoogleApiMetrics('sheets', 'parse_batch_get', 'ok', dur); } catch { /* noop: метрики не критичны */ }
    }
  }

  /**
   * Нормализация ответа spreadsheets.values.get к SheetData
   */
  public toSheetDataFromGet(raw: unknown, fallbackRange: string): SheetData {
    const start = Date.now();
    try {
      const obj = (raw && typeof raw === 'object' ? (raw as { range?: unknown; values?: unknown; majorDimension?: unknown }) : {});
      const range = typeof obj.range === 'string' ? this.normalizeRange(obj.range) : this.normalizeRange(fallbackRange);
      const valuesRaw = Array.isArray(obj.values) ? obj.values : [];
      const values = valuesRaw.map(row => Array.isArray(row) ? row.map(cell => (cell == null ? '' : String(cell))) : []);
      const majorDimension = 'ROWS';
      return { range, majorDimension, values };
    } finally {
      const dur = Date.now() - start;
      try { this.metrics?.updateGoogleApiMetrics('sheets', 'parse_get', 'ok', dur); } catch { /* noop: метрики не критичны */ }
    }
  }

  public buildBatchUpdate(req: BatchUpdateRequest): BatchUpdateRequest {
    // Нормализация всех ranges + значений и базовая валидация
    const normData = req.data.map(d => ({
      range: this.normalizeRange(d.range),
      values: this.normalizeWriteValues(d.values),
    }));
    this.validateBatchWrite(normData);
    return { valueInputOption: req.valueInputOption, data: normData };
  }

  // ----- helpers -----
  private toValueRangeSafe(v: unknown): SheetValueRange | null {
    if (!v || typeof v !== 'object') return null;
    const o = v as { range?: unknown; values?: unknown };
    const range = typeof o.range === 'string' ? this.normalizeRange(o.range) : undefined;
    const values = Array.isArray(o.values) ? o.values : [];
    if (!range) return null;
    const norm = values.map(row => Array.isArray(row) ? row.map(c => this.normalizeCell(c)) : []);
    return { range, values: norm };
  }

  private normalizeCell(c: unknown): string | number | null {
    if (c == null) return null;
    if (typeof c === 'number') return c;
    if (typeof c === 'string') return c;
    // Числа в строковом виде
    if (typeof c === 'object' && c !== null && 'toString' in c) return String(c);
    try { return JSON.stringify(c); } catch { return null; }
  }

  /** Нормализация значений для записи: undefined -> null, объекты -> JSON, boolean -> 'TRUE'/'FALSE' */
  public normalizeWriteValues(values: Array<Array<unknown>>): (string | number | null)[][] {
    return values.map(row =>
      Array.isArray(row)
        ? row.map(v => {
            if (v == null) return null;
            if (typeof v === 'number') return v;
            if (typeof v === 'string') return v;
            if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
            if (typeof v === 'object') {
              try { return JSON.stringify(v); } catch { return String(v); }
            }
            return String(v);
          })
        : []
    );
  }

  /** Простейшая валидация batch write: ограничение на количество ячеек и размеры строк */
  public validateBatchWrite(data: Array<{ range: string; values: (string | number | null)[][] }>, opts?: { maxCells?: number }): void {
    const maxCells = opts?.maxCells ?? 50000; // безопасный дефолт
    let total = 0;
    for (const item of data) {
      const rows = item.values.length;
      const cols = item.values.reduce((m, r) => Math.max(m, r.length), 0);
      total += rows * cols;
      if (total > maxCells) {
        throw new Error(`Превышен лимит ячеек для batch write: ${total} > ${maxCells}`);
      }
    }
  }
}
