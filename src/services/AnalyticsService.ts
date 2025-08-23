import type { BotConfig } from '@/types';
import type { CacheService } from './CacheService';
import { createHash } from 'crypto';

export type Row = Record<string, any>;
export type Condition = {
  field: string;
  op: 'eq' | 'neq' | 'contains' | 'in' | 'gte' | 'lte';
  value: any;
};

export class AnalyticsService {
  private readonly cache: CacheService | null;
  private readonly ttl: number;

  constructor(config?: BotConfig, cache?: CacheService) {
    this.cache = cache ?? null;
    this.ttl = config?.google?.analyticsCacheTTL ?? config?.performance?.cacheTTL ?? 900;
  }

  private keyOf(obj: unknown): string {
    const json = JSON.stringify(obj);
    return createHash('sha1').update(json).digest('hex');
  }

  private async withCache<T>(key: string, compute: () => Promise<T> | T): Promise<T> {
    if (!this.cache) {
      return await Promise.resolve(compute());
    }
    try {
      const cached = await this.cache.get<T>(key);
      if (cached !== null) return cached as T;
      const val = await Promise.resolve(compute());
      // best-effort cache set
      await this.cache.set(key, val, this.ttl).catch(() => undefined);
      return val;
    } catch {
      return await Promise.resolve(compute());
    }
  }

  inferSchema(rows: Row[]): string[] {
    if (!rows.length) return [];
    return Object.keys(rows[0] ?? {});
  }

  filterRows(rows: Row[], conditions: Condition[]): Row[] {
    const norm = (v: any) => (typeof v === 'string' ? v.toLowerCase() : v);
    const pass = (row: Row) =>
      conditions.every(c => {
        const val = row[c.field];
        switch (c.op) {
          case 'eq':
            return norm(val) === norm(c.value);
          case 'neq':
            return norm(val) !== norm(c.value);
          case 'contains':
            return String(val ?? '').toLowerCase().includes(String(c.value ?? '').toLowerCase());
          case 'in':
            return Array.isArray(c.value) && c.value.some((x: any) => norm(x) === norm(val));
          case 'gte':
            return Number(val) >= Number(c.value);
          case 'lte':
            return Number(val) <= Number(c.value);
          default:
            return true;
        }
      });
    return rows.filter(pass);
  }

  groupBy(rows: Row[], keys: string[]): Record<string, Row[]> {
    const map: Record<string, Row[]> = {};
    for (const r of rows) {
      const k = keys.map(k => String(r[k] ?? '')).join('|');
      if (!map[k]) map[k] = [];
      map[k].push(r);
    }
    return map;
  }

  aggregate(rows: Row[], field: string | null, op: 'count' | 'sum'): number {
    if (op === 'count') return rows.length;
    if (!field) return 0;
    return rows.reduce((acc, r) => acc + Number(r[field] ?? 0), 0);
  }

  /**
   * Выполнить типовой пайплайн: фильтр → группировка → агрегация, с кэшированием
   */
  async analyze(params: {
    rows: Row[];
    conditions?: Condition[];
    groupKeys?: string[];
    aggregateField?: string | null;
    aggregateOp?: 'count' | 'sum';
  }): Promise<{ groups: Record<string, Row[]>; metrics: Record<string, number> } | Row[]> {
    const { rows, conditions = [], groupKeys = [], aggregateField = null, aggregateOp = 'count' } =
      params;

    const cacheKey = `analytics:${this.keyOf({ conditions, groupKeys, aggregateField, aggregateOp })}`;
    return this.withCache(cacheKey, async () => {
      const filtered = conditions.length ? this.filterRows(rows, conditions) : rows;
      if (!groupKeys.length) return filtered;
      const grouped = this.groupBy(filtered, groupKeys);
      const metrics: Record<string, number> = {};
      for (const [k, arr] of Object.entries(grouped)) {
        metrics[k] = this.aggregate(arr, aggregateField, aggregateOp);
      }
      return { groups: grouped, metrics };
    }) as Promise<{ groups: Record<string, Row[]>; metrics: Record<string, number> } | Row[]>;
  }

  /**
   * Простой частотный анализ ключевых слов по тексту
   */
  extractKeywords(text: string, opts?: { minLen?: number; topN?: number; stop?: string[] }): Array<{ word: string; count: number }> {
    const minLen = opts?.minLen ?? 3;
    const topN = opts?.topN ?? 20;
    const stop = new Set((opts?.stop ?? ['the','and','a','to','of','в','і','та','на','що','для','це','як','ми','ви','вони'])
      .map(s => s.toLowerCase()));
    const words = String(text || '')
      .toLowerCase()
      .replace(/[^\p{L}\p{N}\s]+/gu, ' ')
      .split(/\s+/)
      .filter(w => w.length >= minLen && !stop.has(w));
    const freq = new Map<string, number>();
    for (const w of words) freq.set(w, (freq.get(w) ?? 0) + 1);
    return Array.from(freq.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, topN)
      .map(([word, count]) => ({ word, count }));
  }

  /**
   * Наивные темы: группируем ключевые слова по стемам (обрезка окончаний) и выбираем топовые
   */
  extractTopics(text: string, opts?: { topN?: number }): Array<{ topic: string; weight: number }> {
    const topN = opts?.topN ?? 10;
    const kw = this.extractKeywords(text, { topN: topN * 3 });
    const stem = (w: string) => w.replace(/(ами|ами|ами|ами|ів|ем|ам|ах|ах|ий|ий|ий|ий|ий|ий|ий|ов|ев|ів|ом|ом|ах|ах|и|и|и|а|я|у|ю|е|о)$/u, '');
    const buckets = new Map<string, number>();
    for (const { word, count } of kw) {
      const s = stem(word);
      buckets.set(s, (buckets.get(s) ?? 0) + count);
    }
    return Array.from(buckets.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, topN)
      .map(([topic, weight]) => ({ topic, weight }));
  }

  /**
   * Табличная аналитика: преднастроенные группировки/срезы
   */
  async analyzeTable(rows: Row[], preset: 'by_status' | 'by_month' | 'by_owner', statusField = 'status', dateField = 'date', ownerField = 'owner'): Promise<{ groups: Record<string, Row[]>; metrics: Record<string, number> }> {
    switch (preset) {
      case 'by_status':
        return (await this.analyze({ rows, groupKeys: [statusField], aggregateOp: 'count' })) as any;
      case 'by_month': {
        const norm = rows.map(r => ({ ...r, __month: this.monthKey(r[dateField]) }));
        return (await this.analyze({ rows: norm, groupKeys: ['__month'], aggregateOp: 'count' })) as any;
      }
      case 'by_owner':
        return (await this.analyze({ rows, groupKeys: [ownerField], aggregateOp: 'count' })) as any;
    }
  }

  private monthKey(v: any): string {
    const d = v ? new Date(v) : null;
    if (!d || Number.isNaN(d.getTime())) return 'unknown';
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    return `${y}-${m}`;
  }
}
