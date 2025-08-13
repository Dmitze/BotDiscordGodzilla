import type { BotConfig } from '@/types';
import { CacheService } from './CacheService';
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
}
