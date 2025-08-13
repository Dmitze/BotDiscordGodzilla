export type Row = Record<string, any>;
export type Condition = {
  field: string;
  op: 'eq' | 'neq' | 'contains' | 'in' | 'gte' | 'lte';
  value: any;
};

export class AnalyticsService {
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
}
