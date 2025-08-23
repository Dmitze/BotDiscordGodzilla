import type { SearchParams } from '@/types';
import type { SearchIndex, SearchQuery } from '@/search/SearchIndex';

export interface SqliteRunArgs {
  searchIndex: SearchIndex;
  params: SearchParams;
}

export async function runSearchSqlite({ searchIndex, params }: SqliteRunArgs) {
  const tags: string[] = [];
  if (params.documentType && params.documentType !== 'all') tags.push(params.documentType);
  if (params.unit) tags.push(params.unit);
  if (params.priority && params.priority !== 'all') tags.push(params.priority);
  const parseDate = (s?: string) => {
    if (!s) return undefined;
    const m = s.match(/(\d{4})-(\d{1,2})-(\d{1,2})|(?:(\d{1,2})\.(\d{1,2})\.(\d{4}))/);
    if (!m) return undefined;
    const y = Number(m[1] ?? m[6]);
    const mon = Number(m[2] ?? m[5]);
    const d = Number(m[3] ?? m[4]);
    const dt = new Date(y, mon - 1, d);
    return isNaN(dt.getTime()) ? undefined : dt.getTime();
  };
  const modifiedFrom = parseDate(params.dateFrom);
  const modifiedTo = parseDate(params.dateTo);

  const filters: any = {};
  if (typeof modifiedFrom === 'number') filters.modifiedFrom = modifiedFrom;
  if (typeof modifiedTo === 'number') filters.modifiedTo = modifiedTo;
  if (tags.length) filters.tags = tags;

  const q: SearchQuery = {
    text: String(params.query || ''),
    limit: Math.max(1, params.limit || 10),
    sample: undefined,
    // always include filters key (even if empty)
    filters,
  } as any;

  return searchIndex.search(q);
}
