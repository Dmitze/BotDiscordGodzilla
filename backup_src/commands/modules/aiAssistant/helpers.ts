import * as xlsx from 'xlsx';
import type { GoogleService } from '@/services/GoogleService';

export function tokenizeName(query: string, max = 5): string {
  const tokens = (query.match(/[\p{L}\p{N}\-_.]{2,}/giu) || [])
    .filter(w => w.length >= 2)
    .slice(0, max);
  return tokens.join(' ').trim();
}

export function findMonthNumber(q: string): number | undefined {
  const monthMap: Record<string, number> = {
    'январ': 1, 'лют': 2, 'фев': 2, 'берез': 3, 'март': 3, 'квіт': 4, 'апрел': 4,
    'май': 5, 'трав': 5, 'июн': 6, 'черв': 6, 'июл': 7, 'лип': 7, 'авг': 8, 'серп': 8,
    'сен': 9, 'верес': 9, 'окт': 10, 'жовт': 10, 'нояб': 11, 'листоп': 11, 'дек': 12, 'груд': 12,
  };
  const key = Object.keys(monthMap).find(k => q.includes(k));
  return key ? monthMap[key] : undefined;
}

export function isImageMime(mt?: string): boolean {
  return !!(mt && /^image\//i.test(mt));
}

export function isDocLikeMime(mt?: string): boolean {
  return mt === 'application/vnd.google-apps.document' ||
    mt === 'application/pdf' ||
    mt === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
    mt === 'application/msword';
}

export async function ensureDriveIndex(googleService: GoogleService, folderId: string) {
  let index = await googleService.getDriveIndex(folderId);
  if (!index) index = await googleService.buildDriveIndex(folderId, { ttlSeconds: 1800, recursive: true, maxDepth: -1 });
  return index;
}

export async function readGoogleSheet(googleService: GoogleService, spreadsheetId: string, range = 'A1:Z2000'): Promise<Array<Record<string, unknown>>> {
  try {
    const sheetTitles = await googleService.listSheets(spreadsheetId);
    const first = sheetTitles[0] || 'Лист1';
    const data = await googleService.getSheetData(spreadsheetId, `${first}!${range}`);
    const rows = (data.values || []) as unknown[];
    if (!rows.length) return [];
    const headerRow = (rows[0] as unknown[] | undefined) ?? [];
    const rest = (rows.slice(1) as unknown[][]) ?? [];
    const headers = headerRow.map(h => String(h ?? '').trim());
    return rest.map((rowArr) => {
      const obj: Record<string, unknown> = {};
      headers.forEach((h, i) => { if (!h) return; obj[h] = (rowArr)[i]; });
      return obj;
    });
  } catch { return []; }
}

export function readExcelBuffer(buf: Buffer): Array<Record<string, unknown>> {
  try {
    const wb = xlsx.read(buf, { type: 'buffer' });
    const firstName = wb.SheetNames[0];
    if (!firstName) return [];
    const sheet = wb.Sheets[firstName];
    if (!sheet) return [];
    return xlsx.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: '' });
  } catch { return []; }
}
