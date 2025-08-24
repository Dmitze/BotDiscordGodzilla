import type { DriveFile } from '@/types/drive';
import type { IParser, ParseInput, ParseResult } from './IParser';
import type { GoogleService } from '@/services/GoogleService';

export class SheetsParser implements IParser {
  canParse(meta: DriveFile): boolean {
    return meta.mimeType === 'application/vnd.google-apps.spreadsheet';
  }
  async parse(input: ParseInput, ctx: {
    exportFile: (fileId: string, mime: string) => Promise<Buffer>;
  }): Promise<ParseResult> {
    if (!input.fileId) throw new Error('fileId required');
    const buf = await ctx.exportFile(input.fileId, 'text/csv');
    const text = buf.toString('utf8');
    return { text, source: 'export', buffer: buf };
  }
}

// Stable Sheets helpers (delegating to GoogleService)
export async function listSheets(service: GoogleService, spreadsheetId: string): Promise<string[]> {
  return service.listSheets(spreadsheetId);
}

export async function findSheetByName(
  service: GoogleService,
  spreadsheetId: string,
  name: string
): Promise<{ title: string; index: number } | null> {
  return service.findSheetByName(spreadsheetId, name);
}

export async function readRange(
  service: GoogleService,
  spreadsheetId: string,
  sheetName: string,
  rangeOrOpts: string | { columnHints?: string[]; headerRow?: number }
): Promise<{ headers: string[]; rows: (string | number | null)[][] }> {
  return service.readRange(spreadsheetId, sheetName, rangeOrOpts);
}
