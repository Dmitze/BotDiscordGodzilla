import type { DriveFile } from '@/types/drive';
import type { IParser, ParseInput, ParseResult } from './IParser';

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
