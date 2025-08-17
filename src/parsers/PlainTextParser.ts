import type { DriveFile } from '@/types/drive';
import type { IParser, ParseInput, ParseResult } from './IParser';

export class PlainTextParser implements IParser {
  private mimes = new Set(['text/plain', 'text/markdown', 'application/json', 'text/csv']);
  canParse(meta: DriveFile): boolean {
    return !!meta.mimeType && this.mimes.has(meta.mimeType);
  }
  async parse(input: ParseInput, ctx: {
    downloadFile: (fileId: string) => Promise<Buffer>;
  }): Promise<ParseResult> {
    if (!input.fileId) throw new Error('fileId required');
    const buf = await ctx.downloadFile(input.fileId);
    return { text: buf.toString('utf8'), source: 'raw', buffer: buf };
  }
}
