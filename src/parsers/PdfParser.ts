import type { DriveFile } from '@/types/drive';
import type { IParser, ParseInput, ParseResult } from './IParser';
import pdfParse from 'pdf-parse';

export class PdfParser implements IParser {
  canParse(meta: DriveFile): boolean {
    return meta.mimeType === 'application/pdf';
  }
  async parse(input: ParseInput, ctx: {
    downloadFile: (fileId: string) => Promise<Buffer>;
  }): Promise<ParseResult> {
    if (!input.fileId) throw new Error('fileId required');
    const buf = await ctx.downloadFile(input.fileId);
    const parsed = await pdfParse(buf);
    return { text: parsed.text || '', source: 'parser', buffer: buf };
  }
}
