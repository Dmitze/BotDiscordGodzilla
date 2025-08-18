import type { DriveFile } from '@/types/drive';
import type { IParser, ParseInput, ParseResult } from './IParser';
import pdfParse from 'pdf-parse';

export class PdfParser implements IParser {
  canParse(meta: DriveFile): boolean {
    return meta.mimeType === 'application/pdf';
  }
  async parse(input: ParseInput, ctx: {
    downloadFile: (fileId: string) => Promise<Buffer>;
    extractTextFromBuffer: (buf: Buffer) => Promise<string>;
  }): Promise<ParseResult> {
    if (!input.fileId) throw new Error('fileId required');
    const buf = await ctx.downloadFile(input.fileId);
    try {
      const parsed = await pdfParse(buf);
      return { text: parsed.text || '', source: 'parser', buffer: buf };
    } catch (_e) {
      // Fallback to OCR if pdf-parse fails
      const text = await ctx.extractTextFromBuffer(buf);
      return { text: text || '', source: 'ocr', buffer: buf };
    }
  }
}
