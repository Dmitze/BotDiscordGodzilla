import type { DriveFile } from '@/types/drive';
import type { IParser, ParseInput, ParseResult } from './IParser';

export class ImageOcrParser implements IParser {
  canParse(meta: DriveFile): boolean {
    return !!meta.mimeType && /^image\//i.test(meta.mimeType);
  }
  async parse(input: ParseInput, ctx: {
    downloadFile: (fileId: string) => Promise<Buffer>;
    extractTextFromBuffer: (buf: Buffer) => Promise<string>;
  }): Promise<ParseResult> {
    if (!input.fileId) throw new Error('fileId required');
    const buf = await ctx.downloadFile(input.fileId);
    const text = await ctx.extractTextFromBuffer(buf);
    return { text, source: 'ocr', buffer: buf };
  }
}
