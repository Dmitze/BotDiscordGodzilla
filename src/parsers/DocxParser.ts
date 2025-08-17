import type { DriveFile } from '@/types/drive';
import type { IParser, ParseInput, ParseResult } from './IParser';
import * as mammoth from 'mammoth';

export class DocxParser implements IParser {
  canParse(meta: DriveFile): boolean {
    return meta.mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' || meta.mimeType === 'application/msword';
  }
  async parse(input: ParseInput, ctx: {
    downloadFile: (fileId: string) => Promise<Buffer>;
  }): Promise<ParseResult> {
    if (!input.fileId) throw new Error('fileId required');
    const buf = await ctx.downloadFile(input.fileId);
    const { value } = await mammoth.extractRawText({ buffer: buf });
    return { text: value || '', source: 'parser', buffer: buf };
  }
}
