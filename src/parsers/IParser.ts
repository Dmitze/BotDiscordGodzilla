import type { DriveFile } from '@/types/drive';

export type ParseInput = {
  fileId?: string;
  buffer?: Buffer;
  mime?: string;
};

export type ParseResult = {
  text: string;
  source: string; // 'export' | 'parser' | 'ocr' | 'raw' | custom
  warnings?: string[];
  buffer?: Buffer; // оригинальный буфер, если доступен (для checksum)
};

export interface IParser {
  canParse(meta: DriveFile): boolean;
  parse(input: ParseInput, ctx: {
    exportFile: (fileId: string, mime: string) => Promise<Buffer>;
    downloadFile: (fileId: string) => Promise<Buffer>;
    extractTextFromImage: (file: DriveFile) => Promise<string>;
    extractTextFromBuffer: (buf: Buffer) => Promise<string>;
  }): Promise<ParseResult>;
}
