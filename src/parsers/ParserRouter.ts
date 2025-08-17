import type { DriveFile } from '@/types/drive';
import type { IParser, ParseResult } from './IParser';
import { t } from '@/i18n';
import logger from '@/utils/logger';

export type ParserContext = {
  exportFile: (fileId: string, mime: string) => Promise<Buffer>;
  downloadFile: (fileId: string) => Promise<Buffer>;
  extractTextFromImage: (file: DriveFile) => Promise<string>;
  extractTextFromBuffer: (buf: Buffer) => Promise<string>;
};

export class ParserRouter {
  private parsers: IParser[] = [];
  constructor(parsers: IParser[]) {
    this.parsers = parsers;
  }

  register(parser: IParser): void {
    this.parsers.push(parser);
  }

  /**
   * Выбирает парсер по метаданным; при ошибке логирует и пробует следующий.
   * Локализация ключей: parsers.fallback.used, parsers.unsupportedMime, parsers.parseError
   */
  async parse(meta: DriveFile, ctx: ParserContext): Promise<ParseResult> {
    const candidates = this.parsers.filter(p => p.canParse(meta));
    if (!candidates.length) {
      const msg = t('parsers.unsupportedMime', { mime: meta.mimeType || 'unknown' }) || `Unsupported MIME: ${meta.mimeType}`;
      logger.warn('ParserRouter: unsupported mime', { mime: meta.mimeType, name: meta.name });
      throw new Error(msg);
    }

    let lastErr: unknown;
    for (const parser of candidates) {
      try {
        return await parser.parse({ fileId: meta.id, mime: meta.mimeType }, ctx);
      } catch (e) {
        lastErr = e;
        logger.warn('ParserRouter: primary parser failed, trying fallback', {
          message: t('parsers.fallback.used') || 'Fallback parser used',
          parser: parser.constructor?.name,
          mime: meta.mimeType,
          error: e instanceof Error ? e.message : String(e),
        });
      }
    }

    const errMsg = t('parsers.parseError') || 'Parsing error';
    throw new Error(`${errMsg}: ${lastErr instanceof Error ? lastErr.message : String(lastErr)}`);
  }
}
