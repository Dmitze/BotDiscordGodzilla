import type { DriveFile } from '@/types/drive';
import type { IParser, ParseResult } from './IParser';
import { t } from '@/i18n';
import logger from '@/utils/logger';
import { withTimeout, retry } from './utils';
import { normalizeUnicode } from './normalize';

export type ParserContext = {
  exportFile: (fileId: string, mime: string) => Promise<Buffer>;
  downloadFile: (fileId: string) => Promise<Buffer>;
  extractTextFromImage: (file: DriveFile) => Promise<string>;
  extractTextFromBuffer: (buf: Buffer) => Promise<string>;
};

export type ParserRouterOptions = {
  timeoutMs?: number; // таймаут на один парсер
  retryAttempts?: number; // количество попыток на один парсер
  retryDelayMs?: number; // задержка между попытками
};

export class ParserRouter {
  private parsers: IParser[] = [];
  private opts: Required<ParserRouterOptions>;
  constructor(parsers: IParser[], options: ParserRouterOptions = {}) {
    this.parsers = parsers;
    this.opts = {
      timeoutMs: options.timeoutMs ?? 10000,
      retryAttempts: options.retryAttempts ?? 1,
      retryDelayMs: options.retryDelayMs ?? 200,
    };
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
        const result = await retry(
          () => withTimeout(parser.parse({ fileId: meta.id, mime: meta.mimeType }, ctx), this.opts.timeoutMs, 'parser.parse'),
          this.opts.retryAttempts,
          this.opts.retryDelayMs,
          'parser.parse'
        );
        // Нормализация текста единообразно по всем парсерам
        if (result && typeof result.text === 'string') {
          result.text = normalizeUnicode(result.text);
        }
        return result;
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
