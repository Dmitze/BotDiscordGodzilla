/**
 * DriveIndexerService
 * Индексация содержимого файлов Google Drive для простого полнотекстового поиска
 */

import type { BotConfig, HealthStatus, ServiceStats } from '@/types';
import type { DriveFile, DriveListResult, DriveListQuery } from '@/types/drive';
import logger from '@/utils/logger';
import { BaseService } from '@/core/BaseService';
import type { GoogleService } from './GoogleService';
import type { SearchIndex } from '@/search/SearchIndex';
import type { CacheService } from './CacheService';
import type SchedulerService from './SchedulerService';
// import { chunkTextForDiscord } from '@/utils/chunk';
import RetryManager from '@/utils/retry';
import { detectLanguage } from '@/nlp/LanguageDetector';

interface BotLike {
  config: BotConfig;
  getService(name: string): any;
}

export interface DriveIndexEntry {
  id: string;
  name: string;
  mimeType: string;
  modifiedTime?: string;
  owners?: string[];
  size?: number;
  text: string; // обрезанный текст
  textLength: number; // исходная длина текста
  updatedAt: number; // epoch ms
}

export interface DriveSearchResult {
  file: Omit<DriveIndexEntry, 'text'> & { snippet: string };
  score: number;
}

const INDEX_PREFIX = 'drive:index:file:'; // per-file запись
const INDEX_KEYS = 'drive:index:keys'; // список fileIds
const MAX_TEXT_STORED = 100_000; // ограничим объём для Redis

export class DriveIndexerService extends BaseService {
  private bot: BotLike;
  private google!: GoogleService;
  private cache!: CacheService;
  private searchIndex: SearchIndex | undefined;
  private metrics?: { incCounter?: (...args: any[]) => void; observeHistogram?: (...args: any[]) => void } | undefined;
  private indexedCount = 0;
  private lastRunAt: number | null = null;
  // simple in-memory queue for batching with limited concurrency
  private queue: DriveFile[] = [];
  private running = 0;
  private queueDrained?: Promise<void>;
  private queueResolve?: () => void;
  
  public initializeServices(
    googleService: GoogleService,
    cacheService: CacheService,
    searchIndex: SearchIndex | undefined,
    metricsService: { incCounter?: (...args: any[]) => void; observeHistogram?: (...args: any[]) => void } | undefined
  ): void {
    this.google = googleService;
    this.cache = cacheService;
    this.searchIndex = searchIndex;
    this.metrics = metricsService;
  }

  private isCronDisabled(): boolean {
    return (
      process.env['NODE_ENV'] === 'test' ||
      String(process.env['DISABLE_CRON']).toLowerCase() === 'true'
    );
  }

  // === queue helpers ===
  private resetQueue(concurrency: number, onSuccess?: () => void) {
    this.queue = [];
    this.running = 0;
    this.queueDrained = new Promise<void>(res => { this.queueResolve = res; });
    // attach runner
    const runNext = () => {
      if (!this.queue.length && this.running === 0) {
        this.queueResolve?.();
        return;
      }
      while (this.running < concurrency && this.queue.length) {
        const file = this.queue.shift()!;
        this.running++;
        this.indexOneFileByMeta(file)
          .then(() => { onSuccess?.(); })
          .catch(err => {
            logger.warn('⚠️ Завдання індексації не виконано', { id: file.id, error: err instanceof Error ? err.message : String(err) });
          })
          .finally(() => {
            this.running--;
            // schedule next tick
            setImmediate(runNext);
          });
      }
    };
    // start pump soon
    setImmediate(runNext);
  }

  private enqueue(file: DriveFile) {
    this.queue.push(file);
  }

  private async waitQueueDrain(): Promise<void> {
    await this.queueDrained;
  }

  constructor(bot: BotLike) {
    super('DriveIndexerService', bot.config);
    this.bot = bot;
  }

  protected async onInitialize(): Promise<void> {
    // Получим зависимости
    this.google = this.bot.getService('google') as GoogleService;
    this.cache = this.bot.getService('cache') as CacheService;
    try {
      this.searchIndex = this.bot.getService('searchIndex') as unknown as SearchIndex;
    } catch {
      this.searchIndex = undefined;
    }
    // Метрики опционально
    try {
      const m = this.bot.getService('metrics');
      if (m) this.metrics = m;
    } catch {
      // ignore if metrics not registered
    }

    if (!this.config.drive?.enableTextIndex) {
      logger.info('🔎 Индексация Drive отключена (DRIVE_ENABLE_TEXT_INDEX=false)');
      return;
    }

    // Зарегистрируем cron-задачу (пропускаємо у тестовому середовищі або коли DISABLE_CRON=true)
    if (this.isCronDisabled()) {
      logger.debug('⏭️ Пропуск реєстрації cron індексації Drive (тест або DISABLE_CRON)');
      return;
    }
    const scheduler = this.bot.getService('scheduler') as SchedulerService | undefined;
    if (scheduler && typeof scheduler.scheduleJob === 'function') {
      const cron = this.config.drive.indexCron || '*/30 * * * *';
      scheduler.scheduleJob('drive-index', cron, async () => {
        try {
          await this.reindexIncremental();
        } catch (e) {
          logger.error('❌ Ошибка фоновой индексации', { error: e instanceof Error ? e.message : String(e) });
        }
      });
      logger.info(`⏰ Задача индексации Drive зарегистрирована: ${cron}`);
    } else {
      logger.warn('⚠️ SchedulerService недоступен — cron для индексации не будет запущен');
    }

    // Выполняем начальную индексацию при запуске
    try {
      logger.info('🚀 Запуск начальной индексации Google Drive...');
      await this.reindexIncremental();
      logger.info('✅ Начальная индексация Google Drive завершена успешно');
    } catch (error) {
      logger.error('❌ Ошибка начальной индексации Google Drive', {
        error: error instanceof Error ? error.message : String(error)
      });
    }

    // формальный await, чтобы удовлетворить линтер (async без await)
    await Promise.resolve();
  }
  
  /** Получить чанки текста, безопасные для Discord */
  public async getTextChunks(fileId: string, max = 1900): Promise<string[]> {
    const text = await this.getText(fileId);
    if (!text) return [];
    
    // Розбиваємо текст на чанки з кращим форматуванням
    const chunks: string[] = [];
    let currentChunk = '';
    
    // Розбиваємо текст на рядки
    const lines = text.split('\n');
    
    for (const line of lines) {
      // Якщо додавання рядка перевищить ліміт, зберігаємо поточний чанк
      if (currentChunk.length + line.length + 1 > max) {
        if (currentChunk) {
          chunks.push(currentChunk.trim());
          currentChunk = '';
        }
        // Якщо один рядок більший за ліміт, розбиваємо його
        if (line.length > max) {
          const words = line.split(' ');
          let currentLine = '';
          
          for (const word of words) {
            if (currentLine.length + word.length + 1 > max) {
              if (currentLine) {
                chunks.push(currentLine.trim());
                currentLine = '';
              }
              // Якщо одне слово більше за ліміт, обрізаємо його
              if (word.length > max) {
                chunks.push(word.substring(0, max - 3) + '...');
                continue;
              }
            }
            currentLine += (currentLine ? ' ' : '') + word;
          }
          
          if (currentLine) {
            currentChunk = currentLine;
          }
        } else {
          currentChunk = line;
        }
      } else {
        currentChunk += (currentChunk ? '\n' : '') + line;
      }
    }
    
    // Додаємо останній чанк, якщо він не порожній
    if (currentChunk) {
      chunks.push(currentChunk.trim());
    }
    
    return chunks;
  }

  /** Переиндексация всех файлов (полная) */
  public async reindexAll(folderId?: string): Promise<void> {
    if (!this.ensureReady()) return;
    const fid = folderId || this.config.google.driveFolderId || this.config.drive.folderId;
    if (!fid) throw new Error('Не указан folderId для индексации');

    logger.info('📚 Запуск полной индексации Drive', { folderId: fid });
    const start = Date.now();
    this.metrics?.incCounter?.('drive_index_runs_total', { mode: 'full' });
    let pageToken: string | undefined = undefined;
    let total = 0;

    const configuredConcAll = Number(this.config.drive?.maxConcurrency ?? 4);
    const concurrency = process.env['NODE_ENV'] === 'test' ? 1 : Math.max(1, Math.min(8, configuredConcAll));
    this.resetQueue(concurrency);
    try {
      do {
        const query: DriveListQuery = pageToken ? { folderId: fid, pageToken } : { folderId: fid };
        const { files, nextPageToken }: DriveListResult = await this.google.listDriveFiles(query);
        for (const f of files) {
          total++;
          this.enqueue(f);
        }
        pageToken = nextPageToken;
      } while (pageToken);
    } finally {
      await this.waitQueueDrain();
    }

    const durationMs = Date.now() - start;
    this.metrics?.observeHistogram?.('drive_index_duration_seconds', durationMs / 1000, { mode: 'full' });
    this.metrics?.incCounter?.('drive_index_files_indexed_total', { mode: 'full', total });
    logger.info('✅ Полная индексация завершена', { total, durationMs });
  }

  /** Инкрементальная индексация: только новые/изменённые */
  public async reindexIncremental(folderId?: string): Promise<void> {
    if (!this.ensureReady()) return;
    const fid = folderId || this.config.google.driveFolderId || this.config.drive.folderId;
    if (!fid) return;

    logger.info('🔄 Запуск инкрементальной индексации', { folderId: fid });
    const start = Date.now();
    this.metrics?.incCounter?.('drive_index_runs_total', { mode: 'incremental' });
    let pageToken: string | undefined = undefined;
    let updated = 0;

    const concurrency = Math.max(1, Math.min(8, Number(this.config.drive?.maxConcurrency ?? 4)));
    this.resetQueue(concurrency, () => { updated++; });
    try {
      do {
        const query: DriveListQuery = pageToken ? { folderId: fid, pageToken } : { folderId: fid };
        const { files, nextPageToken }: DriveListResult = await this.google.listDriveFiles(query);
        for (const f of files) {
          if (await this.needReindex(f)) this.enqueue(f);
        }
        pageToken = nextPageToken;
      } while (pageToken);
    } finally {
      await this.waitQueueDrain();
    }

    const durationMs = Date.now() - start;
    this.metrics?.observeHistogram?.('drive_index_duration_seconds', durationMs / 1000, { mode: 'incremental' });
    this.metrics?.incCounter?.('drive_index_files_indexed_total', { mode: 'incremental', total: updated });
    logger.info('✅ Инкрементальная индексация завершена', { updated, durationMs });
  }

  /** Простая выдача по содержимому */
  public async search(query: string, limit = 10): Promise<DriveSearchResult[]> {
    // If we have SQLite FTS available, use it for better search results
    if (this.searchIndex) {
      try {
        const result = await this.searchIndex.search({
          text: query,
          limit
        });
        
        return result.hits.map(hit => ({
          file: {
            id: hit.fileId,
            name: hit.name,
            mimeType: hit.mimeType || '',
            textLength: hit.textLen || 0,
            updatedAt: Date.now(),
            snippet: hit.snippet || '',
            // Fix: Handle undefined modifiedTime properly
            ...(hit.modifiedTime ? { modifiedTime: new Date(hit.modifiedTime).toISOString() } : {}),
            ...(hit.ownerEmail ? { owners: [hit.ownerEmail] } : {}),
          },
          score: hit.score || 0
        }));
      } catch (error) {
        logger.warn('⚠️ Пошук через FTS не вдався, використовуємо резервний метод', { 
          error: error instanceof Error ? error.message : String(error) 
        });
        // Fall back to the original search method
      }
    }
    
    // Original search method as fallback
    if (!this.ensureReady()) return [];
    const ids = await this.cache.get<string[]>(INDEX_KEYS);
    if (!ids || ids.length === 0) return [];

    const q = query.trim().toLowerCase();
    const results: DriveSearchResult[] = [];
    for (const id of ids) {
      const entry = await this.cache.get<DriveIndexEntry>(INDEX_PREFIX + id);
      if (!entry || !entry.text) continue;
      const textLower = entry.text.toLowerCase();
      const idx = textLower.indexOf(q);
      if (idx >= 0) {
        const snippet = this.makeSnippet(entry.text, idx, q.length);
        const fileBase = {
          id: entry.id,
          name: entry.name,
          mimeType: entry.mimeType,
          textLength: entry.textLength,
          updatedAt: entry.updatedAt,
          snippet,
        } as const;

        const fileObj: Omit<DriveIndexEntry, 'text'> & { snippet: string } = {
          ...(fileBase as any),
          ...(entry.modifiedTime ? { modifiedTime: entry.modifiedTime } : {}),
          ...(entry.owners ? { owners: entry.owners } : {}),
          ...(typeof entry.size === 'number' ? { size: entry.size } : {}),
        };

        results.push({ file: fileObj, score: 1 / (1 + idx) });
      }
      if (results.length >= limit) break;
    }

    // простая сортировка по позиции первого вхождения
    results.sort((a, b) => b.score - a.score);
    return results.slice(0, limit);
  }

  /** Получить полную запись индекса по fileId (если доступна в кэше) */
  public async getEntry(fileId: string): Promise<DriveIndexEntry | null> {
    if (!this.ensureReady()) return null;
    const entry = await this.cache.get<DriveIndexEntry>(INDEX_PREFIX + fileId);
    return entry ?? null;
  }

  /** Получить исходный текст (обрезанный до MAX_TEXT_STORED) для предпросмотра */
  public async getText(fileId: string): Promise<string> {
    const entry = await this.getEntry(fileId);
    return entry?.text ?? '';
  }

  /** Индексация одного файла по метаданным (без повторного запроса метаданных) */
  public async indexOneFileByMeta(file: DriveFile): Promise<void> {
    if (!this.ensureReady()) return;
    if (!this.isIndexableMime(file.mimeType)) {
      this.metrics?.incCounter?.('drive_index_skipped_total', { reason: 'non_indexable_mime', mime: file.mimeType });
      return;
    }

    // основний воркер, загорнутий ретраями (бекоф для 429/5xx/таймаутів)
    const worker = async () => {
      // Спеціальна гілка для Google Sheets: індексувати кожен лист як окремий "віртуальний документ"
      if (
        file.mimeType === 'application/vnd.google-apps.spreadsheet' &&
        this.searchIndex &&
        file.id &&
        typeof (this.google as any).listSheets === 'function' &&
        typeof (this.google as any).getSheetData === 'function'
      ) {
        try {
          const tabs = await (this.google as any).listSheets(file.id);
          const pieces: string[] = [];
          const breadcrumbs = await this.safeGetBreadcrumbs(file.id);
          for (const tab of tabs) {
            const range = `'${tab}'!A1:Z1000`;
            const data = await (this.google as any).getSheetData(file.id, range, { useCache: true, cacheTTL: this.config.drive?.ttlTextSec ?? 300 });
            const values: string[][] = Array.isArray(data?.values)
              ? (data.values as unknown[][]).map((row: unknown[]) =>
                  Array.isArray(row) ? row.map((v: unknown) => (v == null ? '' : String(v))) : []
                )
              : [];
            const textRows = values.map((row: string[]) => row.join(' | '));
            const text = textRows.join('\n');
            pieces.push(`# ${tab}\n${text}`);
            // Upsert в FTS як окремий "документ"
            const lang = detectLanguage(text);
            const owner = Array.isArray(file.owners) && file.owners.length ? file.owners[0] : undefined;
            const path = `/${[...breadcrumbs, file.name, tab].join('/')}`;
            const labels = Array.isArray((file as any).labels) ? ((file as any).labels as string[]) : undefined;
            await this.searchIndex.upsert({
              fileId: `${file.id}:${encodeURIComponent(tab)}`,
              name: `${file.name} — ${tab}`,
              mimeType: file.mimeType,
              ...(owner ? { ownerEmail: owner } : {}),
              ...(typeof file.size === 'number' ? { sizeBytes: file.size } : {}),
              ...(file.modifiedTime ? { modifiedTime: Date.parse(file.modifiedTime) } : {}),
              text,
              language: lang,
              path,
              ...(labels ? { labels } : {}),
              meta: {
                sheet: file.id,
                tab,
                range,
                breadcrumbs,
                lang,
              },
            });
            this.metrics?.incCounter?.('drive_index_file_indexed', { mime: file.mimeType });
          }
          // Зберігаємо агрегований текст у кеш для попереднього перегляду
          await this.saveEntry(file, pieces.join('\n\n'));
          this.indexedCount += Math.max(1, tabs.length);
          return;
        } catch (e) {
          logger.warn('⚠️ Індексація Sheets (віртуальні документи) не вдалася, буде використано fallback', { id: file.id, error: e instanceof Error ? e.message : String(e) });
          // Продовжимо універсальною гілкою нижче
        }
      }

      // Використовуємо уніфіковану екстракцію через парсери з нормалізацією
      let text = '';
      const useLegacy = process.env['NODE_ENV'] === 'test' || typeof (this.google as any).extractTextForChat !== 'function';
      if (useLegacy) {
        text = await (this.google as any).extractTextFromFile({ id: file.id, mimeType: file.mimeType, name: file.name, modifiedTime: file.modifiedTime ?? null });
      } else {
        const res = await (this.google as any).extractTextForChat(file.id);
        text = res?.text || '';
      }
      await this.saveEntry(file, text);
      // Persist to SQLite FTS index (best-effort)
      try {
        if (this.searchIndex && file.id) {
          const modifiedMs = file.modifiedTime ? Date.parse(file.modifiedTime) : undefined;
          const lang = detectLanguage(text);
          const breadcrumbs = await this.safeGetBreadcrumbs(file.id);
          const path = `/${[...breadcrumbs, file.name].join('/')}`;
          const labels = Array.isArray((file as any).labels) ? ((file as any).labels as string[]) : undefined;
          const payload: {
            fileId: string;
            name: string;
            mimeType: string;
            sizeBytes?: number;
            modifiedTime?: number;
            createdTime?: number;
            text: string;
            tags?: string[];
            meta?: unknown;
            ownerEmail?: string;
            language?: string;
            labels?: string[];
            path?: string;
          } = {
            fileId: file.id,
            name: file.name,
            mimeType: file.mimeType,
            text,
            language: lang,
            path,
            ...(labels ? { labels } : {}),
            meta: {
              webViewLink: file.webViewLink,
              parents: file.parents,
              isShortcut: file.isShortcut,
              shortcutTargetId: file.shortcutDetails?.targetId,
              lang,
              breadcrumbs,
            },
          };
          if (Array.isArray(file.owners) && file.owners.length) {
            const oe = (file.owners[0] as any);
            const email = typeof oe === 'string' ? oe : (typeof oe?.emailAddress === 'string' ? oe.emailAddress : undefined);
            if (typeof email === 'string' && email.length > 0) {
              payload.ownerEmail = email;
            }
          }
          if (typeof file.size === 'number') payload.sizeBytes = file.size;
          if (Number.isFinite(modifiedMs as number)) payload.modifiedTime = modifiedMs as number;
          await this.searchIndex.upsert(payload);
        }
      } catch (e) {
        logger.warn('⚠️ Помилка індексації у FTS', { id: file.id, error: e instanceof Error ? e.message : String(e) });
      }
      // Always count per-file indexing regardless of FTS availability
      this.metrics?.incCounter?.('drive_index_file_indexed', { mime: file.mimeType });
    };

    const retry = await RetryManager.execute(worker, {
      maxAttempts: 4,
      delay: 800,
      backoff: 'exponential',
      factor: 2,
      maxDelay: 10_000,
      timeout: 60_000,
      shouldRetry: (err: Error) => {
        const any = err as any;
        const status = typeof any?.status === 'number' ? any.status : undefined;
        const code = typeof any?.code === 'string' ? any.code : undefined;
        const msg = String(any?.message ?? '').toLowerCase();
        return (
          status === 429 || (typeof status === 'number' && status >= 500) ||
          code === 'ECONNRESET' || code === 'ETIMEDOUT' || msg.includes('timeout')
        );
      },
      onRetry: (attempt, error) => {
        this.metrics?.incCounter?.('drive_index_retries_total', { attempt });
        logger.warn('🔁 Повторна спроба індексації файла', { id: file.id, attempt, error: error.message });
      },
    });
    if (!retry.success) throw retry.error;
    
  }
  
  // === helpers ===
  private ensureReady(): boolean {
    if (!this.google) {
      logger.warn('GoogleService недоступен для индексатора');
      return false;
    }
    if (!this.cache) {
      logger.warn('CacheService недоступен для индексатора');
      return false;
    }
    if (!this.config.drive?.enableTextIndex) {
      return false;
    }
    return true;
  }

  private isIndexableMime(mime: string): boolean {
    return (
      mime === 'application/vnd.google-apps.document' ||
      mime === 'application/pdf' ||
      mime === 'application/msword' ||
      mime === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
      mime === 'text/plain' ||
      mime === 'application/vnd.google-apps.spreadsheet'
    );
  }

  private async needReindex(f: DriveFile): Promise<boolean> {
    const key = INDEX_PREFIX + f.id;
    const existing = await this.cache.get<DriveIndexEntry>(key);
    if (!existing) return true;
    // сравниваем modifiedTime
    if (f.modifiedTime && existing.modifiedTime && f.modifiedTime !== existing.modifiedTime) return true;
    // TTL перевірка
    const ttlSec = this.config.drive?.ttlTextSec || 21600; // 6h
    const expired = (Date.now() - (existing.updatedAt || 0)) / 1000 > ttlSec;
    return expired;
  }

  private async saveEntry(file: DriveFile, textRaw: string): Promise<void> {
    const text = (textRaw || '').slice(0, MAX_TEXT_STORED);
    const entryBase = {
      id: file.id,
      name: file.name,
      mimeType: file.mimeType,
      text,
      textLength: (textRaw || '').length,
      updatedAt: Date.now(),
    } as const;

    const entry: DriveIndexEntry = {
      ...entryBase,
      ...(file.modifiedTime ? { modifiedTime: file.modifiedTime } : {}),
      ...(file.owners ? { owners: file.owners } : {}),
      ...(typeof file.size === 'number' ? { size: file.size } : {}),
    } as DriveIndexEntry;

    const key = INDEX_PREFIX + file.id;
    await this.cache.set(key, entry, this.config.drive.ttlTextSec || 21600);

    // поддержка списка ключей
    let ids = (await this.cache.get<string[]>(INDEX_KEYS)) || [];
    if (!ids.includes(file.id)) ids = [...ids, file.id];
    await this.cache.set(INDEX_KEYS, ids, this.config.drive.ttlTextSec || 21600);

    logger.debug('🧾 Индекс обновлен', { id: file.id, name: file.name, len: entry.textLength });
  }

  private makeSnippet(text: string, idx: number, qlen: number): string {
    const start = Math.max(0, idx - 80);
    const end = Math.min(text.length, idx + qlen + 80);
    const prefix = start > 0 ? '…' : '';
    const suffix = end < text.length ? '…' : '';
    return prefix + text.slice(start, end).replace(/\s+/g, ' ') + suffix;
  }

  /**
   * Обчислити breadcrumbs для файлу. Поки що повертаємо порожній список (MVP),
   * щоб не ламати збірку; логіка обчислення людяних шляхів буде додана на етапі 1.3.
   */
  private async safeGetBreadcrumbs(_fileId: string): Promise<string[]> {
    try {
      // TODO: implement folder path resolution via Drive parents chain
      return [];
    } catch {
      return [];
    }
  }

  // === BaseService required hooks ===
  protected async onShutdown(): Promise<void> {
    // no-op
  }

  protected async onHealthCheck(): Promise<HealthStatus> {
    return {
      healthy: true,
      service: this.name,
      timestamp: new Date(),
    } as unknown as HealthStatus; // адаптация к текущему типу HealthStatus в проекте
  }

  protected onGetStats(): Partial<ServiceStats> {
    return {
      indexedCount: this.indexedCount,
      lastRunAt: this.lastRunAt ?? undefined,
    } as Partial<ServiceStats>;
  }
}

export default DriveIndexerService;
