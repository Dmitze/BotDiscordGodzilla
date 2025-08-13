/**
 * DriveIndexerService
 * Индексация содержимого файлов Google Drive для простого полнотекстового поиска
 */

import type { BotConfig, HealthStatus, ServiceStats } from '@/types';
import type { DriveFile, DriveListResult, DriveListQuery } from '@/types/drive';
import logger from '@/utils/logger';
import { BaseService as BaseServiceClass } from '@/core/BaseService';
import type { GoogleService } from './GoogleService';
import type { SearchIndex } from '@/search/SearchIndex';
import type { CacheService } from './CacheService';
import SchedulerService from './SchedulerService';
import { chunkTextForDiscord } from '@/utils/chunk';

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

export class DriveIndexerService extends BaseServiceClass {
  private bot: BotLike;
  private google!: GoogleService;
  private cache!: CacheService;
  private searchIndex: SearchIndex | undefined;
  private metrics?: { incCounter?: (...args: any[]) => void; observeHistogram?: (...args: any[]) => void };
  private indexedCount = 0;
  private lastRunAt: number | null = null;
  
  private isCronDisabled(): boolean {
    return (
      process.env['NODE_ENV'] === 'test' ||
      String(process.env['DISABLE_CRON']).toLowerCase() === 'true'
    );
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
      if (m) this.metrics = m as any;
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

    // формальный await, чтобы удовлетворить линтер (async без await)
    await Promise.resolve();
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

    do {
      const query: DriveListQuery = pageToken ? { folderId: fid, pageToken } : { folderId: fid };
      const { files, nextPageToken }: DriveListResult = await this.google.listDriveFiles(query);
      for (const f of files as DriveFile[]) {
        await this.indexOneFileByMeta(f).catch(err => {
          logger.warn('⚠️ Индексация файла пропущена', { id: f.id, error: err instanceof Error ? err.message : String(err) });
        });
        total++;
      }
      pageToken = nextPageToken;
    } while (pageToken);

    const durationMs = Date.now() - start;
    this.metrics?.observeHistogram?.('drive_index_duration_seconds', durationMs / 1000, { mode: 'full' });
    this.metrics?.incCounter?.('drive_index_files_indexed_total', { mode: 'full', total });
    logger.info('✅ Полная индексация завершена', { total });
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

    do {
      const query: DriveListQuery = pageToken ? { folderId: fid, pageToken } : { folderId: fid };
      const { files, nextPageToken }: DriveListResult = await this.google.listDriveFiles(query);
      for (const f of files as DriveFile[]) {
        const need = await this.needReindex(f);
        if (!need) continue;
        await this.indexOneFileByMeta(f).catch(err => {
          logger.warn('⚠️ Индексация файла пропущена', { id: f.id, error: err instanceof Error ? err.message : String(err) });
        });
        updated++;
      }
      pageToken = nextPageToken;
    } while (pageToken);

    const durationMs = Date.now() - start;
    this.metrics?.observeHistogram?.('drive_index_duration_seconds', durationMs / 1000, { mode: 'incremental' });
    this.metrics?.incCounter?.('drive_index_files_indexed_total', { mode: 'incremental', total: updated });
    logger.info('✅ Инкрементальная индексация завершена', { updated });
  }

  /** Простая выдача по содержимому */
  public async search(query: string, limit = 10): Promise<DriveSearchResult[]> {
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
        } as any;

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

  /** Получить чанки текста, безопасные для Discord */
  public async getTextChunks(fileId: string, max = 1900): Promise<string[]> {
    const text = await this.getText(fileId);
    if (!text) return [];
    return chunkTextForDiscord(text, max);
  }

  /** Индексация одного файла по метаданным (без повторного запроса метаданных) */
  public async indexOneFileByMeta(file: DriveFile): Promise<void> {
    if (!this.ensureReady()) return;
    if (!this.isIndexableMime(file.mimeType)) {
      this.metrics?.incCounter?.('drive_index_skipped_total', { reason: 'non_indexable_mime', mime: file.mimeType });
      return;
    }

    const text = await this.google.extractTextFromFile({ id: file.id, mimeType: file.mimeType, name: file.name, modifiedTime: file.modifiedTime ?? null });
    await this.saveEntry(file, text);
    // Persist to SQLite FTS index (best-effort)
    try {
      if (this.searchIndex && file.id) {
        const modifiedMs = file.modifiedTime ? Date.parse(file.modifiedTime) : undefined;
        const payload: {
          fileId: string;
          name: string;
          mimeType?: string;
          ownerEmail?: string;
          sizeBytes?: number;
          modifiedTime?: number;
          createdTime?: number;
          text: string;
          tags?: string[];
          meta?: unknown;
        } = {
          fileId: file.id,
          name: file.name,
          mimeType: file.mimeType,
          text,
          meta: {
            webViewLink: file.webViewLink,
            parents: file.parents,
            isShortcut: file.isShortcut,
            shortcutTargetId: file.shortcutDetails?.targetId,
          },
        };
        const owner = Array.isArray(file.owners) && file.owners.length ? file.owners[0] : undefined;
        if (owner) payload.ownerEmail = owner;
        if (typeof file.size === 'number') payload.sizeBytes = file.size;
        if (Number.isFinite(modifiedMs as number)) payload.modifiedTime = modifiedMs as number;
        await this.searchIndex.upsert(payload);
      }
    } catch (e) {
      logger.warn('⚠️ Не вдалося оновити SqliteSearchIndex', { id: file.id, error: e instanceof Error ? e.message : String(e) });
    }
    this.indexedCount++;
    this.metrics?.incCounter?.('drive_index_file_indexed', { mime: file.mimeType });
  }

  /** Индексация одного файла по id (сама достанет метаданные) */
  public async indexOne(fileId: string): Promise<void> {
    if (!this.ensureReady()) return;
    const meta = await this.google.getDriveFile(fileId);
    await this.indexOneFileByMeta(meta);
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
      mime === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    );
  }

  private async needReindex(f: DriveFile): Promise<boolean> {
    const key = INDEX_PREFIX + f.id;
    const existing = await this.cache.get<DriveIndexEntry>(key);
    if (!existing) return true;
    // сравниваем modifiedTime
    if (f.modifiedTime && existing.modifiedTime && f.modifiedTime === existing.modifiedTime) return false;
    return true;
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
