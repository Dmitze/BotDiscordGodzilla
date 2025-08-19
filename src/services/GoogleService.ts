/**
 * Google Service з Connection Pool та оптимізацією
 * Покращена продуктивність та стабільність
 */

import { google, type sheets_v4, type drive_v3, type docs_v1 } from 'googleapis';
import type { DocBlock } from '@/types/docs';
import { ImageAnnotatorClient } from '@google-cloud/vision';
import type { MetricsService } from './MetricsService';
import { createHash } from 'crypto';
// Parsers
import { createDefaultParserRouter } from '@/parsers';
import type { BotConfig, HealthStatus, ServiceStats, SheetData, BatchSheetData } from '@/types';
import type { DriveListQuery, DriveListResult, DriveFile } from '@/types/drive';
import { BaseService as BaseServiceClass } from '@/core/BaseService';
import { CacheService } from './CacheService';
import logger from '@/utils/logger';
import { DocsService } from './google/DocsService';
import { SheetsService } from './google/SheetsService';
import { sanitizeTextForChat, normalizeText } from '@/utils/fileProcessor';
import { validateInput } from '@/utils/security';

interface GoogleServiceStats extends ServiceStats {
  requests: number;
  errors: number;
  averageResponseTime: number;
  connectionPoolUsage: number;
  cacheHits: number;
  cacheMisses: number;
}

interface ConnectionInfo {
  inUse: boolean;
  lastUsed: number;
  requestCount: number;
}

interface GoogleServiceOptions {
  useCache?: boolean;
  cacheTTL?: number;
  forceRefresh?: boolean;
  batchSize?: number;
  retryFailed?: boolean;
  cacheResults?: boolean;
  maxRetries?: number;
  valueInputOption?: 'RAW' | 'USER_ENTERED';
  clearCache?: boolean;
}

export class GoogleService extends BaseServiceClass {
  private auth: InstanceType<typeof google.auth.JWT> | null = null;
  private sheets: sheets_v4.Sheets | null = null;
  private drive: drive_v3.Drive | null = null;
  private docs: docs_v1.Docs | null = null;
  private connectionPool = new Map<string, ConnectionInfo>();
  private readonly retryAttempts = 3;
  private readonly retryDelay = 1000;
  private stats: GoogleServiceStats;
  private cacheService: CacheService;
  private metrics?: MetricsService;
  private docsService?: DocsService;
  private sheetsService?: SheetsService;
  // Token-bucket per apiType (drive|sheets|docs)
  private rlTokens = new Map<string, number>();
  private rlLastRefill = new Map<string, number>();

  constructor(config: BotConfig) {
    super('GoogleService', config);
    this.cacheService = new CacheService(config);
    this.stats = {
      service: 'GoogleService',
      uptime: 0,
      requests: 0,
      errors: 0,
      averageResponseTime: 0,
      connectionPoolUsage: 0,
      cacheHits: 0,
      cacheMisses: 0,
    };
    // Инициализация токенов бурстом для известных apiType
    const { burst } = this.getRateConfig();
    ['drive', 'sheets', 'docs'].forEach(t => {
      this.rlTokens.set(t, burst);
      this.rlLastRefill.set(t, Date.now());
    });
  }

  /**
   * Извлечение текста для чата с валидацией/санитизацией, хэш-контролем и фоллбек-парсерами
   * Возвращает текст, источник (parser/export/ocr), контрольную сумму и modifiedTime
   */
  public async extractTextForChat(fileId: string): Promise<{
    text: string;
    checksum: string;
    modifiedTime?: string;
    source: 'export' | 'parser' | 'ocr' | 'raw';
    warnings: string[];
  }> {
    const overallStart = Date.now();
    const meta = await this.getDriveFileMetadata(fileId);
    const mime = String(meta.mimeType || '');
    const modified = meta.modifiedTime ?? '';

    // Используем модификацию и потенциальный контент-хэш для кэш-ключа
    const baseCacheKey = `extract:text:${fileId}:${modified}:${mime}`;
    try {
      const cached = await this.cacheService.get<{ text: string; checksum: string; modifiedTime?: string; source: 'export' | 'parser' | 'ocr' | 'raw'; warnings: string[] }>(baseCacheKey);
      if (cached?.text) {
        this.stats.cacheHits++;
        // Cache hit: lightweight metrics
        try { this.metrics?.recordFileOperation({ operation: 'extract_text_cache', status: 'success', mime, fileId }); } catch {}
        return cached;
      }
    } catch {
      this.stats.cacheMisses++;
    }

    const warnings: string[] = [];
    let buffer: Buffer | null = null;
    let text = '';
    let source: 'export' | 'parser' | 'ocr' | 'raw' = 'raw';

    // Маршрутизация через модульные парсеры
    try {
      const router = createDefaultParserRouter({
        timeoutMs: this.config.drive?.parseTimeoutMs ?? 10000,
        retryAttempts: this.config.drive?.parseRetryAttempts ?? 1,
        retryDelayMs: this.config.drive?.parseRetryDelayMs ?? 200,
      });
      const parsed = await router.parse(
        {
          id: String(meta.id || fileId),
          name: String((meta as any).name || ''),
          mimeType: mime,
          modifiedTime: meta.modifiedTime,
          owners: (meta as any).owners || [],
          webViewLink: (meta as any).webViewLink,
          iconLink: (meta as any).iconLink,
          size: (meta as any).size ? Number((meta as any).size) : undefined,
          isShortcut: (meta as any).shortcutDetails ? true : false,
          shortcutTargetId: (meta as any).shortcutDetails?.targetId,
        } as any,
        {
          exportFile: (id, m) => this.exportFile(id, m),
          downloadFile: (id) => this.downloadFile(id),
          extractTextFromImage: async (_file) => {
            // not used in current parsers; OCR uses buffer path
            const buf = await this.downloadFile(fileId);
            return this.extractTextFromBuffer(buf);
          },
          extractTextFromBuffer: (buf) => this.extractTextFromBuffer(buf),
        }
      );
      text = parsed.text;
      source = (parsed.source as any) ?? 'raw';
      buffer = parsed.buffer ?? null;
    } catch (error) {
      logger.error('❌ Помилка витягання тексту', {
        type: 'processing_error',
        event: 'extract_text_failed',
        component: 'GoogleService',
        fileId,
        mime,
        error: error instanceof Error ? error.message : String(error),
      });
      text = '';
      try { this.metrics?.recordFileOperation({ operation: 'extract_text', status: 'error', mime, fileId }); } catch {}
    }

    // Нормализация и очистка
    text = normalizeText(text);
    const val = validateInput(text, { inputType: 'message' });
    if (!val.isValid) warnings.push(`sanitization: ${val.errors.join(', ')}`);
    const safe = sanitizeTextForChat(text);

    // Контрольная сумма
    const checksum = buffer ? createHash('sha256').update(buffer).digest('hex') : createHash('sha256').update(safe).digest('hex');

    const result = { text: safe, checksum, modifiedTime: modified, source, warnings } as const;
    try {
      const ttl = this.config.drive?.ttlTextSec ?? 300;
      await this.cacheService.set(baseCacheKey, result, ttl);
    } catch { /* istanbul ignore next */ }
    // Final overall metrics for extractTextForChat
    try {
      this.metrics?.recordFileOperation({ operation: 'extract_text', status: 'success', mime, fileId });
      this.metrics?.observeFileOperationLatency('extract_text', mime, Date.now() - overallStart);
      this.metrics?.observeTextSizeBytes(source, Buffer.byteLength(safe, 'utf8'));
      if (mime) this.metrics?.incrementMimeType(mime);
    } catch {}
    return result;
  }

  /**
   * Структурированное содержимое Google Docs
   */
  public async getDocumentBlocks(documentId: string): Promise<DocBlock[]> {
    try {
      const result = await this.executeWithRetry(async () => {
        if (!this.docs) throw new Error('Docs API не ініціалізовано');
        const response = await this.docs.documents.get({ documentId });
        return this.getDocsService().extractBlocksFromDoc(response.data);
      }, 'docs', undefined, 'docs.documents.get');
      return result;
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка отримання структурованого контенту документа', {
          type: 'api_error',
          event: 'docs_blocks_failed',
          component: 'GoogleService',
          service: 'docs',
          documentId,
          errorName: error.name,
          errorMessage: error.message,
          stack: error.stack,
        });
      } else {
        logger.error('❌ Помилка отримання структурованого контенту документа', {
          type: 'api_error',
          event: 'docs_blocks_failed',
          component: 'GoogleService',
          service: 'docs',
          documentId,
          errorMessage: String(error),
        });
      }
      throw error;
    }
  }

  /** Подключение MetricsService (вызывается из ServiceManager) */
  public setMetricsService(ms: MetricsService): void {
    this.metrics = ms;
    // Лениво инициализируем под-сервисы, чтобы они могли писать метрики
    this.docsService = new DocsService(this.metrics);
    this.sheetsService = new SheetsService(this.metrics);
  }

  /** Получить сервис Google Docs parser (без сетевых вызовов) */
  public getDocsService(): DocsService {
    if (!this.docsService) this.docsService = new DocsService(this.metrics);
    return this.docsService;
  }

  /** Получить сервис Google Sheets helper (без сетевых вызовов) */
  public getSheetsService(): SheetsService {
    if (!this.sheetsService) this.sheetsService = new SheetsService(this.metrics);
    return this.sheetsService;
  }

  /** Получение параметров rate-limit из конфига с дефолтами */
  private getRateConfig(): { qps: number; burst: number } {
    const qps = this.config.drive?.rateQps ?? 5;
    const burst = this.config.drive?.rateBurst ?? 10;
    return { qps: Math.max(1, qps), burst: Math.max(1, burst) };
  }

  /** Ожидание до доступности токена по token-bucket; возвращает задержку в мс */
  private async throttle(apiType: string): Promise<number> {
    const { qps, burst } = this.getRateConfig();
    const now = Date.now();
    const last = this.rlLastRefill.get(apiType) ?? now;
    const prevTokens = this.rlTokens.get(apiType) ?? burst;
    const elapsedSec = (now - last) / 1000;
    // Пополнение токенов
    let tokens = Math.min(burst, prevTokens + elapsedSec * qps);
    this.rlLastRefill.set(apiType, now);

    if (tokens >= 1) {
      tokens -= 1;
      this.rlTokens.set(apiType, tokens);
      return 0;
    }

    const need = 1 - tokens;
    const waitMs = Math.ceil((need / qps) * 1000);
    await new Promise(resolve => setTimeout(resolve, waitMs));
    // Списываем токен после ожидания
    const afterTokens = Math.max(0, tokens + (waitMs / 1000) * qps - 1);
    this.rlTokens.set(apiType, afterTokens);
    this.rlLastRefill.set(apiType, Date.now());
    // Метрика "throttled"
    try {
      this.metrics?.updateGoogleApiMetrics(apiType, 'throttle', 'throttled', waitMs);
    } catch (/* istanbul ignore next */ _e) {
      // noop: метрики не критичны
    }
    return waitMs;
  }

  /**
   * Нормализация Google Drive файла к внутреннему типу DriveFile
   */
  private toDriveFile(file: drive_v3.Schema$File): DriveFile {
    const owners = this.getOwnerNamesOrEmails(file);
    const base: DriveFile = {
      id: String(file.id || ''),
      name: String(file.name || ''),
      mimeType: String(file.mimeType || ''),
    };
    const opt = this.buildDriveFileOptional(file, owners);
    return { ...base, ...opt } as DriveFile;
  }

  /**
   * Извлекает список владельцев (email/displayName) с фильтрацией пустых значений
   */
  private getOwnerNamesOrEmails(file: drive_v3.Schema$File): string[] {
    const ownersRaw = Array.isArray(file.owners) ? file.owners : [];
    return ownersRaw
      .map(o => (o?.emailAddress ? o.emailAddress : o?.displayName))
      .filter((v): v is string => typeof v === 'string' && v.length > 0);
  }

  /**
   * Безопасно мапит shortcutDetails в наш внутренний вид
   */
  private mapShortcutDetails(file: drive_v3.Schema$File): { targetId: string; targetMimeType?: string } | undefined {
    const targetId = file.shortcutDetails?.targetId;
    if (typeof targetId !== 'string' || targetId.length === 0) return undefined;
    const out: { targetId: string; targetMimeType?: string } = { targetId };
    const mime = file.shortcutDetails?.targetMimeType;
    if (typeof mime === 'string') out.targetMimeType = mime;
    return out;
  }

  /**
   * Обёртка над getDriveFileMetadata с нормализацией в DriveFile
   */
  public async getDriveFile(fileId: string): Promise<DriveFile> {
    const raw = await this.getDriveFileMetadata(fileId);
    return this.toDriveFile(raw);
  }

  /**
   * Список файлов по запросу DriveListQuery с пагинацией и кэшем
   */
  public async listDriveFiles(query: DriveListQuery): Promise<DriveListResult> {
    const {
      folderId,
      query: nameContains = '',
      mimeIncludes = [],
      ownerAllowlist = [],
      pageSize,
      pageToken,
      dateFrom,
      dateTo,
      sizeMin,
      sizeMax,
      sortBy,
      sortDir,
      highlightChanges,
      sessionKey,
    } = query;

    const size = this.getDriveListPageSize(pageSize);
    const cacheKey = this.buildDriveListCacheKey(query, size);

    // Кэшованный ответ
    try {
      const cached = await this.cacheService.get<DriveListResult>(cacheKey);
      if (cached) {
        this.stats.cacheHits++;
        logger.debug('✅ Кэш листинга Drive', {
          type: 'system',
          event: 'cache_hit',
          component: 'GoogleService',
          folderId,
        });
        return cached;
      }
    } catch (/* istanbul ignore next */ _e) {
      // noop: промах кеша не критичен для выполнения
      this.stats.cacheMisses++;
    }

    const allowedMimeCfg = this.config.drive.allowedMime || ['*'];
    const needMimeFilter = !(allowedMimeCfg.length === 1 && allowedMimeCfg[0] === '*');

    // Построение Drive query (q)
    const q = this.buildDriveQuery(folderId, nameContains, mimeIncludes, allowedMimeCfg, needMimeFilter, dateFrom, dateTo);

    const params = this.buildFilesListParams(q, size, pageToken ?? undefined);
    const start = Date.now();
    const res = await this.executeDriveFilesList(params);

    const duration = Date.now() - start;
    const filesRaw: drive_v3.Schema$File[] = Array.isArray(res.data.files) ? res.data.files : [];
    let result = this.createDriveListResult(filesRaw, ownerAllowlist, res.data.nextPageToken ?? undefined);

    // Пост-фильтры по размеру
    if (typeof sizeMin === 'number' || typeof sizeMax === 'number') {
      const min = typeof sizeMin === 'number' ? sizeMin : -Infinity;
      const max = typeof sizeMax === 'number' ? sizeMax : Infinity;
      result.files = result.files.filter(f => {
        const sz = typeof f.size === 'number' ? f.size : undefined;
        if (sz == null) return false; // если требуется фильтр по размеру, отсутствующие размеры исключаем
        return sz >= min && sz <= max;
      });
    }

    // Сортировка
    if (sortBy) {
      const dir = sortDir === 'desc' ? -1 : 1;
      const cmp = (a: DriveFile, b: DriveFile): number => {
        if (sortBy === 'name') return (a.name || '').localeCompare(b.name || '') * dir;
        if (sortBy === 'size') return ((a.size ?? -1) - (b.size ?? -1)) * dir;
        // modifiedTime
        const da = a.modifiedTime ? Date.parse(a.modifiedTime) : 0;
        const db = b.modifiedTime ? Date.parse(b.modifiedTime) : 0;
        return (da - db) * dir;
      };
      result.files = [...result.files].sort(cmp);
    }

    // Подсветка изменений между запросами одной сессии
    if (highlightChanges && sessionKey) {
      const snapKey = `drive:list:snap:${sessionKey}:${folderId}:${this.hashFiltersForSnapshot({ nameContains, mimeIncludes, ownerAllowlist, dateFrom, dateTo, sizeMin, sizeMax, sortBy, sortDir })}`;
      try {
        const prev = await this.cacheService.get<DriveFile[]>(snapKey);
        if (Array.isArray(prev)) {
          const changes = this.computeDriveListChanges(prev, result.files);
          if (changes.addedIds.length || changes.removedIds.length || changes.modified.length) {
            result = { ...result, changes };
          }
        }
      } catch { /* ignore cache read */ }
      try {
        // сохраняем текущий снимок на короткое время
        await this.cacheService.set(snapKey, result.files, this.config.drive.ttlListSec ?? 300);
      } catch { /* ignore cache write */ }
    }

    await this.cacheDriveListResult(cacheKey, result);

    this.logDriveListSuccess({
      folderId,
      result,
      duration,
      pageTokenPresent: Boolean(pageToken),
    });

    // Рекурсивная ветка (упрощённо): не разворачиваем здесь глубину, оставим на будущий индексатор
    // Можно реализовать отдельным сервисом, чтобы не блокировать текущий шаг.

    return result;
  }

  // -------- Helpers for toDriveFile --------
  private buildDriveFileOptional(
    file: drive_v3.Schema$File,
    owners: string[],
  ): Partial<DriveFile> {
    const webViewLink = (file as { webViewLink?: string | null }).webViewLink ?? undefined;
    const iconLink = (file as { iconLink?: string | null }).iconLink ?? undefined;
    const hideWeb = Boolean(this.config.drive?.hideWebLink);

    const out: Partial<DriveFile> = {};
    if (file.size) out.size = Number(file.size);
    if (file.modifiedTime) out.modifiedTime = file.modifiedTime;
    if (owners.length > 0) out.owners = owners;
    if (file.parents) out.parents = file.parents;
    if (webViewLink && !hideWeb) out.webViewLink = webViewLink;
    if (iconLink) out.iconLink = iconLink;
    if (file.mimeType === 'application/vnd.google-apps.shortcut') out.isShortcut = true;
    const sd = this.mapShortcutDetails(file);
    if (sd) out.shortcutDetails = sd;
    return out;
  }

  // -------- Helpers for listDriveFiles --------
  private getDriveListPageSize(pageSize: number | undefined): number {
    return Math.max(5, Math.min(100, pageSize ?? this.config.drive.pageSize));
  }

  private buildDriveListCacheKey(q: DriveListQuery, size: number): string {
    const { folderId, query, mimeIncludes = [], ownerAllowlist = [], pageToken, recursive = false, maxDepth = 20, dateFrom, dateTo, sizeMin, sizeMax, sortBy, sortDir, highlightChanges, sessionKey } = q;
    // Включаем текущую конфигурацию allowedMime в ключ кэша, чтобы разные фильтры MIME не делили один и тот же кэш
    const allowedMimeCfg = (this.config.drive.allowedMime && this.config.drive.allowedMime.length > 0)
      ? this.config.drive.allowedMime
      : ['*'];
    const allowedKey = allowedMimeCfg.join('.');
    return `drive:list:v3:${folderId}:${query ?? ''}:${mimeIncludes.join('.')}:${ownerAllowlist.join('.')}:${allowedKey}:${size}:${pageToken ?? ''}:${recursive}:${maxDepth}:${dateFrom ?? ''}:${dateTo ?? ''}:${sizeMin ?? ''}:${sizeMax ?? ''}:${sortBy ?? ''}:${sortDir ?? ''}:${highlightChanges ? '1' : '0'}:${sessionKey ?? ''}`;
  }

  private getDriveListFields(): string {
    return [
      'nextPageToken',
      "files(id,name,mimeType,size,modifiedTime,parents,webViewLink,iconLink,shortcutDetails(targetId,targetMimeType),owners(displayName,emailAddress))",
    ].join(',');
  }

  private buildFilesListParams(q: string, size: number, pageToken?: string): drive_v3.Params$Resource$Files$List {
    const params: drive_v3.Params$Resource$Files$List = {
      q,
      pageSize: size,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
      fields: this.getDriveListFields(),
      corpora: 'allDrives',
    };
    if (pageToken) params.pageToken = pageToken;
    return params;
  }

  private async executeDriveFilesList(params: drive_v3.Params$Resource$Files$List) {
    return this.executeWithRetry(async () => {
      if (!this.drive) throw new Error('Drive API не инициализовано');
      return this.drive.files.list(params);
    }, 'drive', undefined, 'drive.files.list');
  }

  private createDriveListResult(filesRaw: drive_v3.Schema$File[], ownerAllowlist: string[], nextPageToken?: string): DriveListResult {
    let files = filesRaw.map(f => this.toDriveFile(f));
    files = this.filterFilesByOwners(files, ownerAllowlist);
    const result: DriveListResult = { files };
    if (nextPageToken) result.nextPageToken = nextPageToken;
    return result;
  }

  /**
   * Пост-фильтр по списку разрешённых владельцев
   */
  private filterFilesByOwners(files: DriveFile[], ownerAllowlist?: string[]): DriveFile[] {
    if (!ownerAllowlist || ownerAllowlist.length === 0) return files;
    const allow = new Set(ownerAllowlist.map(s => s.toLowerCase()));
    return files.filter(f => (f.owners || []).some(o => allow.has(o.toLowerCase())));
  }

  /**
   * Кэширование результата листинга Drive с безопасной обработкой ошибок
   */
  private async cacheDriveListResult(cacheKey: string, result: DriveListResult): Promise<void> {
    try {
      const ttl = this.config.drive.ttlListSec ?? 60;
      await this.cacheService.set(cacheKey, result, ttl);
    } catch (/* istanbul ignore next */ _e) {
      // noop: кэширование не критично
    }
  }

  /**
   * Структурное логирование успешного листинга Drive
   */
  private logDriveListSuccess(args: { folderId: string; result: DriveListResult; duration: number; pageTokenPresent: boolean }): void {
    const { folderId, result, duration, pageTokenPresent } = args;
    logger.info('📄 Листинг Drive завершён', {
      type: 'api_request',
      event: 'drive_list_success',
      component: 'GoogleService',
      folderId,
      count: result.files.length,
      duration,
      pageToken: pageTokenPresent ? 'yes' : 'no',
      nextPageToken: result.nextPageToken ? 'yes' : 'no',
    });
  }

  /**
   * Сборка строки фильтра q для Drive API
   */
  private buildDriveQuery(
    folderId: string,
    nameContains: string,
    mimeIncludes: string[],
    allowedMimeCfg: string[],
    needMimeFilter: boolean,
    dateFrom?: string,
    dateTo?: string,
  ): string {
    const qParts: string[] = [
      `'${folderId}' in parents`,
      'trashed = false',
    ];
    if (nameContains) {
      // Простейший contains; Drive API чувствителен к регистру, но достаточно для начала
      const esc = nameContains.replace(/['\\]/g, '\\$&');
      qParts.push(`name contains '${esc}'`);
    }
    if (dateFrom) qParts.push(`modifiedTime >= '${dateFrom}'`);
    if (dateTo) qParts.push(`modifiedTime <= '${dateTo}'`);
    if (mimeIncludes && mimeIncludes.length > 0) {
      const ors = mimeIncludes.map(m => `mimeType='${m}'`).join(' or ');
      qParts.push(`(${ors})`);
    }
    if (needMimeFilter) {
      const ors = allowedMimeCfg.map(m => `mimeType='${m}'`).join(' or ');
      qParts.push(`(${ors})`);
    }
    return qParts.join(' and ');
  }

  /**
   * Считаем изменения между старым и новым списком файлов
   */
  private computeDriveListChanges(
    prev: DriveFile[],
    next: DriveFile[],
  ): {
    addedIds: string[];
    removedIds: string[];
    modified: Array<{
      id: string;
      fields: Array<'name' | 'mimeType' | 'size' | 'modifiedTime' | 'owners' | 'parents' | 'webViewLink'>;
    }>;
  } {
    const prevMap = new Map(prev.map(f => [f.id, f] as const));
    const nextMap = new Map(next.map(f => [f.id, f] as const));
    const addedIds: string[] = [];
    const removedIds: string[] = [];
    const modified: Array<{ id: string; fields: Array<'name' | 'mimeType' | 'size' | 'modifiedTime' | 'owners' | 'parents' | 'webViewLink'>; }> = [];

    for (const id of nextMap.keys()) if (!prevMap.has(id)) addedIds.push(id);
    for (const id of prevMap.keys()) if (!nextMap.has(id)) removedIds.push(id);

    const eqArr = (a?: string[], b?: string[]): boolean => {
      const aa = Array.isArray(a) ? [...a].sort() : [];
      const bb = Array.isArray(b) ? [...b].sort() : [];
      if (aa.length !== bb.length) return false;
      for (let i = 0; i < aa.length; i++) if (aa[i] !== bb[i]) return false;
      return true;
    };

    for (const [id, n] of nextMap) {
      const p = prevMap.get(id);
      if (!p) continue;
      const changed: Array<'name' | 'mimeType' | 'size' | 'modifiedTime' | 'owners' | 'parents' | 'webViewLink'> = [];
      if ((p.name || '') !== (n.name || '')) changed.push('name');
      if ((p.mimeType || '') !== (n.mimeType || '')) changed.push('mimeType');
      if ((p.size ?? -1) !== (n.size ?? -1)) changed.push('size');
      if ((p.modifiedTime || '') !== (n.modifiedTime || '')) changed.push('modifiedTime');
      if (!eqArr(p.owners, n.owners)) changed.push('owners');
      if (!eqArr(p.parents, n.parents)) changed.push('parents');
      if ((p.webViewLink || '') !== (n.webViewLink || '')) changed.push('webViewLink');
      if (changed.length) modified.push({ id, fields: changed });
    }

    return { addedIds, removedIds, modified };
  }

  /**
   * Стабильный хэш набора фильтров для ключа снапшота изменений
   */
  private hashFiltersForSnapshot(filters: {
    nameContains: string;
    mimeIncludes: string[];
    ownerAllowlist: string[];
    dateFrom?: string | undefined;
    dateTo?: string | undefined;
    sizeMin?: number | undefined;
    sizeMax?: number | undefined;
    sortBy?: string | undefined;
    sortDir?: string | undefined;
  }): string {
    const src = JSON.stringify(filters);
    return createHash('md5').update(src).digest('hex');
  }

  /**
   * Загрузка бинарного содержимого файла (Drive files.get alt=media → Buffer)
   */
  public async downloadFile(fileId: string): Promise<Buffer> {
    const cacheKey = `drive:file:bin:${fileId}`;
    try {
      const cached = await this.cacheService.get<Buffer>(cacheKey);
      if (cached) {
        this.stats.cacheHits++;
        return cached;
      }
    } catch {
      this.stats.cacheMisses++;
    }

    const start = Date.now();
    const resp = await this.executeWithRetry(async () => {
      if (!this.drive) throw new Error('Drive API не инициализовано');
      // Важно: просим поток
      const r = await this.drive.files.get({
        fileId,
        alt: 'media' as unknown as string,
      } as drive_v3.Params$Resource$Files$Get, { responseType: 'stream' });
      return r as unknown as { data: unknown };
    }, 'drive', undefined, 'drive.files.get.media');

    const buf = await this.dataToBuffer(resp.data as unknown as any);
    try {
      const ttl = this.config.drive.ttlTextSec ?? 300;
      await this.cacheService.set(cacheKey, buf, ttl);
    } catch (/* istanbul ignore next */ _e) {
      // noop: ошибка записи в кеш не критична
    }

    logger.info('⬇️ Файл загружен из Drive', {
      type: 'api_request',
      event: 'drive_download_success',
      component: 'GoogleService',
      fileId,
      size: buf.byteLength,
      duration: Date.now() - start,
    });

    return buf;
  }

  /**
   * Экспорт родных документов Google (Docs/Sheets/Slides) в другой MIME (pdf/txt/csv/…)
   */
  public async exportFile(fileId: string, mimeOut: string): Promise<Buffer> {
    const cacheKey = `drive:file:export:${fileId}:${mimeOut}`;
    try {
      const cached = await this.cacheService.get<Buffer>(cacheKey);
      if (cached) {
        this.stats.cacheHits++;
        return cached;
      }
    } catch {
      this.stats.cacheMisses++;
    }

    const start = Date.now();
    const resp = await this.executeWithRetry(async () => {
      if (!this.drive) throw new Error('Drive API не инициализовано');
      const r = await this.drive.files.export(
        { fileId, mimeType: mimeOut } as drive_v3.Params$Resource$Files$Export,
        { responseType: 'stream' },
      );
      return r as unknown as { data: unknown };
    }, 'drive', undefined, 'drive.files.export');

    const buf = await this.dataToBuffer(resp.data as unknown as any);
    try {
      const ttl = this.config.drive.ttlTextSec ?? 300;
      await this.cacheService.set(cacheKey, buf, ttl);
    } catch (/* istanbul ignore next */ _e) {
      // noop: ошибка записи в кеш не критична
    }

    logger.info('📤 Файл экспортирован из Drive', {
      type: 'api_request',
      event: 'drive_export_success',
      component: 'GoogleService',
      fileId,
      mimeOut,
      size: buf.byteLength,
      duration: Date.now() - start,
    });

    return buf;
  }

  /**
   * Расчёт и кеширование SHA-256 для файла Drive по fileId
   */
  public async getFileChecksum(fileId: string): Promise<string> {
    const cacheKey = `drive:file:checksum:${fileId}`;
    try {
      const cached = await this.cacheService.get<string>(cacheKey);
      if (cached) {
        this.stats.cacheHits++;
        return cached;
      }
    } catch {
      this.stats.cacheMisses++;
    }
    const buf = await this.downloadFile(fileId);
    const hash = createHash('sha256').update(buf).digest('hex');
    try {
      const ttl = this.config.drive.ttlTextSec ?? 300;
      await this.cacheService.set(cacheKey, hash, ttl);
    } catch { /* istanbul ignore next */ }
    return hash;
  }

  /**
   * Загрузка бинарного содержимого вместе с контрольной суммой
   */
  public async downloadFileWithHash(fileId: string): Promise<{ buffer: Buffer; checksum: string }> {
    const buffer = await this.downloadFile(fileId);
    const checksum = createHash('sha256').update(buffer).digest('hex');
    return { buffer, checksum };
  }

  /**
   * Экспорт файла вместе с контрольной суммой
   */
  public async exportFileWithHash(fileId: string, mimeOut: string): Promise<{ buffer: Buffer; checksum: string }> {
    const buffer = await this.exportFile(fileId, mimeOut);
    const checksum = createHash('sha256').update(buffer).digest('hex');
    return { buffer, checksum };
  }

  /**
   * Безопасно собрать поток в Buffer
   */
  private async streamToBuffer(stream: NodeJS.ReadableStream): Promise<Buffer> {
    return new Promise<Buffer>((resolve, reject) => {
      const chunks: Buffer[] = [];
      stream.on('data', (d: unknown) => {
        if (Buffer.isBuffer(d)) chunks.push(d);
        else if (d instanceof Uint8Array) chunks.push(Buffer.from(d));
        else if (typeof d === 'string') chunks.push(Buffer.from(d));
        else chunks.push(Buffer.from(String(d)));
      });
      stream.on('end', () => resolve(Buffer.concat(chunks)));
      stream.on('error', reject);
    });
  }

  private async dataToBuffer(data: unknown): Promise<Buffer> {
    // Если это уже Buffer
    if (Buffer.isBuffer(data)) return data;
    // Node stream
    if (data && typeof data === 'object' && typeof (data as any).on === 'function') {
      return this.streamToBuffer(data as unknown as NodeJS.ReadableStream);
    }
    // Uint8Array / ArrayBuffer
    if (data instanceof Uint8Array) return Buffer.from(data);
    if (typeof ArrayBuffer !== 'undefined' && data instanceof ArrayBuffer) return Buffer.from(new Uint8Array(data));
    // Строка
    if (typeof data === 'string') return Buffer.from(data);
    // Нечто JSON-подобное
    try {
      return Buffer.from(JSON.stringify(data ?? ''));
    } catch {
      return Buffer.from(String(data ?? ''));
    }
  }

  /**
   * OCR для локального буфера (без Drive)
   */
  public async extractTextFromBuffer(buf: Buffer): Promise<string> {
    try {
      const hash = createHash('sha1').update(buf).digest('hex');
      const cacheKey = `ocr:buffer:${hash}`;
      const ttl = this.config.google.ocrCacheTTL ?? this.config.performance.cacheTTL ?? 3600;

      const cached = await this.cacheService.get<string>(cacheKey);
      if (cached) {
        this.stats.cacheHits++;
        return cached;
      }
      this.stats.cacheMisses++;

      const provider = this.config.google.ocrProvider ?? 'vision';
      let text = '';
      if (provider === 'off') text = '';
      else if (provider === 'tesseract') text = await this.ocrWithTesseract(buf);
      else text = await this.ocrWithVision(buf);

      if (text && text.trim()) {
        await this.cacheService.set(cacheKey, text, ttl);
      }
      return text;
    } catch (error) {
      logger.error('❌ Помилка OCR для буфера', {
        type: 'processing_error',
        event: 'buffer_ocr_failed',
        component: 'GoogleService',
        error: error instanceof Error ? error.message : String(error),
      });
      return '';
    }
  }

  /**
   * OCR зображень з фича-флагом (Vision/Tesseract) та кэшуванням
   */
  public async extractTextFromImage(file: drive_v3.Schema$File): Promise<string> {
    try {
      const mime = file.mimeType || '';
      if (!/^image\//i.test(mime)) return '';
      if (!file.id) return '';
      const fileId = file.id;
      const buf = await this.downloadDriveFile(fileId);

      // Ключ кэша на основе fileId + modifiedTime (если есть) + хэш контента
      const modified = file.modifiedTime ?? '';
      const hash = createHash('sha1').update(buf).digest('hex');
      const cacheKey = `ocr:image:${fileId}:${modified}:${hash}`;

      const ttl = this.config.google.ocrCacheTTL ?? this.config.performance.cacheTTL ?? 3600;
      const cached = await this.cacheService.get<string>(cacheKey);
      if (cached) {
        this.stats.cacheHits++;
        return cached;
      }
      this.stats.cacheMisses++;

      const provider = this.config.google.ocrProvider ?? 'vision';
      let text = '';
      if (provider === 'off') {
        text = '';
      } else if (provider === 'tesseract') {
        text = await this.ocrWithTesseract(buf);
      } else {
        text = await this.ocrWithVision(buf);
      }

      if (text && text.trim().length > 0) {
        await this.cacheService.set(cacheKey, text, ttl);
      }
      return text;
    } catch (error) {
      logger.error('❌ Помилка OCR для зображення', {
        type: 'processing_error',
        event: 'image_ocr_failed',
        component: 'GoogleService',
        fileId: file.id,
        error: error instanceof Error ? error.message : String(error),
      });
      return '';
    }
  }

  /**
   * OCR через Google Vision (по умолчанию)
   */
  private async ocrWithVision(buf: Buffer): Promise<string> {
    type VisionEntityAnnotation = { description?: string | null };
    const client = new ImageAnnotatorClient();
    const [result] = await client.textDetection({ image: { content: buf } });
    const annotations = (result?.textAnnotations ?? []) as VisionEntityAnnotation[];
    const first = annotations[0];
    if (first?.description && typeof first.description === 'string') {
      return first.description;
    }
    const descriptions = annotations
      .map(a => (typeof a?.description === 'string' ? a.description : undefined))
      .filter((d): d is string => typeof d === 'string' && d.length > 0);
    return descriptions.join('\n');
  }

  /**
   * OCR через офлайн Tesseract
   * Требует установленные зависимости tesseract.js и tesseract.js-node,
   * а также локальные traineddata (config.google.tesseractLangPath)
   */
  private async ocrWithTesseract(buf: Buffer): Promise<string> {
    try {
      // Динамический импорт, чтобы не ломать окружение без зависимости
      const mod = await import('tesseract.js');
      type TesseractWorker = {
        loadLanguage(langs: string): Promise<void>;
        initialize(langs: string): Promise<void>;
        recognize(image: Buffer): Promise<{ data: { text?: string } }>;
        terminate(): Promise<void>;
      };
      type CreateWorker = (opts?: { langPath?: string; cachePath?: string; logger?: (m: unknown) => void }) => Promise<TesseractWorker>;
      const createWorker = mod.createWorker as unknown as CreateWorker;

      const langs = this.config.google.tesseractLangs || 'eng';
      const langPath = this.config.google.tesseractLangPath;

      const workerOpts: { langPath?: string; logger?: (m: unknown) => void } = {};
      if (typeof langPath === 'string') workerOpts.langPath = langPath;
      // workerOpts.logger = m => logger.debug('tesseract', { progress: (m as any).progress });
      const worker = await createWorker(workerOpts);

      try {
        await worker.loadLanguage(langs);
        await worker.initialize(langs);

        const { data } = await worker.recognize(buf);
        const text = typeof data?.text === 'string' ? data.text : '';
        return text;
      } finally {
        await worker.terminate();
      }
    } catch (e) {
      logger.warn('⚠️ OCR (tesseract) не доступен или завершился с ошибкой', {
        type: 'system',
        event: 'tesseract_ocr_unavailable',
        component: 'GoogleService',
        errorMessage: e instanceof Error ? e.message : String(e),
      });
      return '';
    }
  }

  /**
   * Отримати метадані файлу Google Drive
   */
  public async getDriveFileMetadata(fileId: string): Promise<drive_v3.Schema$File> {
    try {
      const file = await this.executeWithRetry(async () => {
        if (!this.drive) throw new Error('Drive API не ініціалізовано');
        const res = await this.drive.files.get({
          fileId,
          fields: 'id,name,mimeType,size,modifiedTime,owners(displayName),parents',
        });
        return res.data;
      }, 'drive', undefined, 'drive.files.get');
      return file;
    } catch (error) {
      logger.error('❌ Помилка отримання метаданих файлу Drive', {
        type: 'api_error',
        event: 'drive_get_metadata_failed',
        component: 'GoogleService',
        fileId,
        error: String(error),
      });
      throw error;
    }
  }

  /**
   * Завантаження бінарного файлу з Drive (не Google Docs/Sheets типи)
   */
  public async downloadDriveFile(fileId: string): Promise<Buffer> {
    try {
      const data = await this.executeWithRetry(async () => {
        if (!this.drive) throw new Error('Drive API не ініціалізовано');
        const res = await this.drive.files.get(
          { fileId, alt: 'media' },
          { responseType: 'arraybuffer' }
        );
        return Buffer.from(res.data as any);
      }, 'drive', undefined, 'drive.files.get.media');
      return data;
    } catch (error) {
      logger.error('❌ Помилка завантаження файла Drive', {
        type: 'api_error',
        event: 'drive_download_failed',
        component: 'GoogleService',
        fileId,
        error: String(error),
      });
      throw error;
    }
  }

  /**
   * Експорт файлу Google Docs/Sheets/Slides у вказаний MIME тип
   */
  public async exportDriveFile(fileId: string, mimeType: string): Promise<Buffer> {
    try {
      const data = await this.executeWithRetry(async () => {
        if (!this.drive) throw new Error('Drive API не ініціалізовано');
        const res = await this.drive.files.export(
          { fileId, mimeType },
          { responseType: 'arraybuffer' }
        );
        return Buffer.from(res.data as any);
      }, 'drive', undefined, 'drive.files.export');
      return data;
    } catch (error) {
      logger.error('❌ Помилка експорту файла Drive', {
        type: 'api_error',
        event: 'drive_export_failed',
        component: 'GoogleService',
        fileId,
        mimeType,
        error: String(error),
      });
      throw error;
    }
  }

  /**
   * Ініціалізація Google сервісів
   */
  protected async onInitialize(): Promise<void> {
    try {
      logger.info('🔧 Ініціалізація Google Service...', {
        type: 'system',
        event: 'google_service_init',
        component: 'GoogleService',
      });

      // Ініціалізація кешу
      await this.cacheService.initialize();

      // Створення автентифікації
      await this.initializeAuth();

      // Ініціалізація API клієнтів
      this.initializeAPIs();

      // Створення connection pool
      this.initializeConnectionPool();

      logger.info('✅ Google Service ініціалізовано', {
        type: 'system',
        event: 'google_service_init_success',
        component: 'GoogleService',
      });
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка ініціалізації Google Service', {
          type: 'system',
          event: 'google_service_init_failed',
          component: 'GoogleService',
          errorName: error.name,
          errorMessage: error.message,
          stack: error.stack,
          severity: 'critical',
        });
      } else {
        logger.error('❌ Помилка ініціалізації Google Service', {
          type: 'system',
          event: 'google_service_init_failed',
          component: 'GoogleService',
          errorMessage: String(error),
          severity: 'critical',
        });
      }
      throw error;
    }
  }

  /**
   * Список файлів у папці Google Drive з опціями рекурсії та фільтрації
   */
  public async listDriveFilesInFolder(
    folderId: string,
    opts: {
      recursive?: boolean;
      type?: 'sheet' | 'folder' | 'any';
      query?: string; // частина імені файла (case-insensitive contains)
      limit?: number;
      pageToken?: string;
      maxDepth?: number; // обмеження глибини рекурсії
    } = {}
  ): Promise<drive_v3.Schema$File[]> {
    const {
      recursive = false,
      type = 'any',
      query = '',
      limit = 100,
      pageToken,
      maxDepth = 20,
    } = opts;

    const cacheKey = `drive:list:${folderId}:${type}:${query}:${recursive}:${maxDepth}:${pageToken ?? ''}:${limit}`;

    // Кеш
    try {
      const cached = await this.cacheService.get<drive_v3.Schema$File[]>(cacheKey);
      if (cached) {
        this.stats.cacheHits++;
        logger.debug('✅ Використано кешовані дані Drive list', {
          type: 'system',
          event: 'cache_hit',
          component: 'GoogleService',
          folderId,
        });
        return cached;
      }
    } catch (e) {
      this.stats.cacheMisses++;
    }

    // Побудова запитів
    const { qFiles, qFolders } = this.buildListInFolderQueries(folderId, type, query);

    const filesHere = await this.fetchAllDriveList(qFiles, limit, pageToken);

    // Якщо шукаємо таблиці — конвертуємо ярлики на таблиці у "віртуальні" записи з targetId та
    // додаємо Excel-файли у результат
    const results: drive_v3.Schema$File[] = type === 'sheet' ? this.mapSheetFiles(filesHere) : [...filesHere];

    if (recursive && (maxDepth > 0 || maxDepth <= -1)) {
      // Отримуємо підпапки в поточній папці (включаючи ярлики на папки)
      const foldersLevel = await this.fetchAllDriveList(qFolders, limit, pageToken);
      const folders = this.extractFoldersFromLevel(foldersLevel);

      for (const folder of folders) {
        try {
          const sub = await this.listDriveFilesInFolder(folder.id, {
            recursive: true,
            type,
            query,
            limit,
            maxDepth: maxDepth <= -1 ? -1 : maxDepth - 1,
          });
          results.push(...sub);
        } catch (err) {
          logger.warn('⚠️ Неможливо отримати вміст підпапки', {
            type: 'api_error',
            event: 'drive_list_subfolder_failed',
            component: 'GoogleService',
            folderId: folder.id,
          });
        }
      }
    }

    // Додаткове логування, якщо результати порожні — допомагає діагностувати доступ/вміст
    if (results.length === 0) {
      logger.info('ℹ️ Drive list: порожній результат', {
        type: 'system',
        event: 'drive_list_empty',
        component: 'GoogleService',
        folderId,
        query,
        recursive,
        maxDepth,
        filterType: type,
      });
    }

    // Кешуємо на короткий термін (30с), щоб не перевищувати квоти
    try {
      await this.cacheService.set(cacheKey, results, 30);
    } catch (/* istanbul ignore next */ _e) {
      // noop: ошибка записи в кеш не критична
    }

    return results;
  }

  /**
   * Построение запросов files/folders для listDriveFilesInFolder
   */
  private buildListInFolderQueries(
    folderId: string,
    type: 'sheet' | 'folder' | 'any',
    query: string
  ): { qFiles: string; qFolders: string } {
    // MIME фильтры
    const filesMimeFilter =
      type === 'sheet'
        ?
            " and (" +
            [
              "mimeType='application/vnd.google-apps.spreadsheet'",
              "mimeType='application/vnd.google-apps.shortcut'",
              "mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'",
              "mimeType='application/vnd.ms-excel'",
            ].join(' or ') +
            ")"
        : type === 'folder'
          ? " and mimeType='application/vnd.google-apps.folder'"
          : '';
    const foldersMimeFilter =
      " and (mimeType='application/vnd.google-apps.folder' or mimeType='application/vnd.google-apps.shortcut')";

    // name contains фильтр
    const nameFilter = query ? ` and name contains '${query.replace(/'/g, "\\'")}'` : '';

    const base = `'${folderId}' in parents and trashed=false`;
    const qFiles = `${base}${filesMimeFilter}${nameFilter}`;
    const qFolders = `${base}${foldersMimeFilter}`;
    return { qFiles, qFolders };
  }

  /**
   * Выбрать все страницы результата files.list по запросу q
   */
  private async fetchAllDriveList(
    q: string,
    limit: number,
    startPageToken?: string
  ): Promise<drive_v3.Schema$File[]> {
    const all: drive_v3.Schema$File[] = [];
    let token: string | undefined = startPageToken;
    do {
      const resp = await this.executeWithRetry(async () => {
        if (!this.drive) throw new Error('Drive API не ініціалізовано');
        const params: drive_v3.Params$Resource$Files$List = {
          q,
          fields:
            'nextPageToken, files(id,name,mimeType,size,modifiedTime,parents,shortcutDetails(targetId,targetMimeType))',
          pageSize: Math.min(limit, 1000),
          ...(token ? { pageToken: token } : {}),
        };
        return await this.drive.files.list(params);
      }, 'drive', undefined, 'drive.files.list');
      const files = resp.data.files || [];
      all.push(...files);
      token = resp.data.nextPageToken || undefined;
    } while (token && all.length < limit);
    return all;
  }

  /**
   * Нормализация списка файлов под тип 'sheet': разворачивает ярлыки на spreadsheets, добавляет Excel
   */
  private mapSheetFiles(files: drive_v3.Schema$File[]): drive_v3.Schema$File[] {
    return files
      .map(f => {
        if (f.mimeType === 'application/vnd.google-apps.shortcut') {
          const targetId = f.shortcutDetails?.targetId;
          const targetMime = f.shortcutDetails?.targetMimeType;
          if (typeof targetId === 'string' && targetMime === 'application/vnd.google-apps.spreadsheet') {
            return {
              id: targetId,
              name: f.name,
              mimeType: 'application/vnd.google-apps.spreadsheet',
              parents: f.parents,
            } as drive_v3.Schema$File;
          }
          return null;
        }
        const isGs = f.mimeType === 'application/vnd.google-apps.spreadsheet';
        const isXlsx = f.mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
        const isXls = f.mimeType === 'application/vnd.ms-excel';
        return isGs || isXlsx || isXls ? f : null;
      })
      .filter((x): x is drive_v3.Schema$File => Boolean(x));
  }

  /**
   * Выделение подпапок из уровня результатов, включая ярлыки на папки
   */
  private extractFoldersFromLevel(level: drive_v3.Schema$File[]): Array<{ id: string; name?: string }> {
    return level
      .map<{ id: string; name?: string } | null>(f => {
        if (f.mimeType === 'application/vnd.google-apps.folder') {
          if (!f.id) return null;
          return f.name ? { id: f.id, name: f.name } : { id: f.id };
        }
        if (
          f.mimeType === 'application/vnd.google-apps.shortcut' &&
          f.shortcutDetails?.targetMimeType === 'application/vnd.google-apps.folder'
        ) {
          const targetId = f.shortcutDetails?.targetId;
          if (typeof targetId === 'string') return f.name ? { id: targetId, name: f.name } : { id: targetId };
        }
        return null;
      })
      .filter((x): x is { id: string; name?: string } => x !== null);
  }

  /**
   * Отримати список листів у таблиці
   */
  public async listSheets(spreadsheetId: string): Promise<string[]> {
    const cacheKey = `sheets:list:${spreadsheetId}`;
    try {
      const cached = await this.cacheService.get<string[]>(cacheKey);
      if (cached) {
        this.stats.cacheHits++;
        return cached;
      }
    } catch {
      this.stats.cacheMisses++;
    }

    const titles = await this.executeWithRetry(async () => {
      if (!this.sheets) throw new Error('Sheets API не ініціалізовано');
      const res = await this.sheets.spreadsheets.get({ spreadsheetId });
      const sheets = res.data.sheets || [];
      return sheets.map(s => s.properties?.title || '').filter(Boolean);
    }, 'sheets', undefined, 'sheets.spreadsheets.get');

    try {
      await this.cacheService.set(cacheKey, titles, 60);
    } catch (/* istanbul ignore next */ _e) {
      // noop: ошибка записи в кеш не критична
    }
    return titles;
  }

  /**
   * Знайти таблиці за ім'ям у папці (з опцією рекурсії)
   */
  public async findSpreadsheetsByNameInFolder(
    namePart: string,
    rootFolderId: string,
    recursive: boolean = true,
    maxDepth: number = 3
  ): Promise<drive_v3.Schema$File[]> {
    const files = await this.listDriveFilesInFolder(rootFolderId, {
      recursive,
      maxDepth,
      type: 'sheet',
      query: namePart,
      limit: 500,
    });
    // Повертаємо тільки spreadsheets
    return files.filter(f => f.mimeType === 'application/vnd.google-apps.spreadsheet');
  }

  /**
   * Ініціалізація автентифікації
   */
  private async initializeAuth(): Promise<void> {
    try {
      // Перевірка наявності credentials
      if (!this.config.google.credentials) {
        throw new Error('Google credentials не налаштовано');
      }

      // Створення JWT автентифікації
      const jwt = new google.auth.JWT(
        this.config.google.credentials.client_email,
        undefined,
        this.config.google.credentials.private_key,
        [
          'https://www.googleapis.com/auth/spreadsheets',
          'https://www.googleapis.com/auth/drive',
          'https://www.googleapis.com/auth/documents',
        ]
      );

      // Авторизація
      await jwt.authorize();
      this.auth = jwt;
      logger.info('✅ Google автентифікація успішна', {
        type: 'system',
        event: 'google_auth_success',
        component: 'GoogleService',
      });
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка Google автентифікації', {
          type: 'api_error',
          event: 'google_auth_failed',
          component: 'GoogleService',
          errorName: error.name,
          errorMessage: error.message,
          stack: error.stack,
          service: 'google',
        });
      } else {
        logger.error('❌ Помилка Google автентифікації', {
          type: 'api_error',
          event: 'google_auth_failed',
          component: 'GoogleService',
          service: 'google',
          errorMessage: String(error),
        });
      }
      throw error;
    }
  }

  /**
   * Ініціалізація API клієнтів
   */
  private initializeAPIs(): void {
    try {
      if (!this.auth) throw new Error('Auth client is not initialized');
      // Google Sheets API
      this.sheets = google.sheets({ version: 'v4', auth: this.auth });

      // Google Drive API
      this.drive = google.drive({ version: 'v3', auth: this.auth });

      // Google Docs API
      this.docs = google.docs({ version: 'v1', auth: this.auth });

      logger.info('✅ Google API клієнти ініціалізовано', {
        type: 'system',
        event: 'google_api_init_success',
        component: 'GoogleService',
      });
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка ініціалізації Google API', {
          type: 'api_error',
          event: 'google_api_init_failed',
          component: 'GoogleService',
          errorName: error.name,
          errorMessage: error.message,
          stack: error.stack,
          service: 'google',
        });
      } else {
        logger.error('❌ Помилка ініціалізації Google API', {
          type: 'api_error',
          event: 'google_api_init_failed',
          component: 'GoogleService',
          service: 'google',
          errorMessage: String(error),
        });
      }
      throw error;
    }
  }

  /**
   * Ініціалізація Connection Pool
   */
  private initializeConnectionPool(): void {
    try {
      const apiTypes = ['sheets', 'drive', 'docs'];

      for (const apiType of apiTypes) {
        this.connectionPool.set(apiType, {
          inUse: false,
          lastUsed: Date.now(),
          requestCount: 0,
        });
      }

      logger.info('✅ Connection Pool ініціалізовано', {
        type: 'system',
        event: 'connection_pool_init_success',
        component: 'GoogleService',
      });
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка ініціалізації Connection Pool', {
          type: 'system',
          event: 'connection_pool_init_failed',
          component: 'GoogleService',
          errorName: error.name,
          errorMessage: error.message,
          stack: error.stack,
        });
      } else {
        logger.error('❌ Помилка ініціалізації Connection Pool', {
          type: 'system',
          event: 'connection_pool_init_failed',
          component: 'GoogleService',
          errorMessage: String(error),
        });
      }
      throw error;
    }
  }

  /**
   * Отримання з'єднання з пулу
   */
  private getConnection(apiType: string): boolean {
    const connection = this.connectionPool.get(apiType);
    if (!connection) {
      return false;
    }

    if (connection.inUse) {
      return false;
    }

    connection.inUse = true;
    connection.lastUsed = Date.now();
    connection.requestCount++;
    return true;
  }

  /**
   * Звільнення з'єднання
   */
  private releaseConnection(apiType: string): void {
    const connection = this.connectionPool.get(apiType);
    if (connection) {
      connection.inUse = false;
    }
  }

  /**
   * Виконання операції з retry
   */
  private async executeWithRetry<T>(
    operation: () => Promise<T>,
    apiType: string,
    maxRetries: number = this.retryAttempts,
    endpoint: string = 'unknown'
  ): Promise<T> {
    let lastError: Error | null = null;

    for (let attempt = 0; attempt <= maxRetries; attempt++) {
      try {
        // Rate-limit перед взятием соединения
        await this.throttle(apiType);

        const connection = this.getConnection(apiType);
        if (!connection) {
          throw new Error(`Немає доступних з'єднань для ${apiType}`);
        }

        const startTime = Date.now();
        const result = await operation();
        const duration = Date.now() - startTime;

        this.releaseConnection(apiType);
        this.updateStats(true, duration);

        // Метрики: успешный запрос
        try {
          this.metrics?.updateGoogleApiMetrics(apiType, endpoint, 'success', duration);
        } catch (/* istanbul ignore next */ _e) {
          // noop: метрики не критичны
        }

        return result;
      } catch (error) {
        lastError = error as Error;
        this.releaseConnection(apiType);
        this.updateStats(false, 0);

        // Метрики: ошибка попытки
        try {
          const errDuration = 0;
          this.metrics?.updateGoogleApiMetrics(apiType, endpoint, 'error', errDuration);
        } catch (/* istanбул ignore next */ _e) {
          // noop: ошибка записи в кеш не критична
        }

        if (attempt < maxRetries) {
          const delay = this.retryDelay * Math.pow(2, attempt);
          await new Promise(resolve => setTimeout(resolve, delay));
        }
      }
    }

    throw lastError || new Error('Всі спроби виконання невдалі');
  }

  /**
   * Отримання даних з Google Sheets
   */
  public async getSheetData(
    spreadsheetId: string,
    range: string,
    options: GoogleServiceOptions = {}
  ): Promise<SheetData> {
    const { useCache = true, cacheTTL = 300, forceRefresh = false } = options;

    try {
      const normRange = this.getSheetsService().normalizeRange(range);
      // Перевірка кешу
      if (useCache && !forceRefresh) {
        const cacheKey = `sheets:${spreadsheetId}:${normRange}`;
        try {
          const cached = await this.cacheService.get<SheetData>(cacheKey);
          if (cached) {
            this.stats.cacheHits++;
            logger.debug('✅ Використано кешовані дані Sheets', {
              type: 'system',
              event: 'cache_hit',
              component: 'GoogleService',
              spreadsheetId: spreadsheetId.substring(0, 10) + '...',
              range: normRange,
              rowsCount: cached.values.length,
            });
            return cached;
          } else {
            this.stats.cacheMisses++;
          }
        } catch (cacheError) {
          if (cacheError instanceof Error) {
            logger.warn('⚠️ Помилка читання з кешу Sheets', {
              type: 'system',
              event: 'cache_read_failed',
              component: 'CacheService',
              errorName: cacheError.name,
              errorMessage: cacheError.message,
              stack: cacheError.stack,
            });
          } else {
            logger.warn('⚠️ Помилка читання з кешу Sheets', {
              type: 'system',
              event: 'cache_read_failed',
              component: 'CacheService',
              errorMessage: String(cacheError),
            });
          }
          this.stats.cacheMisses++;
        }
      }

      const result = await this.executeWithRetry(async () => {
        if (!this.sheets) throw new Error('Sheets API не ініціалізовано');

        const response = await this.sheets.spreadsheets.values.get({
          spreadsheetId,
          range: normRange,
        });

        // Унифицированная нормализация через SheetsService
        return this.getSheetsService().toSheetDataFromGet(response.data, normRange);
      }, 'sheets', undefined, 'sheets.spreadsheets.values.get');

      // Збереження в кеш
      if (useCache) {
        const cacheKey = `sheets:${spreadsheetId}:${normRange}`;
        try {
          // CacheService expects TTL in seconds; do not convert to ms
          await this.cacheService.set(cacheKey, result, cacheTTL);
          logger.debug('💾 Дані Sheets збережено в кеш', {
            type: 'system',
            event: 'cache_write',
            component: 'GoogleService',
            spreadsheetId: spreadsheetId.substring(0, 10) + '...',
            range: normRange,
            rowsCount: result.values.length,
            ttl: `${cacheTTL}s`,
          });
        } catch (cacheError) {
          if (cacheError instanceof Error) {
            logger.warn('⚠️ Помилка збереження в кеш Sheets', {
              type: 'system',
              event: 'cache_write_failed',
              component: 'CacheService',
              errorName: cacheError.name,
              errorMessage: cacheError.message,
              stack: cacheError.stack,
            });
          } else {
            logger.warn('⚠️ Помилка збереження в кеш Sheets', {
              type: 'system',
              event: 'cache_write_failed',
              component: 'CacheService',
              errorMessage: String(cacheError),
            });
          }
        }
      }

      return result;
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка отримання даних з Sheets', {
          type: 'api_error',
          event: 'sheets_get_failed',
          component: 'GoogleService',
          errorName: error.name,
          errorMessage: error.message,
          stack: error.stack,
          service: 'sheets',
          spreadsheetId,
          range,
        });
      } else {
        logger.error('❌ Помилка отримання даних з Sheets', {
          type: 'api_error',
          event: 'sheets_get_failed',
          component: 'GoogleService',
          service: 'sheets',
          spreadsheetId,
          range,
          errorMessage: String(error),
        });
      }
      throw error;
    }
  }

  /**
   * Запис даних в Google Sheets
   */
  public async writeSheetData(
    spreadsheetId: string,
    range: string,
    values: string[][],
    options: GoogleServiceOptions = {}
  ): Promise<void> {
    const { valueInputOption = 'RAW', clearCache = true } = options;

    try {
      const normRange = this.getSheetsService().normalizeRange(range);
      const normValues = this.getSheetsService().normalizeWriteValues(values);
      await this.executeWithRetry(async () => {
        if (!this.sheets) throw new Error('Sheets API не ініціалізовано');

        await this.sheets.spreadsheets.values.update({
          spreadsheetId,
          range: normRange,
          valueInputOption,
          requestBody: {
            values: normValues,
          },
        });
      }, 'sheets', undefined, 'sheets.spreadsheets.values.update');

      // Очищення кешу
      if (clearCache) {
        const cacheKey = `sheets:${spreadsheetId}:${normRange}`;
        try {
          await this.cacheService.delete(cacheKey);
          logger.debug('🗑️ Кеш Sheets очищено', {
            type: 'system',
            event: 'cache_delete',
            component: 'GoogleService',
            spreadsheetId: spreadsheetId.substring(0, 10) + '...',
            range: normRange,
          });
        } catch (cacheError) {
          if (cacheError instanceof Error) {
            logger.warn('⚠️ Помилка очищення кешу Sheets', {
              type: 'system',
              event: 'cache_delete_failed',
              component: 'CacheService',
              errorName: cacheError.name,
              errorMessage: cacheError.message,
              stack: cacheError.stack,
            });
          } else {
            logger.warn('⚠️ Помилка очищення кешу Sheets', {
              type: 'system',
              event: 'cache_delete_failed',
              component: 'CacheService',
              errorMessage: String(cacheError),
            });
          }
        }
      }
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка запису в Sheets', {
          type: 'api_error',
          event: 'sheets_write_failed',
          component: 'GoogleService',
          errorName: error.name,
          errorMessage: error.message,
          stack: error.stack,
          service: 'sheets',
          spreadsheetId,
          range,
        });
      } else {
        logger.error('❌ Помилка запису в Sheets', {
          type: 'api_error',
          event: 'sheets_write_failed',
          component: 'GoogleService',
          service: 'sheets',
          spreadsheetId,
          range,
          errorMessage: String(error),
        });
      }
      throw error;
    }
  }

  /**
   * Batch отримання даних з Google Sheets
   */
  public async batchGetSheetData(
    spreadsheetId: string,
    ranges: string[],
    options: GoogleServiceOptions = {}
  ): Promise<BatchSheetData> {
    const { batchSize = 10, retryFailed = true, maxRetries = 3 } = options;

    try {
      const normRanges = ranges.map(r => this.getSheetsService().normalizeRange(r));
      const chunks = this.chunkArray(normRanges, batchSize);
      const results: SheetData[] = [];
      const failedRanges: string[] = [];

      for (const chunk of chunks) {
        try {
          const result = await this.executeWithRetry(
            async () => {
              if (!this.sheets) throw new Error('Sheets API не ініціалізовано');

              const response = await this.sheets.spreadsheets.values.batchGet({
                spreadsheetId,
                ranges: chunk,
              });

              // Делегируем парсинг структуре SheetsService
              const parsed = this.getSheetsService().parseBatchGet(response.data);
              return parsed.valueRanges;
            },
            'sheets',
            maxRetries,
            'sheets.spreadsheets.values.batchGet'
          );

          results.push(
            ...result.map(vr => ({
              range: vr.range,
              majorDimension: 'ROWS',
              values: (vr.values || []).map(row => row.map(cell => (cell == null ? '' : String(cell)))),
            }))
          );
        } catch (error) {
          if (error instanceof Error) {
            logger.error('❌ Помилка batch запиту', {
              type: 'api_error',
              event: 'sheets_batch_get_failed',
              errorName: error.name,
              errorMessage: error.message,
              stack: error.stack,
              service: 'sheets',
              spreadsheetId,
              ranges: chunk,
            });
          } else {
            logger.error('❌ Помилка batch запиту', {
              type: 'api_error',
              event: 'sheets_batch_get_failed',
              service: 'sheets',
              spreadsheetId,
              ranges: chunk,
              errorMessage: String(error),
            });
          }
          if (retryFailed) {
            failedRanges.push(...chunk);
          }
        }
      }

      // Повторна спроба для невдалих ranges
      if (retryFailed && failedRanges.length > 0) {
        for (const range of failedRanges) {
          try {
            const result = await this.getSheetData(spreadsheetId, range, { useCache: false });
            results.push(result);
          } catch (error) {
            if (error instanceof Error) {
              logger.error(`❌ Повторна спроба невдала для range: ${range}`, {
                type: 'api_error',
                event: 'sheets_retry_get_failed',
                errorName: error.name,
                errorMessage: error.message,
                stack: error.stack,
                service: 'sheets',
                spreadsheetId,
                range,
              });
            } else {
              logger.error(`❌ Повторна спроба невдала для range: ${range}`, {
                type: 'api_error',
                event: 'sheets_retry_get_failed',
                service: 'sheets',
                spreadsheetId,
                range,
                errorMessage: String(error),
              });
            }
          }
        }
      }

      return {
        valueRanges: results,
        spreadsheetId,
      };
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка batch отримання даних', {
          type: 'api_error',
          event: 'sheets_batch_get_failed_final',
          errorName: error.name,
          errorMessage: error.message,
          stack: error.stack,
          service: 'sheets',
          spreadsheetId,
        });
      } else {
        logger.error('❌ Помилка batch отримання даних', {
          type: 'api_error',
          event: 'sheets_batch_get_failed_final',
          service: 'sheets',
          spreadsheetId,
          errorMessage: String(error),
        });
      }
      throw error;
    }
  }

  /**
   * Batch запис даних в Google Sheets
   */
  public async batchWriteSheetData(
    spreadsheetId: string,
    data: Array<{ range: string; values: string[][] }>,
    options: GoogleServiceOptions = {}
  ): Promise<void> {
    const { batchSize = 10, retryFailed = true, maxRetries = 3, clearCache = true, valueInputOption = 'RAW' } = options;

    try {
      const normData = data.map(d => ({ range: this.getSheetsService().normalizeRange(d.range), values: d.values }));
      const chunks = this.chunkArray(normData, batchSize);
      const failedBatches: Array<{ range: string; values: string[][] }> = [];

      for (const chunk of chunks) {
        try {
          await this.executeWithRetry(
            async () => {
              if (!this.sheets) throw new Error('Sheets API не ініціалізовано');

              // Використовуємо values.batchUpdate, щоб не залежати від sheetId
              const requestBody = this.getSheetsService().buildBatchUpdate({
                valueInputOption: valueInputOption === 'USER_ENTERED' ? 'USER_ENTERED' : 'RAW',
                data: chunk.map(it => ({ range: it.range, values: this.getSheetsService().normalizeWriteValues(it.values) })),
              });

              await this.sheets.spreadsheets.values.batchUpdate({
                spreadsheetId,
                requestBody,
              });
            },
            'sheets',
            maxRetries,
            'sheets.spreadsheets.values.batchUpdate'
          );

          // Очищення кешу
          if (clearCache) {
            for (const item of chunk) {
              const cacheKey = `sheets:${spreadsheetId}:${item.range}`;
              try {
                await this.cacheService.delete(cacheKey);
                logger.debug('🗑️ Кеш Sheets очищено', {
                  type: 'system',
                  event: 'cache_delete',
                  component: 'GoogleService',
                  spreadsheetId: spreadsheetId.substring(0, 10) + '...',
                  range: item.range,
                });
              } catch (cacheError) {
                if (cacheError instanceof Error) {
                  logger.warn('⚠️ Помилка очищення кешу Sheets', {
                    type: 'system',
                    event: 'cache_delete_failed',
                    component: 'CacheService',
                    errorName: cacheError.name,
                    errorMessage: cacheError.message,
                    stack: cacheError.stack,
                  });
                } else {
                  logger.warn('⚠️ Помилка очищення кешу Sheets', {
                    type: 'system',
                    event: 'cache_delete_failed',
                    component: 'CacheService',
                    errorMessage: String(cacheError),
                  });
                }
              }
            }
          }
        } catch (error) {
          if (error instanceof Error) {
            logger.error('❌ Помилка batch запису', {
              type: 'api_error',
              event: 'sheets_batch_write_failed',
              component: 'GoogleService',
              errorName: error.name,
              errorMessage: error.message,
              stack: error.stack,
              service: 'sheets',
              spreadsheetId,
            });
          } else {
            logger.error('❌ Помилка batch запису', {
              type: 'api_error',
              event: 'sheets_batch_write_failed',
              component: 'GoogleService',
              service: 'sheets',
              spreadsheetId,
              errorMessage: String(error),
            });
          }
          if (retryFailed) {
            failedBatches.push(...chunk);
          }
        }
      }

      // Повторна спроба для невдалих batch
      if (retryFailed && failedBatches.length > 0) {
        for (const item of failedBatches) {
          try {
            await this.writeSheetData(spreadsheetId, item.range, item.values, { useCache: false });
          } catch (error) {
            if (error instanceof Error) {
              logger.error(`❌ Повторна спроба невдала для range: ${item.range}`, {
                type: 'api_error',
                event: 'sheets_retry_write_failed',
                component: 'GoogleService',
                errorName: error.name,
                errorMessage: error.message,
                stack: error.stack,
                service: 'sheets',
                spreadsheetId,
                range: item.range,
              });
            } else {
              logger.error(`❌ Повторна спроба невдала для range: ${item.range}`, {
                type: 'api_error',
                event: 'sheets_retry_write_failed',
                component: 'GoogleService',
                service: 'sheets',
                spreadsheetId,
                range: item.range,
                errorMessage: String(error),
              });
            }
          }
        }
      }
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка batch запису даних', {
          type: 'api_error',
          event: 'sheets_batch_write_failed_final',
          component: 'GoogleService',
          errorName: error.name,
          errorMessage: error.message,
          stack: error.stack,
          service: 'sheets',
          spreadsheetId,
        });
      } else {
        logger.error('❌ Помилка batch запису даних', {
          type: 'api_error',
          event: 'sheets_batch_write_failed_final',
          component: 'GoogleService',
          service: 'sheets',
          spreadsheetId,
          errorMessage: String(error),
        });
      }
      throw error;
    }
  }

  /**
   * Пошук файлів в Google Drive
   */
  public async searchFiles(
    query: string,
    _options: GoogleServiceOptions = {}
  ): Promise<drive_v3.Schema$File[]> {
    try {
      const result = await this.executeWithRetry(async () => {
        if (!this.drive) throw new Error('Drive API не ініціалізовано');

        const response = await this.drive.files.list({
          q: query,
          fields: 'files(id,name,mimeType,size,modifiedTime)',
          pageSize: 100,
        });

        return response.data.files || [];
      }, 'drive', undefined, 'drive.files.list');

      return result;
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка пошуку файлів', {
          type: 'api_error',
          event: 'drive_search_failed',
          component: 'GoogleService',
          errorName: error.name,
          errorMessage: error.message,
          stack: error.stack,
          service: 'drive',
          query,
        });
      } else {
        logger.error('❌ Помилка пошуку файлів', {
          type: 'api_error',
          event: 'drive_search_failed',
          component: 'GoogleService',
          service: 'drive',
          query,
          errorMessage: String(error),
        });
      }
      throw error;
    }
  }

  /**
   * Отримання метаданих файлу
   */
  public async getFileMetadata(
    fileId: string,
    fields: string = '*'
  ): Promise<drive_v3.Schema$File> {
    try {
      const result = await this.executeWithRetry(async () => {
        if (!this.drive) throw new Error('Drive API не ініціалізовано');

        const response = await this.drive.files.get({
          fileId,
          fields,
        });

        return response.data;
      }, 'drive', undefined, 'drive.files.get');

      return result;
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка отримання метаданих файлу', {
          type: 'api_error',
          event: 'drive_metadata_failed',
          component: 'GoogleService',
          errorName: error.name,
          errorMessage: error.message,
          stack: error.stack,
          service: 'drive',
          fileId,
        });
      } else {
        logger.error('❌ Помилка отримання метаданих файлу', {
          type: 'api_error',
          event: 'drive_metadata_failed',
          component: 'GoogleService',
          service: 'drive',
          fileId,
          errorMessage: String(error),
        });
      }
      throw error;
    }
  }

  /**
   * Отримання контенту документа
   */
  public async getDocumentContent(documentId: string): Promise<string> {
    try {
      const result = await this.executeWithRetry(async () => {
        if (!this.docs) throw new Error('Docs API не ініціалізовано');

        const response = await this.docs.documents.get({
          documentId,
        });

        // Парсинг контенту документа через узкий DocsService
        const content = this.getDocsService().extractTextFromDoc(response.data);
        return content;
      }, 'docs', undefined, 'docs.documents.get');

      return result;
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка отримання контенту документа', {
          type: 'api_error',
          event: 'docs_content_failed',
          component: 'GoogleService',
          errorName: error.name,
          errorMessage: error.message,
          stack: error.stack,
          service: 'docs',
          documentId,
        });
      } else {
        logger.error('❌ Помилка отримання контенту документа', {
          type: 'api_error',
          event: 'docs_content_failed',
          component: 'GoogleService',
          service: 'docs',
          documentId,
          errorMessage: String(error),
        });
      }
      throw error;
    }
  }

  

  /**
   * Отримання статистики з'єднань
   */
  public getConnectionStats(): Record<string, ConnectionInfo> {
    const stats: Record<string, ConnectionInfo> = {};
    for (const [apiType, connection] of this.connectionPool.entries()) {
      stats[apiType] = { ...connection };
    }
    return stats;
  }

  /**
   * Health check
   */
  protected async onHealthCheck(): Promise<HealthStatus> {
    try {
      // Перевірка автентифікації
      if (!this.auth) {
        return {
          healthy: false,
          service: this.name,
          error: 'Auth не ініціалізовано',
        };
      }

      // Перевірка API клієнтів
      if (!this.sheets || !this.drive || !this.docs) {
        return {
          healthy: false,
          service: this.name,
          error: 'API клієнти не ініціалізовано',
        };
      }

      // Тестовий запит до Sheets API
      try {
        await this.sheets.spreadsheets.get({
          spreadsheetId: this.config.google.spreadsheetId,
          ranges: ['A1:A1'],
        });
      } catch (error) {
        return {
          healthy: false,
          service: this.name,
          error: `Помилка тестового запиту: ${String(error)}`,
        };
      }

      return {
        healthy: true,
        service: this.name,
        details: {
          connectionPoolSize: this.connectionPool.size,
          requests: this.stats.requests,
          errors: this.stats.errors,
          averageResponseTime: this.stats.averageResponseTime,
        },
      };
    } catch (error) {
      return {
        healthy: false,
        service: this.name,
        error: `Health check failed: ${String(error)}`,
      };
    }
  }

  /**
   * Завершення роботи
   */
  protected async onShutdown(): Promise<void> {
    try {
      // Зупинка кеш сервісу
      await this.cacheService.shutdown();

      // Очищення connection pool
      this.connectionPool.clear();

      // Скидання API клієнтів
      this.sheets = null;
      this.drive = null;
      this.docs = null;
      this.auth = null;

      logger.info('✅ Google Service зупинено', {
        type: 'system',
        event: 'google_service_shutdown_success',
        component: 'GoogleService',
      });
    } catch (error) {
      if (error instanceof Error) {
        logger.error('❌ Помилка зупинки Google Service', {
          type: 'system',
          event: 'google_service_shutdown_failed',
          component: 'GoogleService',
          errorName: error.name,
          errorMessage: error.message,
          stack: error.stack,
        });
      } else {
        logger.error('❌ Помилка зупинки Google Service', {
          type: 'system',
          event: 'google_service_shutdown_failed',
          component: 'GoogleService',
          errorMessage: String(error),
        });
      }
      throw error;
    }
  }

  /**
   * Отримання статистики
   */
  protected onGetStats(): Partial<GoogleServiceStats> {
    return this.stats;
  }

  /**
   * Розбивка масиву на чанки
   */
  private chunkArray<T>(array: T[], chunkSize: number): T[][] {
    const chunks: T[][] = [];
    for (let i = 0; i < array.length; i += chunkSize) {
      chunks.push(array.slice(i, i + chunkSize));
    }
    return chunks;
  }

  /**
   * Оновлення статистики
   */
  private updateStats(success: boolean, duration: number): void {
    this.stats.requests++;
    if (!success) {
      this.stats.errors++;
    }

    // Оновлення середнього часу відповіді
    const totalTime = this.stats.averageResponseTime * (this.stats.requests - 1) + duration;
    this.stats.averageResponseTime = totalTime / this.stats.requests;

    // Оновлення використання connection pool
    let inUseConnections = 0;
    for (const connection of this.connectionPool.values()) {
      if (connection.inUse) {
        inUseConnections++;
      }
    }
    this.stats.connectionPoolUsage = (inUseConnections / this.connectionPool.size) * 100;
  }

  /**
   * Побудова індексу файлів у папці Drive та збереження у кеші
   */
  public async buildDriveIndex(
    folderId: string,
    opts: { ttlSeconds?: number; recursive?: boolean; maxDepth?: number } = {}
  ): Promise<drive_v3.Schema$File[]> {
    const { ttlSeconds = 3600, recursive = true, maxDepth = -1 } = opts;
    const files = await this.listDriveFilesInFolder(folderId, {
      recursive,
      type: 'any',
      limit: 10000,
      maxDepth,
    });
    const key = `drive:index:${folderId}`;
    try {
      await this.cacheService.set(key, files, ttlSeconds);
      logger.info('📇 Індекс Drive побудовано', {
        type: 'system',
        event: 'drive_index_built',
        component: 'GoogleService',
        folderId: folderId.substring(0, 10) + '...',
        count: files.length,
        ttl: `${ttlSeconds}s`,
      });
    } catch (e) {
      logger.warn('⚠️ Не вдалося зберегти індекс Drive у кеш', {
        type: 'system',
        event: 'drive_index_cache_failed',
        component: 'GoogleService',
        error: e instanceof Error ? e.message : String(e),
      });
    }
    return files;
  }

  /**
   * Прочитати індекс файлів з кешу
   */
  public async getDriveIndex(folderId: string): Promise<drive_v3.Schema$File[] | null> {
    const key = `drive:index:${folderId}`;
    try {
      const cached = await this.cacheService.get<drive_v3.Schema$File[]>(key);
      return cached ?? null;
    } catch (e) {
      logger.warn('⚠️ Не вдалося прочитати індекс Drive з кешу', {
        type: 'system',
        event: 'drive_index_cache_read_failed',
        component: 'GoogleService',
        error: e instanceof Error ? e.message : String(e),
      });
      return null;
    }
  }

  /**
   * Витягти текст з Google Docs документа
   */
  private async extractTextFromGoogleDoc(documentId: string): Promise<string> {
    const start = Date.now();
    if (!this.docs) throw new Error('Docs API не ініціалізовано');
    const res = await this.docs.documents.get({ documentId });
    const body = res.data.body;
    if (!body || !Array.isArray(body.content)) return '';
    const parts: string[] = [];
    for (const el of body.content) {
      if (!this.isParagraphElementContainer(el)) continue;
      const elements = el.paragraph.elements ?? [];
      for (const run of elements) {
        const text = run.textRun?.content;
        if (typeof text === 'string' && text.length > 0) parts.push(text);
      }
    }
    const text = parts.join('');
    try {
      this.metrics?.recordFileOperation({ operation: 'docs_get', status: 'success', mime: 'application/vnd.google-apps.document', fileId: documentId });
      this.metrics?.observeFileOperationLatency('docs_get', 'application/vnd.google-apps.document', Date.now() - start);
      this.metrics?.observeTextSizeBytes('parser', Buffer.byteLength(text, 'utf8'));
      this.metrics?.incrementMimeType('application/vnd.google-apps.document');
    } catch {}
    return text;
  }

  /** Type guard: элемент тела Google Docs, который содержит параграф с элементами */
  private isParagraphElementContainer(
    el: docs_v1.Schema$StructuralElement | undefined
  ): el is docs_v1.Schema$StructuralElement & { paragraph: { elements?: docs_v1.Schema$ParagraphElement[] } } {
    return Boolean(el && el.paragraph && Array.isArray(el.paragraph.elements));
  }

  /**
   * Отримати текстовий контент з файлу за його MIME типом (Docs/PDF/Word)
   */
  public async extractTextFromFile(file: drive_v3.Schema$File): Promise<string> {
    const mime = file.mimeType || '';
    try {
      if (!file.id) return '';
      // MIME counter
      try { if (mime) this.metrics?.incrementMimeType(mime); } catch {}
      if (mime === 'application/vnd.google-apps.document') {
        const started = Date.now();
        const t = await this.extractTextFromGoogleDoc(file.id);
        try {
          this.metrics?.recordFileOperation({ operation: 'extract_file_text', status: 'success', mime, fileId: file.id });
          this.metrics?.observeFileOperationLatency('extract_file_text', mime, Date.now() - started);
          this.metrics?.observeTextSizeBytes('parser', Buffer.byteLength(t, 'utf8'));
        } catch {}
        return t;
      }
      if (mime === 'application/pdf') {
        const started = Date.now();
        const t = await this.extractFromPdf(file.id);
        try {
          this.metrics?.recordFileOperation({ operation: 'extract_file_text', status: 'success', mime, fileId: file?.id ?? null });
          this.metrics?.observeFileOperationLatency('extract_file_text', mime, Date.now() - started);
          this.metrics?.observeTextSizeBytes('parser', Buffer.byteLength(t, 'utf8'));
        } catch {}
        return t;
      }
      if (this.isWordMime(mime)) {
        const started = Date.now();
        const t = await this.extractFromDocWord(file.id);
        try {
          this.metrics?.recordFileOperation({ operation: 'extract_file_text', status: 'success', mime, fileId: file?.id ?? null });
          this.metrics?.observeFileOperationLatency('extract_file_text', mime, Date.now() - started);
          this.metrics?.observeTextSizeBytes('parser', Buffer.byteLength(t, 'utf8'));
        } catch {}
        return t;
      }
      return '';
    } catch (error) {
      logger.error('❌ Помилка екстракції тексту з файлу', {
        type: 'processing_error',
        event: 'file_text_extract_failed',
        component: 'GoogleService',
        fileId: file.id,
        mime,
        error: error instanceof Error ? error.message : String(error),
      });
      try { this.metrics?.recordFileOperation({ operation: 'extract_file_text', status: 'error', mime, fileId: file?.id ?? null }); } catch {}
      return '';
    }
  }

  /**
   * MIME-предикат для Word-документов
   */
  private isWordMime(mime: string): boolean {
    return (
      mime === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
      mime === 'application/msword'
    );
  }

  /** Извлечение текста из PDF */
  private async extractFromPdf(fileId: string): Promise<string> {
    const dlStart = Date.now();
    const buf = await this.downloadFile(fileId);
    try {
      this.metrics?.recordFileOperation({ operation: 'download', status: 'success', mime: 'application/pdf', fileId });
      this.metrics?.observeFileOperationLatency('download', 'application/pdf', Date.now() - dlStart);
      this.metrics?.observeTextSizeBytes('download_bytes', buf.length);
    } catch {}
    const parseStart = Date.now();
    const { default: pdfParse } = await import('pdf-parse');
    const parsed = await pdfParse(buf);
    const text = parsed.text || '';
    try {
      this.metrics?.recordFileOperation({ operation: 'parse_pdf', status: 'success', mime: 'application/pdf', fileId });
      this.metrics?.observeFileOperationLatency('parse_pdf', 'application/pdf', Date.now() - parseStart);
      this.metrics?.observeTextSizeBytes('parser', Buffer.byteLength(text, 'utf8'));
    } catch {}
    return text;
  }

  /** Извлечение текста из DOC/DOCX */
  private async extractFromDocWord(fileId: string): Promise<string> {
    const dlStart = Date.now();
    const buf = await this.downloadFile(fileId);
    try {
      this.metrics?.recordFileOperation({ operation: 'download', status: 'success', mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', fileId });
      this.metrics?.observeFileOperationLatency('download', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', Date.now() - dlStart);
      this.metrics?.observeTextSizeBytes('download_bytes', buf.length);
    } catch {}
    const parseStart = Date.now();
    const mammoth = await import('mammoth');
    const result = await mammoth.extractRawText({ buffer: buf });
    const text = result.value || '';
    try {
      this.metrics?.recordFileOperation({ operation: 'parse_word', status: 'success', mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', fileId });
      this.metrics?.observeFileOperationLatency('parse_word', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', Date.now() - parseStart);
      this.metrics?.observeTextSizeBytes('parser', Buffer.byteLength(text, 'utf8'));
    } catch {}
    return text;
  }
}
