/**
 * Google Service з Connection Pool та оптимізацією
 * Покращена продуктивність та стабільність
 */

import { google } from 'googleapis';
import type { sheets_v4, drive_v3, docs_v1 } from 'googleapis';
import { ImageAnnotatorClient } from '@google-cloud/vision';
import type { Readable } from 'stream';
import { createHash } from 'crypto';
import pdfParse from 'pdf-parse';
import * as mammoth from 'mammoth';
import type { BotConfig, HealthStatus, ServiceStats, SheetData, BatchSheetData } from '@/types';
import type { DriveListQuery, DriveListResult, DriveFile } from '@/types/drive';
import { BaseService as BaseServiceClass } from '@/core/BaseService';
import { CacheService } from './CacheService';
import logger from '@/utils/logger';

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
  valueInputOption?: string;
  clearCache?: boolean;
}

export class GoogleService extends BaseServiceClass {
  private auth: any = null;
  private sheets: sheets_v4.Sheets | null = null;
  private drive: drive_v3.Drive | null = null;
  private docs: docs_v1.Docs | null = null;
  private connectionPool = new Map<string, ConnectionInfo>();
  private readonly retryAttempts = 3;
  private readonly retryDelay = 1000;
  private stats: GoogleServiceStats;
  private cacheService: CacheService;

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
  }

  /**
   * Нормализация Google Drive файла к внутреннему типу DriveFile
   */
  private toDriveFile(file: drive_v3.Schema$File): DriveFile {
    const ownersRaw = (file.owners || []) as Array<{ displayName?: string; emailAddress?: string }>;
    const owners: string[] = ownersRaw
      .map(o => o?.emailAddress || o?.displayName)
      .filter((v): v is string => typeof v === 'string' && v.length > 0);

    const df: DriveFile = {
      id: String(file.id || ''),
      name: String(file.name || ''),
      mimeType: String(file.mimeType || ''),
    };

    if (file.size) df.size = Number(file.size);
    if (file.modifiedTime) df.modifiedTime = file.modifiedTime;
    if (owners.length > 0) df.owners = owners;
    if (file.parents) df.parents = file.parents as string[];
    const webViewLink = (file as { webViewLink?: string }).webViewLink;
    if (webViewLink && !this.config.drive?.hideWebLink) df.webViewLink = webViewLink;
    const iconLink = (file as { iconLink?: string }).iconLink;
    if (iconLink) df.iconLink = iconLink;
    const isShortcut = file.mimeType === 'application/vnd.google-apps.shortcut';
    if (isShortcut) df.isShortcut = true;
    const sd = (file as { shortcutDetails?: { targetId?: string; targetMimeType?: string } }).shortcutDetails;
    if (sd?.targetId) {
      df.shortcutDetails = { targetId: String(sd.targetId) };
      if (sd.targetMimeType) df.shortcutDetails.targetMimeType = sd.targetMimeType;
    }

    return df;
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
      recursive = false,
      maxDepth = 20,
    } = query;

    const size = Math.max(5, Math.min(100, pageSize ?? this.config.drive.pageSize));
    const cacheKey = `drive:list:v2:${folderId}:${nameContains}:${(mimeIncludes || []).join('.')}:${(ownerAllowlist || []).join('.')}:${size}:${pageToken ?? ''}:${recursive}:${maxDepth}`;

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
    } catch {
      this.stats.cacheMisses++;
    }

    const allowedMimeCfg = this.config.drive.allowedMime || ['*'];
    const needMimeFilter = !(allowedMimeCfg.length === 1 && allowedMimeCfg[0] === '*');

    // Построение Drive query (q)
    const qParts: string[] = [
      `'${folderId}' in parents`,
      'trashed = false',
    ];
    if (nameContains) {
      // Простейший contains; Drive API чувствителен к регистру, но для начала достаточно
      const esc = nameContains.replace(/['\\]/g, '\\$&');
      qParts.push(`name contains '${esc}'`);
    }
    if (mimeIncludes && mimeIncludes.length > 0) {
      const ors = mimeIncludes.map(m => `mimeType='${m}'`).join(' or ');
      qParts.push(`(${ors})`);
    }
    if (needMimeFilter) {
      const ors = allowedMimeCfg.map(m => `mimeType='${m}'`).join(' or ');
      qParts.push(`(${ors})`);
    }

    const q = qParts.join(' and ');

    const fields = [
      'nextPageToken',
      "files(id,name,mimeType,size,modifiedTime,parents,webViewLink,iconLink,shortcutDetails,targetId,owners(displayName,emailAddress))",
    ].join(',');

    const start = Date.now();
    const res = await this.executeWithRetry(async () => {
      if (!this.drive) throw new Error('Drive API не инициализовано');
      const params: drive_v3.Params$Resource$Files$List = {
        q,
        pageSize: size,
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
        fields,
        corpora: 'allDrives',
      };
      if (pageToken) params.pageToken = pageToken;
      return this.drive.files.list(params);
    }, 'drive');

    const duration = Date.now() - start;
    const filesRaw = (res.data.files || []) as drive_v3.Schema$File[];
    let files = filesRaw.map(f => this.toDriveFile(f));

    // ownerAllowlist пост-фильтр, если задан
    if (ownerAllowlist && ownerAllowlist.length > 0) {
      const allow = new Set(ownerAllowlist.map(s => s.toLowerCase()));
      files = files.filter(f => (f.owners || []).some(o => allow.has(o.toLowerCase())));
    }

    const result: DriveListResult = { files };
    if (res.data.nextPageToken) result.nextPageToken = res.data.nextPageToken;

    // Кэшируем
    try {
      const ttl = this.config.drive.ttlListSec ?? 60;
      await this.cacheService.set(cacheKey, result, ttl);
    } catch {}

    logger.info('📄 Листинг Drive завершён', {
      type: 'api_request',
      event: 'drive_list_success',
      component: 'GoogleService',
      folderId,
      count: result.files.length,
      duration,
      pageToken: pageToken ? 'yes' : 'no',
      nextPageToken: result.nextPageToken ? 'yes' : 'no',
    });

    // Рекурсивная ветка (упрощённо): не разворачиваем здесь глубину, оставим на будущий индексатор
    // Можно реализовать отдельным сервисом, чтобы не блокировать текущий шаг.

    return result;
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
      // Типовой контракт: alt='media' возвращает поток
      const r = await this.drive.files.get({ fileId, alt: 'media' as unknown as string } as drive_v3.Params$Resource$Files$Get);
      return r as unknown as { data: Readable };
    }, 'drive');

    const buf = await this.streamToBuffer(resp.data as unknown as NodeJS.ReadableStream);
    try {
      const ttl = this.config.drive.ttlTextSec ?? 300;
      await this.cacheService.set(cacheKey, buf, ttl);
    } catch {}

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
      const r = await this.drive.files.export({ fileId, mimeType: mimeOut } as drive_v3.Params$Resource$Files$Export);
      return r as unknown as { data: Readable };
    }, 'drive');

    const buf = await this.streamToBuffer(resp.data as unknown as NodeJS.ReadableStream);
    try {
      const ttl = this.config.drive.ttlTextSec ?? 300;
      await this.cacheService.set(cacheKey, buf, ttl);
    } catch {}

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
   * Безопасно собрать поток в Buffer
   */
  private async streamToBuffer(stream: NodeJS.ReadableStream): Promise<Buffer> {
    return new Promise<Buffer>((resolve, reject) => {
      const chunks: Buffer[] = [];
      stream.on('data', (d: unknown) => {
        if (Buffer.isBuffer(d)) chunks.push(d);
        else if (typeof d === 'string') chunks.push(Buffer.from(d));
        else chunks.push(Buffer.from(String(d)));
      });
      stream.on('end', () => resolve(Buffer.concat(chunks)));
      stream.on('error', reject);
    });
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
      const buf = await this.downloadDriveFile(file.id!);

      // Ключ кэша на основе fileId + modifiedTime (если есть) + хэш контента
      const modified = (file as any).modifiedTime || '';
      const hash = createHash('sha1').update(buf).digest('hex');
      const cacheKey = `ocr:image:${file.id}:${modified}:${hash}`;

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
    const client = new ImageAnnotatorClient();
    const [result] = await client.textDetection({ image: { content: buf } });
    const annotations = result?.textAnnotations ?? [];
    const first = annotations[0];
    if (first && (first as any).description) return String((first as any).description);
    const descriptions = annotations
      .map(a => (a && 'description' in a ? ((a as any).description as string | undefined) : undefined))
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
      const { createWorker } = await import('tesseract.js');

      const langs = this.config.google.tesseractLangs || 'eng';
      const langPath = this.config.google.tesseractLangPath;

      const worker = await createWorker({
        // logger: m => logger.debug('tesseract', { progress: m.progress }),
        langPath, // если задан, tesseract.js загрузит traineddata локально
        cachePath: undefined,
      } as any);

      try {
        await (worker as any).loadLanguage(langs);
        await (worker as any).initialize(langs);

        const { data } = (await (worker as any).recognize(buf)) as any;
        const text: string = data?.text ?? '';
        return text;
      } finally {
        await (worker as any).terminate();
      }
    } catch (err) {
      logger.error('❌ Помилка Tesseract OCR', {
        type: 'processing_error',
        event: 'tesseract_ocr_failed',
        component: 'GoogleService',
        error: err instanceof Error ? err.message : String(err),
      });
      // Фоллбек на Vision, если доступно
      try {
        return await this.ocrWithVision(buf);
      } catch {
        return '';
      }
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
      }, 'drive');
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
      }, 'drive');
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
      }, 'drive');
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
      await this.initializeAPIs();

      // Створення connection pool
      await this.initializeConnectionPool();

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

    // Побудова MIME фільтрів
    // filesMimeFilter — для файлів поточного типу
    // foldersMimeFilter — для визначення підпапок (включаючи ярлики на папки)
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

    // Побудова name contains фільтра
    const nameFilter = query ? ` and name contains '${query.replace(/'/g, "\\'")}'` : '';

    // Підготовка запитів: окремо для файлів потрібного типу та для підпапок
    const qFiles = `'${folderId}' in parents and trashed=false${filesMimeFilter}${nameFilter}`;
    const qFolders = `'${folderId}' in parents and trashed=false${foldersMimeFilter}`;

    // Пагінація: вибірка усіх сторінок
    const fetchAll = async (q: string): Promise<drive_v3.Schema$File[]> => {
      const all: drive_v3.Schema$File[] = [];
      let token: string | undefined = pageToken;
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
        }, 'drive');
        const files = resp.data.files || [];
        all.push(...files);
        token = resp.data.nextPageToken || undefined;
      } while (token && all.length < limit);
      return all;
    };

    const filesHere = await fetchAll(qFiles);

    // Якщо шукаємо таблиці — конвертуємо ярлики на таблиці у "віртуальні" записи з targetId та
    // додаємо Excel-файли у результат
    let results: drive_v3.Schema$File[] =
      type === 'sheet'
        ? filesHere
            .map(f => {
              if (f.mimeType === 'application/vnd.google-apps.shortcut') {
                const targetId = (f as any).shortcutDetails?.targetId as string | undefined;
                const targetMime = (f as any).shortcutDetails?.targetMimeType as string | undefined;
                if (targetId && targetMime === 'application/vnd.google-apps.spreadsheet') {
                  // Повертаємо об'єкт як ніби це сам spreadsheet
                  return {
                    id: targetId,
                    name: f.name,
                    mimeType: 'application/vnd.google-apps.spreadsheet',
                    parents: f.parents,
                  } as drive_v3.Schema$File;
                }
                return null;
              }
              // Додаємо Google Таблиці та Excel-файли
              const isGs = f.mimeType === 'application/vnd.google-apps.spreadsheet';
              const isXlsx =
                f.mimeType ===
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
              const isXls = f.mimeType === 'application/vnd.ms-excel';
              return isGs || isXlsx || isXls ? f : null;
            })
            .filter((x): x is drive_v3.Schema$File => Boolean(x))
        : [...filesHere];

    if (recursive && (maxDepth > 0 || maxDepth <= -1)) {
      // Отримуємо підпапки в поточній папці (включаючи ярлики на папки)
      const foldersLevel = await fetchAll(qFolders);
      const folders = foldersLevel
        .map<{ id: string; name: string | undefined } | null>(f => {
          if (f.mimeType === 'application/vnd.google-apps.folder') {
            return { id: f.id!, name: f.name ?? undefined };
          }
          if (
            f.mimeType === 'application/vnd.google-apps.shortcut' &&
            (f as any).shortcutDetails?.targetMimeType === 'application/vnd.google-apps.folder'
          ) {
            const targetId = (f as any).shortcutDetails?.targetId as string | undefined;
            if (targetId) return { id: targetId, name: f.name ?? undefined };
          }
          return null;
        })
        .filter((x): x is { id: string; name: string | undefined } => x !== null);

      for (const folder of folders) {
        try {
          const sub = await this.listDriveFilesInFolder(folder.id!, {
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
    } catch {}

    return results;
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
    }, 'sheets');

    try {
      await this.cacheService.set(cacheKey, titles, 60);
    } catch {}
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
      this.auth = new google.auth.JWT(
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
      await this.auth.authorize();
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
  private async initializeAPIs(): Promise<void> {
    try {
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
  private async initializeConnectionPool(): Promise<void> {
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
    maxRetries: number = this.retryAttempts
  ): Promise<T> {
    let lastError: Error | null = null;

    for (let attempt = 0; attempt <= maxRetries; attempt++) {
      try {
        const connection = this.getConnection(apiType);
        if (!connection) {
          throw new Error(`Немає доступних з'єднань для ${apiType}`);
        }

        const startTime = Date.now();
        const result = await operation();
        const duration = Date.now() - startTime;

        this.releaseConnection(apiType);
        this.updateStats(true, duration);

        return result;
      } catch (error) {
        lastError = error as Error;
        this.releaseConnection(apiType);
        this.updateStats(false, 0);

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
      // Перевірка кешу
      if (useCache && !forceRefresh) {
        const cacheKey = `sheets:${spreadsheetId}:${range}`;
        try {
          const cached = await this.cacheService.get<SheetData>(cacheKey);
          if (cached) {
            this.stats.cacheHits++;
            logger.debug('✅ Використано кешовані дані Sheets', {
              type: 'system',
              event: 'cache_hit',
              component: 'GoogleService',
              spreadsheetId: spreadsheetId.substring(0, 10) + '...',
              range,
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
          range,
        });

        return {
          range: response.data.range || range,
          majorDimension: response.data.majorDimension || 'ROWS',
          values: response.data.values || [],
        };
      }, 'sheets');

      // Збереження в кеш
      if (useCache) {
        const cacheKey = `sheets:${spreadsheetId}:${range}`;
        try {
          // CacheService expects TTL in seconds; do not convert to ms
          await this.cacheService.set(cacheKey, result, cacheTTL);
          logger.debug('💾 Дані Sheets збережено в кеш', {
            type: 'system',
            event: 'cache_write',
            component: 'GoogleService',
            spreadsheetId: spreadsheetId.substring(0, 10) + '...',
            range,
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
      await this.executeWithRetry(async () => {
        if (!this.sheets) throw new Error('Sheets API не ініціалізовано');

        await this.sheets.spreadsheets.values.update({
          spreadsheetId,
          range,
          valueInputOption,
          requestBody: {
            values,
          },
        });
      }, 'sheets');

      // Очищення кешу
      if (clearCache) {
        const cacheKey = `sheets:${spreadsheetId}:${range}`;
        try {
          await this.cacheService.delete(cacheKey);
          logger.debug('🗑️ Кеш Sheets очищено', {
            type: 'system',
            event: 'cache_delete',
            component: 'GoogleService',
            spreadsheetId: spreadsheetId.substring(0, 10) + '...',
            range,
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
      const chunks = this.chunkArray(ranges, batchSize);
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

              return response.data.valueRanges || [];
            },
            'sheets',
            maxRetries
          );

          results.push(
            ...result.map(vr => ({
              range: vr.range ?? '',
              majorDimension: vr.majorDimension ?? 'ROWS',
              values: vr.values ?? [],
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
    const { batchSize = 10, retryFailed = true, maxRetries = 3, clearCache = true } = options;

    try {
      const chunks = this.chunkArray(data, batchSize);
      const failedBatches: Array<{ range: string; values: string[][] }> = [];

      for (const chunk of chunks) {
        try {
          await this.executeWithRetry(
            async () => {
              if (!this.sheets) throw new Error('Sheets API не ініціалізовано');

              // Використовуємо values.batchUpdate, щоб не залежати від sheetId
              const valueUpdates = chunk.map(item => ({
                range: item.range,
                values: item.values,
              }));

              await this.sheets.spreadsheets.values.batchUpdate({
                spreadsheetId,
                requestBody: {
                  data: valueUpdates,
                  valueInputOption: 'RAW',
                },
              });
            },
            'sheets',
            maxRetries
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
      }, 'drive');

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
      }, 'drive');

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

        // Парсинг контенту документа
        const content = this.parseDocumentContent(response.data);
        return content;
      }, 'docs');

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
   * Парсинг контенту документа
   */
  private parseDocumentContent(document: docs_v1.Schema$Document): string {
    if (!document.body?.content) {
      return '';
    }

    let content = '';
    for (const element of document.body.content) {
      if (element.paragraph) {
        for (const element2 of element.paragraph.elements || []) {
          if (element2.textRun?.content) {
            content += element2.textRun.content;
          }
        }
        content += '\n';
      }
    }

    return content.trim();
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
    if (!this.docs) throw new Error('Docs API не ініціалізовано');
    const res = await this.docs.documents.get({ documentId });
    const body = res.data.body;
    if (!body || !body.content) return '';
    const parts: string[] = [];
    for (const el of body.content) {
      const p = (el as any).paragraph;
      if (!p || !p.elements) continue;
      for (const run of p.elements) {
        const tr = (run as any).textRun;
        if (tr && tr.content) parts.push(tr.content);
      }
    }
    return parts.join('');
  }

  /**
   * Отримати текстовий контент з файлу за його MIME типом (Docs/PDF/Word)
   */
  public async extractTextFromFile(file: drive_v3.Schema$File): Promise<string> {
    const mime = file.mimeType || '';
    try {
      if (mime === 'application/vnd.google-apps.document') {
        return await this.extractTextFromGoogleDoc(file.id!);
      }
      if (mime === 'application/pdf') {
        const buf = await this.downloadDriveFile(file.id!);
        const parsed = await pdfParse(buf as unknown as Buffer);
        return parsed.text || '';
      }
      if (
        mime === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
        mime === 'application/msword'
      ) {
        const buf = await this.downloadDriveFile(file.id!);
        const result = await mammoth.extractRawText({ buffer: buf });
        return result.value || '';
      }
      return '';
    } catch (error) {
      logger.error('❌ Помилка екстракції тексту з файлу', {
        type: 'processing_error',
        event: 'file_text_extract_failed',
        component: 'GoogleService',
        fileId: file.id,
        mimeType: mime,
        error: error instanceof Error ? error.message : String(error),
      });
      return '';
    }
  }
}
