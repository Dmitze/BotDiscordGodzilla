import { google } from 'googleapis';
import type { BotConfig } from '@/types';
import type { MetricsService } from './MetricsService';
import { DocsService } from './google/DocsService';
import { CacheService } from './CacheService';
import logger from '@/utils/logger';
import { createHash } from 'crypto';
import { chunkTextByTokens } from '@/utils/textChunker';
import type { SearchIndex } from '@/search/SearchIndex';
import type { EmbeddingsProvider } from './EmbeddingsService';

/**
 * GoogleDocsService - Сервіс для роботи з Google Docs
 * Реалізує методи listDocs, getDocContent, indexDoc, searchDoc, summarizeDoc
 * Інтегрується з існуючим GoogleService для автентифікації
 */
export class GoogleDocsService {
  private docsService: DocsService;
  private cacheService: CacheService;
  private metrics?: MetricsService;
  private searchIndex?: SearchIndex;

  constructor(
    private readonly config: BotConfig,
    private readonly googleAuth: InstanceType<typeof google.auth.JWT>,
    metrics?: MetricsService
  ) {
    this.docsService = new DocsService(metrics);
    this.cacheService = new CacheService(config);
    this.metrics = metrics;
  }

  /**
   * Встановити сервіс індексації
   * @param searchIndex Сервіс індексації
   */
  public setSearchIndex(searchIndex: SearchIndex): void {
    this.searchIndex = searchIndex;
  }

  /**
   * Встановити сервіс ембеддінгів
   * @param embeddingsService Сервіс ембеддінгів
   */
  public setEmbeddingsService(embeddingsService: EmbeddingsProvider): void {
    this.embeddingsService = embeddingsService;
  }

  /**
   * Отримати список доступних Google Docs документів
   * @param folderId - ID папки Google Drive для пошуку документів (опціонально)
   * @param query - Пошуковий запит для фільтрації документів (опціонально)
   * @returns Масив об'єктів з інформацією про документи
   */
  public async listDocs(folderId?: string, query?: string): Promise<Array<{
    id: string;
    name: string;
    mimeType: string;
    modifiedTime?: string;
    owners?: Array<{ displayName?: string; emailAddress?: string }>;
  }>> {
    const startTime = Date.now();
    try {
      // Ключ кешу для списку документів
      const cacheKey = `docs:list:${folderId || 'root'}:${query || 'all'}`;
      
      // Спроба отримати з кешу
      try {
        const cached = await this.cacheService.get(cacheKey);
        if (cached) {
          logger.debug('✅ Отримано список документів з кешу', {
            type: 'cache',
            event: 'docs_list_cache_hit',
            component: 'GoogleDocsService',
            folderId,
            query,
          });
          return cached as any;
        }
      } catch (error) {
        logger.debug('⚠️ Не вдалося отримати список документів з кешу', {
          type: 'cache',
          event: 'docs_list_cache_miss',
          component: 'GoogleDocsService',
          folderId,
          query,
          error: error instanceof Error ? error.message : String(error),
        });
      }

      // Створення клієнта Docs API
      const docs = google.docs({
        version: 'v1',
        auth: this.googleAuth,
      });

      // Побудова запиту для пошуку Google Docs
      const qParts: string[] = [
        'mimeType=\'application/vnd.google-apps.document\'',
        'trashed = false',
      ];

      if (folderId) {
        qParts.push(`'${folderId}' in parents`);
      }

      if (query) {
        const escapedQuery = query.replace(/['\\]/g, '\\$&');
        qParts.push(`name contains '${escapedQuery}'`);
      }

      const q = qParts.join(' and ');

      // Виконання запиту до Drive API для отримання списку документів
      const drive = google.drive({
        version: 'v3',
        auth: this.googleAuth,
      });

      const response = await drive.files.list({
        q,
        pageSize: 100,
        fields: 'files(id,name,mimeType,modifiedTime,owners(displayName,emailAddress))',
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
        corpora: 'allDrives',
      });

      const files = response.data.files || [];
      const docsList = files.map(file => ({
        id: file.id || '',
        name: file.name || '',
        mimeType: file.mimeType || '',
        modifiedTime: file.modifiedTime ?? undefined,
        owners: file.owners,
      }));

      // Кешування результату
      try {
        const ttl = this.config.drive?.ttlListSec ?? 300; // 5 хвилин за замовчуванням
        await this.cacheService.set(cacheKey, docsList, ttl);
      } catch (error) {
        logger.warn('⚠️ Не вдалося закешувати список документів', {
          type: 'cache',
          event: 'docs_list_cache_set_failed',
          component: 'GoogleDocsService',
          folderId,
          query,
          error: error instanceof Error ? error.message : String(error),
        });
      }

      // Метрики
      try {
        this.metrics?.updateGoogleApiMetrics('docs', 'list', 'ok', Date.now() - startTime);
      } catch (error) {
        logger.debug('⚠️ Не вдалося записати метрики списку документів', {
          type: 'metrics',
          event: 'docs_list_metrics_failed',
          component: 'GoogleDocsService',
          error: error instanceof Error ? error.message : String(error),
        });
      }

      logger.info('📄 Отримано список Google Docs документів', {
        type: 'api_request',
        event: 'docs_list_success',
        component: 'GoogleDocsService',
        folderId,
        query,
        count: docsList.length,
        duration: Date.now() - startTime,
      });

      return docsList;
    } catch (error) {
      // Метрики помилок
      try {
        this.metrics?.updateGoogleApiMetrics('docs', 'list', 'error', Date.now() - startTime);
      } catch { /* noop */ }

      logger.error('❌ Помилка отримання списку Google Docs документів', {
        type: 'api_error',
        event: 'docs_list_failed',
        component: 'GoogleDocsService',
        folderId,
        query,
        error: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });

      throw error;
    }
  }

  /**
   * Отримати вміст Google Docs документа
   * @param documentId - ID документа Google Docs
   * @returns Об'єкт з вмістом документа
   */
  public async getDocContent(documentId: string): Promise<{
    title: string;
    content: string;
    blocks: any[];
    modifiedTime?: string;
  }> {
    const startTime = Date.now();
    try {
      // Ключ кешу для вмісту документа
      const cacheKey = `docs:content:${documentId}`;
      
      // Спроба отримати з кешу
      try {
        const cached = await this.cacheService.get(cacheKey);
        if (cached) {
          logger.debug('✅ Отримано вміст документа з кешу', {
            type: 'cache',
            event: 'docs_content_cache_hit',
            component: 'GoogleDocsService',
            documentId,
          });
          return cached as any;
        }
      } catch (error) {
        logger.debug('⚠️ Не вдалося отримати вміст документа з кешу', {
          type: 'cache',
          event: 'docs_content_cache_miss',
          component: 'GoogleDocsService',
          documentId,
          error: error instanceof Error ? error.message : String(error),
        });
      }

      // Створення клієнта Docs API
      const docs = google.docs({
        version: 'v1',
        auth: this.googleAuth,
      });

      // Отримання документа
      const response = await docs.documents.get({
        documentId,
        fields: 'title,body,documentStyle,headers,footers,footnotes,lists,tables,revisions',
      });

      const document = response.data;
      const title = document.title || '';
      
      // Отримання текстового вмісту
      const content = this.docsService.extractTextFromDoc(document);
      
      // Отримання структурованих блоків
      const blocks = this.docsService.extractBlocksFromDoc(document);
      
      const result = {
        title,
        content,
        blocks,
        modifiedTime: document.suggestionsViewMode as string || undefined, // This is a workaround, actual modifiedTime should come from Drive metadata
      };

      // Кешування результату
      try {
        const ttl = this.config.drive?.ttlTextSec ?? 300; // 5 хвилин за замовчуванням
        await this.cacheService.set(cacheKey, result, ttl);
      } catch (error) {
        logger.warn('⚠️ Не вдалося закешувати вміст документа', {
          type: 'cache',
          event: 'docs_content_cache_set_failed',
          component: 'GoogleDocsService',
          documentId,
          error: error instanceof Error ? error.message : String(error),
        });
      }

      // Метрики
      try {
        this.metrics?.updateGoogleApiMetrics('docs', 'get', 'ok', Date.now() - startTime);
      } catch (error) {
        logger.debug('⚠️ Не вдалося записати метрики вмісту документа', {
          type: 'metrics',
          event: 'docs_content_metrics_failed',
          component: 'GoogleDocsService',
          error: error instanceof Error ? error.message : String(error),
        });
      }

      logger.info('📄 Отримано вміст Google Docs документа', {
        type: 'api_request',
        event: 'docs_content_success',
        component: 'GoogleDocsService',
        documentId,
        title,
        contentLength: content.length,
        blocksCount: blocks.length,
        duration: Date.now() - startTime,
      });

      return result;
    } catch (error) {
      // Метрики помилок
      try {
        this.metrics?.updateGoogleApiMetrics('docs', 'get', 'error', Date.now() - startTime);
      } catch { /* noop */ }

      logger.error('❌ Помилка отримання вмісту Google Docs документа', {
        type: 'api_error',
        event: 'docs_content_failed',
        component: 'GoogleDocsService',
        documentId,
        error: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });

      throw error;
    }
  }

  /**
   * Індексувати Google Docs документ для пошуку
   * @param documentId - ID документа Google Docs
   * @returns Результат індексації
   */
  public async indexDoc(documentId: string): Promise<{
    success: boolean;
    documentId: string;
    indexedAt: string;
    contentHash: string;
    wordCount: number;
  }> {
    const startTime = Date.now();
    try {
      logger.info('🔍 Початок індексації Google Docs документа', {
        type: 'indexing',
        event: 'docs_index_start',
        component: 'GoogleDocsService',
        documentId,
      });

      // Отримання вмісту документа
      const docContent = await this.getDocContent(documentId);
      
      // Генерація хешу вмісту
      const contentHash = createHash('sha256').update(docContent.content).digest('hex');
      
      // Підрахунок слів
      const wordCount = docContent.content.trim().split(/\s+/).filter(Boolean).length;
      
      // Інтеграція з існуючою системою індексації (SqliteSearchIndex)
      if (this.searchIndex) {
        // Розбиття тексту на частини згідно з планом (800-1200 токенів з перекриттям 100 токенів)
        const chunks = chunkTextByTokens(
          docContent.content,
          1000, // target tokens
          800,  // min tokens
          1200, // max tokens
          100   // overlap tokens
        );
        
        logger.info('📄 Розбито документ на частини', {
          type: 'indexing',
          event: 'docs_chunking_complete',
          component: 'GoogleDocsService',
          documentId,
          chunkCount: chunks.length,
        });
        
        // Індексація кожної частини
        for (let i = 0; i < chunks.length; i++) {
          const chunk = chunks[i];
          
          // Створення унікального ID для частини
          const chunkId = `${documentId}_chunk_${i}`;
          
          // Підготовка метаданих для індексації
          const docToIndex = {
            fileId: chunkId,
            name: `${docContent.title} (частина ${i + 1})`,
            mimeType: 'application/vnd.google-apps.document.chunk',
            text: chunk?.text ?? '',
            tags: ['google-docs', 'chunk'],
            meta: {
              originalDocumentId: documentId,
              originalDocumentName: docContent.title,
              chunkIndex: i,
              chunkStart: chunk?.start ?? 0,
              chunkEnd: chunk?.end ?? 0,
              chunkTokenCount: chunk?.tokenCount ?? 0,
            },
            language: 'uk', // Оскільки це український бот
          };
          
          // Індексація частини
          await this.searchIndex.upsert(docToIndex);
        }
        
        // Також індексуємо весь документ як єдине ціле для глобального пошуку
        const fullDocToIndex = {
          fileId: documentId,
          name: docContent.title,
          mimeType: 'application/vnd.google-apps.document',
          text: docContent.content,
          tags: ['google-docs', 'full-document'],
          meta: {
            wordCount: wordCount,
            chunkCount: chunks.length,
          },
          language: 'uk',
        };
        
        await this.searchIndex.upsert(fullDocToIndex);
        
        logger.info('✅ Успішно проіндексовано Google Docs документ та його частини', {
          type: 'indexing',
          event: 'docs_index_success',
          component: 'GoogleDocsService',
          documentId,
          contentHash,
          wordCount,
          chunkCount: chunks.length,
          duration: Date.now() - startTime,
        });
      } else {
        logger.warn('⚠️ Сервіс індексації не доступний, індексація пропущена', {
          type: 'indexing',
          event: 'docs_index_skipped',
          component: 'GoogleDocsService',
          documentId,
        });
      }

      const result = {
        success: true,
        documentId,
        indexedAt: new Date().toISOString(),
        contentHash,
        wordCount,
      };

      // Метрики
      try {
        this.metrics?.updateGoogleApiMetrics('docs', 'index', 'ok', Date.now() - startTime);
      } catch (error) {
        logger.debug('⚠️ Не вдалося записати метрики індексації документа', {
          type: 'metrics',
          event: 'docs_index_metrics_failed',
          component: 'GoogleDocsService',
          error: error instanceof Error ? error.message : String(error),
        });
      }

      return result;
    } catch (error) {
      // Метрики помилок
      try {
        this.metrics?.updateGoogleApiMetrics('docs', 'index', 'error', Date.now() - startTime);
      } catch { /* noop */ }

      logger.error('❌ Помилка індексації Google Docs документа', {
        type: 'indexing_error',
        event: 'docs_index_failed',
        component: 'GoogleDocsService',
        documentId,
        error: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });

      return {
        success: false,
        documentId,
        indexedAt: new Date().toISOString(),
        contentHash: '',
        wordCount: 0,
      };
    }
  }

  /**
   * Пошук у Google Docs документі
   * @param documentId - ID документа Google Docs
   * @param query - Пошуковий запит
   * @returns Результати пошуку
   */
  public async searchDoc(documentId: string, query: string): Promise<Array<{
    blockIndex: number;
    blockType: string;
    content: string;
    matchPosition: number;
    relevanceScore: number;
  }>> {
    const startTime = Date.now();
    try {
      logger.info('🔍 Початок пошуку в Google Docs документі', {
        type: 'search',
        event: 'docs_search_start',
        component: 'GoogleDocsService',
        documentId,
        query,
      });

      // Отримання структурованого вмісту документа
      const docContent = await this.getDocContent(documentId);
      
      const results: Array<{
        blockIndex: number;
        blockType: string;
        content: string;
        matchPosition: number;
        relevanceScore: number;
      }> = [];
      
      const lowerQuery = query.toLowerCase();
      
      // Пошук у блоках
      docContent.blocks.forEach((block, index) => {
        let content = '';
        let blockType = 'unknown';
        
        if (block.kind === 'paragraph') {
          content = block.text;
          blockType = 'paragraph';
        } else if (block.kind === 'heading') {
          content = block.text;
          blockType = `heading-${block.level}`;
        } else if (block.kind === 'listItem') {
          content = block.text;
          blockType = 'list-item';
        } else if (block.kind === 'table') {
          // Для таблиць об'єднуємо вміст всіх комірок
          content = block.rows.map((row: any) =>
            row.cells.map((cell: any) => cell.text).join(' ')
          ).join(' ');
          blockType = 'table';
        } else if (block.kind === 'footnote') {
          content = block.text;
          blockType = 'footnote';
        }
        
        const lowerContent = content.toLowerCase();
        const matchPosition = lowerContent.indexOf(lowerQuery);
        
        if (matchPosition !== -1) {
          // Простий розрахунок релевантності на основі позиції та довжини
          const relevanceScore = 1 / (1 + matchPosition / content.length);
          
          results.push({
            blockIndex: index,
            blockType,
            content,
            matchPosition,
            relevanceScore,
          });
        }
      });
      
      // Сортування за релевантністю
      results.sort((a, b) => b.relevanceScore - a.relevanceScore);

      // Метрики
      try {
        this.metrics?.updateGoogleApiMetrics('docs', 'search', 'ok', Date.now() - startTime);
      } catch (error) {
        logger.debug('⚠️ Не вдалося записати метрики пошуку в документі', {
          type: 'metrics',
          event: 'docs_search_metrics_failed',
          component: 'GoogleDocsService',
          error: error instanceof Error ? error.message : String(error),
        });
      }

      logger.info('✅ Успішно виконано пошук в Google Docs документі', {
        type: 'search',
        event: 'docs_search_success',
        component: 'GoogleDocsService',
        documentId,
        query,
        resultsCount: results.length,
        duration: Date.now() - startTime,
      });

      return results;
    } catch (error) {
      // Метрики помилок
      try {
        this.metrics?.updateGoogleApiMetrics('docs', 'search', 'error', Date.now() - startTime);
      } catch { /* noop */ }

      logger.error('❌ Помилка пошуку в Google Docs документі', {
        type: 'search_error',
        event: 'docs_search_failed',
        component: 'GoogleDocsService',
        documentId,
        query,
        error: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });

      throw error;
    }
  }

  /**
   * Згенерувати резюме Google Docs документа
   * @param documentId - ID документа Google Docs
   * @returns Резюме документа
   */
  public async summarizeDoc(documentId: string): Promise<{
    title: string;
    summary: string;
    keyPoints: string[];
    wordCount: number;
    readingTimeMinutes: number;
  }> {
    const startTime = Date.now();
    try {
      logger.info('📝 Початок генерації резюме Google Docs документа', {
        type: 'summarization',
        event: 'docs_summarize_start',
        component: 'GoogleDocsService',
        documentId,
      });

      // Отримання вмісту документа
      const docContent = await this.getDocContent(documentId);
      
      // Підрахунок слів
      const wordCount = docContent.content.trim().split(/\s+/).filter(Boolean).length;
      
      // Оцінка часу читання (приблизно 200 слів на хвилину)
      const readingTimeMinutes = Math.ceil(wordCount / 200);
      
      // Вилучення ключових точок з заголовків
      const keyPoints: string[] = [];
      docContent.blocks.forEach(block => {
        if (block.kind === 'heading' && block.level && block.level <= 3) {
          keyPoints.push(block.text);
        }
      });
      
      // Просте резюме на основі першого абзацу та ключових точок
      let summary = '';
      const firstParagraph = docContent.blocks.find(block => block.kind === 'paragraph');
      if (firstParagraph) {
        // Обмежуємо перший абзац 200 словами
        const words = firstParagraph.text.split(/\s+/);
        summary = words.slice(0, 200).join(' ');
        if (words.length > 200) {
          summary += '...';
        }
      }

      const result = {
        title: docContent.title,
        summary,
        keyPoints,
        wordCount,
        readingTimeMinutes,
      };

      // Метрики
      try {
        this.metrics?.updateGoogleApiMetrics('docs', 'summarize', 'ok', Date.now() - startTime);
      } catch (error) {
        logger.debug('⚠️ Не вдалося записати метрики резюмування документа', {
          type: 'metrics',
          event: 'docs_summarize_metrics_failed',
          component: 'GoogleDocsService',
          error: error instanceof Error ? error.message : String(error),
        });
      }

      logger.info('✅ Успішно згенеровано резюме Google Docs документа', {
        type: 'summarization',
        event: 'docs_summarize_success',
        component: 'GoogleDocsService',
        documentId,
        title: docContent.title,
        wordCount,
        keyPointsCount: keyPoints.length,
        duration: Date.now() - startTime,
      });

      return result;
    } catch (error) {
      // Метрики помилок
      try {
        this.metrics?.updateGoogleApiMetrics('docs', 'summarize', 'error', Date.now() - startTime);
      } catch { /* noop */ }

      logger.error('❌ Помилка генерації резюме Google Docs документа', {
        type: 'summarization_error',
        event: 'docs_summarize_failed',
        component: 'GoogleDocsService',
        documentId,
        error: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });

      throw error;
    }
  }
}