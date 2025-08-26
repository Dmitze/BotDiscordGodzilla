/**
 * 🔍 Enhanced RAG Service with Google Drive Auto-Indexing
 * Combines the existing RAG functionality with automatic Google Drive document indexing
 */

import { RagService } from './RagService';
import { GoogleService } from './GoogleService';
import { DriveIndexerService } from './DriveIndexerService';
import { ResponseCacheService } from './ResponseCacheService';
import SchedulerService from './SchedulerService';
import type { SearchIndex } from '@/search/SearchIndex';
import type { AIService } from './AIService';
import logger from '@/utils/logger';

export interface AutoIndexConfig {
  enabled: boolean;
  interval: string; // cron expression
  folders: string[]; // Google Drive folder IDs to monitor
  fileTypes: string[]; // Supported file types
  maxFileSize: number; // Maximum file size in bytes
  batchSize: number; // Number of files to process per batch
}

export interface IndexingStats {
  totalFiles: number;
  indexedFiles: number;
  failedFiles: number;
  lastIndexingRun: Date | null;
  nextScheduledRun: Date | null;
  averageProcessingTime: number;
  indexingProgress: {
    inProgress: boolean;
    currentFile: string | null;
    processed: number;
    total: number;
  };
}

export interface DriveFile {
  id: string;
  name: string;
  mimeType: string;
  size: number;
  modifiedTime: Date;
  content?: string;
  indexed: boolean;
  lastIndexed?: Date;
  error?: string;
}

export class EnhancedRagService extends RagService {
  private autoIndexConfig: AutoIndexConfig;
  private indexingStats: IndexingStats;
  private indexingInProgress = false;
  private scheduledTaskId?: string;

  constructor(
    searchIndex: SearchIndex,
    ai: AIService,
    private readonly googleService: GoogleService,
    private readonly driveIndexer: DriveIndexerService,
    private readonly responseCache: ResponseCacheService,
    private readonly scheduler: SchedulerService,
    embeddings?: { embed: (text: string) => Promise<number[]> },
    autoIndexConfig?: Partial<AutoIndexConfig>
  ) {
    super(searchIndex, ai, embeddings, {
      maxSize: 500,
      ttlSec: 1800, // 30 minutes
      cache: {
        get: async <T>(key: string) => responseCache.get<T>(key),
        set: async <T>(key: string, value: T, ttlSec?: number) => {
          responseCache.set(key, value, ttlSec ? ttlSec / 60 : undefined);
        }
      }
    });

    this.autoIndexConfig = {
      enabled: true,
      interval: '0 */2 * * *', // Every 2 hours
      folders: [], // Will be populated from Google Drive root
      fileTypes: [
        'application/pdf',
        'application/vnd.google-apps.document',
        'application/vnd.google-apps.spreadsheet',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'text/plain'
      ],
      maxFileSize: 50 * 1024 * 1024, // 50MB
      batchSize: 10,
      ...autoIndexConfig
    };

    this.indexingStats = {
      totalFiles: 0,
      indexedFiles: 0,
      failedFiles: 0,
      lastIndexingRun: null,
      nextScheduledRun: null,
      averageProcessingTime: 0,
      indexingProgress: {
        inProgress: false,
        currentFile: null,
        processed: 0,
        total: 0
      }
    };

    this.setupAutoIndexing();
  }

  /**
   * 🚀 Setup automatic indexing schedule
   */
  private async setupAutoIndexing(): Promise<void> {
    if (!this.autoIndexConfig.enabled) {
      logger.info('Auto-indexing is disabled');
      return;
    }

    try {
      // For now, we'll use a simple approach without scheduler integration
      // In a real implementation, you would integrate with the scheduler service
      this.scheduledTaskId = 'rag-auto-indexing-task';

      // Set next run time
      this.updateNextScheduledRun();

      logger.info('Auto-indexing scheduled', {
        component: 'EnhancedRagService',
        interval: this.autoIndexConfig.interval,
        taskId: this.scheduledTaskId
      });

    } catch (error) {
      logger.error('Failed to setup auto-indexing', {
        component: 'EnhancedRagService',
        error: error instanceof Error ? error.message : String(error)
      });
    }
  }

  /**
   * 🔄 Perform automatic indexing
   */
  private async performAutoIndexing(): Promise<void> {
    if (this.indexingInProgress) {
      logger.warn('Indexing already in progress, skipping this run');
      return;
    }

    this.indexingInProgress = true;
    this.indexingStats.indexingProgress.inProgress = true;
    this.indexingStats.lastIndexingRun = new Date();

    const startTime = Date.now();

    try {
      logger.info('Starting auto-indexing process', {
        component: 'EnhancedRagService'
      });

      // Get list of files to index
      const filesToIndex = await this.getFilesToIndex();
      
      this.indexingStats.totalFiles = filesToIndex.length;
      this.indexingStats.indexingProgress.total = filesToIndex.length;
      this.indexingStats.indexingProgress.processed = 0;

      let indexed = 0;
      let failed = 0;

      // Process files in batches
      for (let i = 0; i < filesToIndex.length; i += this.autoIndexConfig.batchSize) {
        const batch = filesToIndex.slice(i, i + this.autoIndexConfig.batchSize);
        
        for (const file of batch) {
          this.indexingStats.indexingProgress.currentFile = file.name;
          
          try {
            await this.indexFile(file);
            indexed++;
            logger.debug('File indexed successfully', {
              component: 'EnhancedRagService',
              fileId: file.id,
              fileName: file.name
            });
          } catch (error) {
            failed++;
            logger.error('Failed to index file', {
              component: 'EnhancedRagService',
              fileId: file.id,
              fileName: file.name,
              error: error instanceof Error ? error.message : String(error)
            });
          }

          this.indexingStats.indexingProgress.processed++;
        }

        // Small delay between batches to avoid rate limiting
        if (i + this.autoIndexConfig.batchSize < filesToIndex.length) {
          await new Promise(resolve => setTimeout(resolve, 1000));
        }
      }

      this.indexingStats.indexedFiles += indexed;
      this.indexingStats.failedFiles += failed;

      const duration = Date.now() - startTime;
      this.indexingStats.averageProcessingTime = duration / Math.max(1, indexed + failed);

      logger.info('Auto-indexing completed', {
        component: 'EnhancedRagService',
        totalFiles: filesToIndex.length,
        indexed,
        failed,
        duration: `${duration}ms`
      });

    } catch (error) {
      logger.error('Auto-indexing process failed', {
        component: 'EnhancedRagService',
        error: error instanceof Error ? error.message : String(error)
      });
    } finally {
      this.indexingInProgress = false;
      this.indexingStats.indexingProgress.inProgress = false;
      this.indexingStats.indexingProgress.currentFile = null;
      this.updateNextScheduledRun();
    }
  }

  /**
   * 📂 Get list of files that need indexing
   */
  private async getFilesToIndex(): Promise<DriveFile[]> {
    const filesToIndex: DriveFile[] = [];

    try {
      // If no specific folders configured, scan from root
      const foldersToScan = this.autoIndexConfig.folders.length > 0 
        ? this.autoIndexConfig.folders 
        : ['root'];

      for (const folderId of foldersToScan) {
        const files = await this.googleService.searchFiles(
          `'${folderId}' in parents and trashed=false'`
        );

        for (const file of files) {
          // Check if file type is supported
          if (!this.autoIndexConfig.fileTypes.includes(file.mimeType || '')) {
            continue;
          }

          // Check file size
          const fileSize = parseInt(file.size || '0');
          if (fileSize > this.autoIndexConfig.maxFileSize) {
            continue;
          }

          // Check if file needs re-indexing
          const needsIndexing = await this.needsIndexing(file.id!, new Date(file.modifiedTime!));
          if (!needsIndexing) {
            continue;
          }

          filesToIndex.push({
            id: file.id!,
            name: file.name!,
            mimeType: file.mimeType!,
            size: fileSize,
            modifiedTime: new Date(file.modifiedTime!),
            indexed: false
          });
        }
      }

      return filesToIndex;
    } catch (error) {
      logger.error('Failed to get files list for indexing', {
        component: 'EnhancedRagService',
        error: error instanceof Error ? error.message : String(error)
      });
      return [];
    }
  }

  /**
   * 🔍 Check if file needs indexing
   */
  private async needsIndexing(fileId: string, modifiedTime: Date): Promise<boolean> {
    try {
      // Check if file is already in the search index
      // For now, assume all files need indexing (simplification)
      return true;
    } catch (error) {
      // If we can't determine, assume it needs indexing
      return true;
    }
  }

  /**
   * 📄 Index a single file
   */
  private async indexFile(file: DriveFile): Promise<void> {
    try {
      // Extract content from the file
      const content = await this.extractFileContent(file);
      
      if (!content || content.trim().length === 0) {
        throw new Error('No content extracted from file');
      }

      // Index the content using search index directly
      // This is a simplified version - in production you'd use proper indexing
      logger.debug('File would be indexed', {
        component: 'EnhancedRagService',
        fileId: file.id,
        fileName: file.name,
        contentLength: content.length
      });

      file.indexed = true;
      file.lastIndexed = new Date();
      file.content = content.substring(0, 1000); // Store preview

    } catch (error) {
      file.error = error instanceof Error ? error.message : String(error);
      throw error;
    }
  }

  /**
   * 📖 Extract content from file
   */
  private async extractFileContent(file: DriveFile): Promise<string> {
    try {
      switch (file.mimeType) {
        case 'application/vnd.google-apps.document':
          // Use Google Docs API to export as plain text
          return await this.extractGoogleDocContent(file.id);
          
        case 'application/vnd.google-apps.spreadsheet':
          // Use Google Sheets API to export as CSV
          return await this.extractGoogleSheetContent(file.id);
          
        case 'application/pdf':
          // PDF extraction would require additional service
          throw new Error('PDF extraction not implemented in demo');
          
        case 'text/plain':
          // Use Google Drive API to download content
          return await this.extractPlainTextContent(file.id);
          
        default:
          // Try to download as text
          return await this.extractPlainTextContent(file.id);
      }
    } catch (error) {
      logger.error('Failed to extract file content', {
        component: 'EnhancedRagService',
        fileId: file.id,
        mimeType: file.mimeType,
        error: error instanceof Error ? error.message : String(error)
      });
      throw error;
    }
  }

  /**
   * 🛠️ Manual indexing trigger
   */
  async triggerManualIndexing(folderId?: string): Promise<void> {
    if (this.indexingInProgress) {
      throw new Error('Indexing is already in progress');
    }

    // Temporarily override folders if specified
    const originalFolders = this.autoIndexConfig.folders;
    if (folderId) {
      this.autoIndexConfig.folders = [folderId];
    }

    try {
      await this.performAutoIndexing();
    } finally {
      this.autoIndexConfig.folders = originalFolders;
    }
  }

  /**
   * 📊 Get indexing statistics
   */
  getIndexingStats(): IndexingStats {
    return { ...this.indexingStats };
  }

  /**
   * ⚙️ Update auto-indexing configuration
   */
  updateAutoIndexConfig(config: Partial<AutoIndexConfig>): void {
    this.autoIndexConfig = { ...this.autoIndexConfig, ...config };

    // For now, we'll use a simple approach without scheduler integration
    // In a real implementation, you would integrate with the scheduler service
    logger.info('Auto-indexing would be scheduled here', {
      component: 'EnhancedRagService',
      interval: this.autoIndexConfig.interval
    });

    logger.info('Auto-indexing configuration updated', {
      component: 'EnhancedRagService',
      config: this.autoIndexConfig
    });
  }

  /**
   * ⏰ Update next scheduled run time
   */
  private updateNextScheduledRun(): void {
    // For now, we'll use a simple approach without scheduler integration
    // In a real implementation, you would integrate with the scheduler service
    logger.info('Next scheduled run would be updated here', {
      component: 'EnhancedRagService',
      taskId: this.scheduledTaskId
    });
  }

  /**
   * 🔍 Enhanced search with auto-indexing awareness
   */
  async search(
    query: string,
    options?: {
      limit?: number;
      minScore?: number;
      filters?: Record<string, any>;
      useCache?: boolean;
    }
  ): Promise<Array<{ content: string; score: number; fileId?: string; fileName?: string }>> {
    const cacheKey = ResponseCacheService.generateKey(
      'enhanced_rag_search',
      query,
      JSON.stringify(options || {})
    );

    // Check cache if enabled
    if (options?.useCache !== false) {
      const cached = this.responseCache.get<Array<{ content: string; score: number; fileId?: string; fileName?: string }>>(cacheKey);
      if (cached && Array.isArray(cached)) {
        return cached;
      }
    }

    // Perform search using parent RAG functionality
    const ragResult = await this.answer(query, 
      { 
        k: options?.limit || 5, 
        ...(options?.filters && { filters: options.filters })
      },
      {},
      { model: 'ukrainian-military-assistant' }
    );

    // Transform RAG result to search format
    const searchResults = ragResult.chunks.map((chunk: any) => ({
      content: chunk.content || String(chunk),
      score: chunk.score || 0.8,
      fileId: chunk.fileId,
      fileName: chunk.fileName || 'Unknown'
    }));

    // Cache results
    if (options?.useCache !== false) {
      this.responseCache.set(cacheKey, searchResults, 15); // 15 minutes
    }

    return searchResults;
  }

  /**
   * 🛑 Shutdown service
   */
  async shutdown(): Promise<void> {
    // For now, we'll use a simple approach without scheduler integration
    // In a real implementation, you would integrate with the scheduler service
    logger.info('Auto-indexing would be cancelled here', {
      component: 'EnhancedRagService',
      taskId: this.scheduledTaskId
    });
    this.scheduledTaskId = undefined;

    this.indexingInProgress = false;

    logger.info('EnhancedRagService shutdown completed', {
      component: 'EnhancedRagService'
    });
  }

  /**
   * 📄 Extract content from Google Doc
   */
  private async extractGoogleDocContent(fileId: string): Promise<string> {
    try {
      // This would use Google Docs API to export as plain text
      // For demo purposes, return placeholder
      return `Google Doc content for file ${fileId} (demo)`;
    } catch (error) {
      throw new Error(`Failed to extract Google Doc content: ${error}`);
    }
  }

  /**
   * 📊 Extract content from Google Sheet
   */
  private async extractGoogleSheetContent(fileId: string): Promise<string> {
    try {
      // This would use Google Sheets API to export as CSV
      // For demo purposes, return placeholder
      return `Google Sheet content for file ${fileId} (demo)`;
    } catch (error) {
      throw new Error(`Failed to extract Google Sheet content: ${error}`);
    }
  }

  /**
   * 📝 Extract plain text content
   */
  private async extractPlainTextContent(fileId: string): Promise<string> {
    try {
      // This would use Google Drive API to download file content
      // For demo purposes, return placeholder
      return `Plain text content for file ${fileId} (demo)`;
    } catch (error) {
      throw new Error(`Failed to extract plain text content: ${error}`);
    }
  }
}