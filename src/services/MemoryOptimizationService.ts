import { BaseService } from '@/core/BaseService';
import type { BotConfig } from '@/types';
import logger from '@/utils/logger';

export interface MemoryStats {
  heapUsed: number;
  heapTotal: number;
  rss: number;
  external: number;
  arrayBuffers: number;
  percentageUsed: number;
}

export interface MemoryOptimizationConfig {
  maxHeapUsage: number; // in bytes
  gcInterval: number; // in milliseconds
  cleanupThreshold: number; // percentage (0-100)
  streamChunkSize: number; // in bytes
  compressionThreshold: number; // in bytes
}

export interface DocumentChunk {
  id: string;
  content: string;
  position: number;
  size: number;
  compressed: boolean;
}

export interface MemoryPressureEvent {
  timestamp: Date;
  memoryUsage: MemoryStats;
  actionTaken: string;
  threshold: number;
}

export class MemoryOptimizationService extends BaseService {
  private config: MemoryOptimizationConfig;
  private documentCache: Map<string, DocumentChunk[]> = new Map();
  private memoryPressureEvents: MemoryPressureEvent[] = [];
  private gcIntervalId: NodeJS.Timeout | null = null;
  private readonly MAX_PRESSURE_EVENTS = 100;
  private readonly MAX_CACHE_SIZE = 1000;
  
  constructor(config: BotConfig) {
    super('MemoryOptimizationService', config);
    
    this.config = {
      maxHeapUsage: config.memory?.maxHeapUsage || 1024 * 1024 * 1024, // 1GB default
      gcInterval: config.memory?.gcInterval || 60000, // 1 minute default
      cleanupThreshold: config.memory?.cleanupThreshold || 80, // 80% default
      streamChunkSize: config.memory?.streamChunkSize || 64 * 1024, // 64KB default
      compressionThreshold: config.memory?.compressionThreshold || 1024 * 1024 // 1MB default
    };
  }

  /**
   * Initialize memory optimization service
   */
  protected async onInitialize(): Promise<void> {
    // Start garbage collection interval
    this.startGcInterval();
    
    logger.info('Memory optimization service initialized', {
      component: 'MemoryOptimizationService',
      config: this.config
    });
  }

  /**
   * Start garbage collection interval
   */
  private startGcInterval(): void {
    if (this.gcIntervalId) {
      clearInterval(this.gcIntervalId);
    }
    
    this.gcIntervalId = setInterval(() => {
      this.performGarbageCollection();
    }, this.config.gcInterval);
    
    logger.debug('Garbage collection interval started', {
      component: 'MemoryOptimizationService',
      interval: this.config.gcInterval
    });
  }

  /**
   * Get current memory statistics
   */
  getMemoryStats(): MemoryStats {
    const memoryUsage = process.memoryUsage();
    
    return {
      heapUsed: memoryUsage.heapUsed,
      heapTotal: memoryUsage.heapTotal,
      rss: memoryUsage.rss,
      external: memoryUsage.external,
      arrayBuffers: memoryUsage.arrayBuffers || 0,
      percentageUsed: Math.round((memoryUsage.heapUsed / memoryUsage.heapTotal) * 100)
    };
  }

  /**
   * Check if memory usage is above threshold
   */
  isMemoryPressureHigh(): boolean {
    const stats = this.getMemoryStats();
    return stats.percentageUsed > this.config.cleanupThreshold;
  }

  /**
   * Perform garbage collection and cleanup
   */
  performGarbageCollection(): void {
    const stats = this.getMemoryStats();
    
    logger.debug('Performing garbage collection', {
      component: 'MemoryOptimizationService',
      memoryUsage: `${stats.percentageUsed}%`,
      heapUsed: `${Math.round(stats.heapUsed / 1024 / 1024)}MB`
    });
    
    // Check if we need to take action
    if (this.isMemoryPressureHigh()) {
      this.handleMemoryPressure(stats);
    }
    
    // Force garbage collection if available (Node.js flag --expose-gc needed)
    if (global.gc) {
      global.gc();
      logger.debug('Forced garbage collection', {
        component: 'MemoryOptimizationService'
      });
    }
  }

  /**
   * Handle memory pressure situations
   */
  private handleMemoryPressure(stats: MemoryStats): void {
    logger.warn('High memory pressure detected', {
      component: 'MemoryOptimizationService',
      memoryUsage: `${stats.percentageUsed}%`,
      threshold: `${this.config.cleanupThreshold}%`
    });
    
    // Record the memory pressure event
    this.recordMemoryPressureEvent(stats, 'cleanup_cache');
    
    // Clean up document cache
    this.cleanupDocumentCache();
    
    // If still high, take more aggressive action
    if (this.isMemoryPressureHigh()) {
      this.recordMemoryPressureEvent(this.getMemoryStats(), 'aggressive_cleanup');
      this.aggressiveCleanup();
    }
  }

  /**
   * Record a memory pressure event
   */
  private recordMemoryPressureEvent(stats: MemoryStats, action: string): void {
    const event: MemoryPressureEvent = {
      timestamp: new Date(),
      memoryUsage: { ...stats },
      actionTaken: action,
      threshold: this.config.cleanupThreshold
    };
    
    this.memoryPressureEvents.push(event);
    
    // Maintain event log size
    if (this.memoryPressureEvents.length > this.MAX_PRESSURE_EVENTS) {
      this.memoryPressureEvents = this.memoryPressureEvents.slice(-this.MAX_PRESSURE_EVENTS);
    }
  }

  /**
   * Clean up document cache
   */
  private cleanupDocumentCache(): void {
    const initialSize = this.documentCache.size;
    
    // Remove oldest entries if we're above the cache size limit
    if (this.documentCache.size > this.MAX_CACHE_SIZE) {
      const keysToRemove = Array.from(this.documentCache.keys())
        .slice(0, Math.floor(this.MAX_CACHE_SIZE * 0.1)); // Remove 10%
      
      for (const key of keysToRemove) {
        this.documentCache.delete(key);
      }
      
      logger.info('Document cache partially cleared', {
        component: 'MemoryOptimizationService',
        removedEntries: keysToRemove.length,
        remainingEntries: this.documentCache.size
      });
    }
    
    // If we're still under pressure, clear more aggressively
    if (this.isMemoryPressureHigh() && initialSize === this.documentCache.size) {
      // Clear 50% of the cache
      const keysToRemove = Array.from(this.documentCache.keys())
        .slice(0, Math.floor(this.documentCache.size * 0.5));
      
      for (const key of keysToRemove) {
        this.documentCache.delete(key);
      }
      
      logger.info('Document cache aggressively cleared', {
        component: 'MemoryOptimizationService',
        removedEntries: keysToRemove.length,
        remainingEntries: this.documentCache.size
      });
    }
  }

  /**
   * Aggressive cleanup when memory pressure is still high
   */
  private aggressiveCleanup(): void {
    // Clear all document caches
    const clearedEntries = this.documentCache.size;
    this.documentCache.clear();
    
    logger.warn('Aggressive memory cleanup performed', {
      component: 'MemoryOptimizationService',
      clearedEntries
    });
  }

  /**
   * Process a large document with memory optimization
   */
  async processLargeDocument(documentId: string, content: string): Promise<DocumentChunk[]> {
    try {
      // Check if document is already cached
      const cachedChunks = this.documentCache.get(documentId);
      if (cachedChunks) {
        logger.debug('Returning cached document chunks', {
          component: 'MemoryOptimizationService',
          documentId,
          chunkCount: cachedChunks.length
        });
        return cachedChunks;
      }
      
      // Split document into chunks
      const chunks = this.splitDocumentIntoChunks(documentId, content);
      
      // Compress large chunks if needed
      const processedChunks = await Promise.all(
        chunks.map(chunk => this.processChunk(chunk))
      );
      
      // Cache the processed chunks
      this.documentCache.set(documentId, processedChunks);
      
      // Check memory after processing
      if (this.isMemoryPressureHigh()) {
        this.performGarbageCollection();
      }
      
      logger.info('Large document processed successfully', {
        component: 'MemoryOptimizationService',
        documentId,
        chunkCount: processedChunks.length,
        totalSize: content.length
      });
      
      return processedChunks;
    } catch (error) {
      logger.error('Error processing large document', {
        component: 'MemoryOptimizationService',
        documentId,
        error: error instanceof Error ? error.message : String(error)
      });
      
      throw error;
    }
  }

  /**
   * Split document into manageable chunks
   */
  private splitDocumentIntoChunks(documentId: string, content: string): DocumentChunk[] {
    const chunks: DocumentChunk[] = [];
    const chunkSize = this.config.streamChunkSize;
    
    for (let i = 0; i < content.length; i += chunkSize) {
      const chunkContent = content.substring(i, i + chunkSize);
      const chunk: DocumentChunk = {
        id: `${documentId}-chunk-${i / chunkSize}`,
        content: chunkContent,
        position: i,
        size: chunkContent.length,
        compressed: false
      };
      
      chunks.push(chunk);
    }
    
    return chunks;
  }

  /**
   * Process individual chunk (compress if large)
   */
  private async processChunk(chunk: DocumentChunk): Promise<DocumentChunk> {
    // Only compress chunks larger than threshold
    if (chunk.size > this.config.compressionThreshold) {
      try {
        const compressedContent = await this.compressContent(chunk.content);
        return {
          ...chunk,
          content: compressedContent,
          compressed: true
        };
      } catch (error) {
        logger.warn('Failed to compress chunk, storing uncompressed', {
          component: 'MemoryOptimizationService',
          chunkId: chunk.id,
          error: error instanceof Error ? error.message : String(error)
        });
        
        return chunk;
      }
    }
    
    return chunk;
  }

  /**
   * Compress content using a simple algorithm
   */
  private async compressContent(content: string): Promise<string> {
    // In a real implementation, you would use a proper compression library
    // For now, we'll simulate compression by removing extra whitespace
    return content.replace(/\s+/g, ' ');
  }

  /**
   * Decompress content
   */
  private async decompressContent(content: string): Promise<string> {
    // In a real implementation, you would use a proper decompression library
    // For now, we'll just return the content as-is
    return content;
  }

  /**
   * Get document chunks (decompress if needed)
   */
  async getDocumentChunks(documentId: string): Promise<DocumentChunk[]> {
    const chunks = this.documentCache.get(documentId);
    
    if (!chunks) {
      return [];
    }
    
    // Decompress chunks if needed
    const decompressedChunks = await Promise.all(
      chunks.map(async chunk => {
        if (chunk.compressed) {
          try {
            const decompressedContent = await this.decompressContent(chunk.content);
            return {
              ...chunk,
              content: decompressedContent,
              compressed: false
            };
          } catch (error) {
            logger.warn('Failed to decompress chunk', {
              component: 'MemoryOptimizationService',
              chunkId: chunk.id,
              error: error instanceof Error ? error.message : String(error)
            });
            return chunk;
          }
        }
        return chunk;
      })
    );
    
    return decompressedChunks;
  }

  /**
   * Stream process large content
   */
  async *streamProcessContent(content: string): AsyncGenerator<string, void, unknown> {
    const chunkSize = this.config.streamChunkSize;
    
    for (let i = 0; i < content.length; i += chunkSize) {
      const chunk = content.substring(i, i + chunkSize);
      
      // Process the chunk (in a real implementation, you might do more here)
      const processedChunk = chunk.trim();
      
      yield processedChunk;
      
      // Check memory pressure during streaming
      if (this.isMemoryPressureHigh()) {
        this.performGarbageCollection();
      }
    }
  }

  /**
   * Get memory pressure events
   */
  getMemoryPressureEvents(limit?: number): MemoryPressureEvent[] {
    const events = [...this.memoryPressureEvents];
    return limit ? events.slice(-limit) : events;
  }

  /**
   * Get document cache statistics
   */
  getCacheStats(): {
    cachedDocuments: number;
    totalChunks: number;
    averageChunksPerDocument: number;
  } {
    const cachedDocuments = this.documentCache.size;
    let totalChunks = 0;
    
    for (const chunks of this.documentCache.values()) {
      totalChunks += chunks.length;
    }
    
    const averageChunksPerDocument = cachedDocuments > 0 ? totalChunks / cachedDocuments : 0;
    
    return {
      cachedDocuments,
      totalChunks,
      averageChunksPerDocument
    };
  }

  /**
   * Clear document cache
   */
  clearDocumentCache(): void {
    const clearedEntries = this.documentCache.size;
    this.documentCache.clear();
    
    logger.info('Document cache cleared', {
      component: 'MemoryOptimizationService',
      clearedEntries
    });
  }

  /**
   * Generate memory optimization report
   */
  generateReport(): {
    memoryStats: MemoryStats;
    cacheStats: {
      cachedDocuments: number;
      totalChunks: number;
      averageChunksPerDocument: number;
    };
    pressureEvents: MemoryPressureEvent[];
    config: MemoryOptimizationConfig;
    recommendations: string[];
  } {
    const memoryStats = this.getMemoryStats();
    const cacheStats = this.getCacheStats();
    const pressureEvents = this.getMemoryPressureEvents(10);
    
    const recommendations: string[] = [];
    
    // Generate recommendations based on current state
    if (memoryStats.percentageUsed > 80) {
      recommendations.push('Memory usage is high - consider increasing available memory or optimizing document processing');
    }
    
    if (cacheStats.cachedDocuments > this.MAX_CACHE_SIZE * 0.8) {
      recommendations.push('Document cache is nearly full - consider adjusting cache size limits');
    }
    
    if (pressureEvents.length > 5) {
      recommendations.push(`Frequent memory pressure events (${pressureEvents.length}) detected - review document processing patterns`);
    }
    
    return {
      memoryStats,
      cacheStats,
      pressureEvents,
      config: { ...this.config },
      recommendations
    };
  }

  /**
   * Adjust configuration dynamically
   */
  updateConfig(newConfig: Partial<MemoryOptimizationConfig>): void {
    Object.assign(this.config, newConfig);
    
    logger.info('Memory optimization configuration updated', {
      component: 'MemoryOptimizationService',
      newConfig
    });
    
    // Restart GC interval if interval changed
    if (newConfig.gcInterval !== undefined) {
      this.startGcInterval();
    }
  }

  /**
   * Shutdown the service
   */
  async shutdown(): Promise<void> {
    if (this.gcIntervalId) {
      clearInterval(this.gcIntervalId);
      this.gcIntervalId = null;
    }
    
    // Clear caches
    this.documentCache.clear();
    this.memoryPressureEvents = [];
    
    logger.info('Memory optimization service shutdown', {
      component: 'MemoryOptimizationService'
    });
  }

  /**
   * Force immediate garbage collection
   */
  forceGarbageCollection(): void {
    if (global.gc) {
      const beforeStats = this.getMemoryStats();
      global.gc();
      const afterStats = this.getMemoryStats();
      
      logger.info('Forced garbage collection completed', {
        component: 'MemoryOptimizationService',
        before: `${Math.round(beforeStats.heapUsed / 1024 / 1024)}MB`,
        after: `${Math.round(afterStats.heapUsed / 1024 / 1024)}MB`,
        freed: `${Math.round((beforeStats.heapUsed - afterStats.heapUsed) / 1024 / 1024)}MB`
      });
    } else {
      logger.warn('Garbage collection not exposed - start Node.js with --expose-gc flag', {
        component: 'MemoryOptimizationService'
      });
    }
  }
}