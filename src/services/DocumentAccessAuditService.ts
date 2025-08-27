import { BaseService } from '@/core/BaseService';
import type { BotConfig } from '@/types';
import type { DriveFile } from '@/types/drive';
import logger from '@/utils/logger';
import { CacheService } from '@/services/CacheService';

export interface DocumentAccessLog {
  id: string;
  userId: string;
  userName: string;
  fileId: string;
  fileName: string;
  accessType: 'view' | 'edit' | 'download' | 'share' | 'delete' | 'move' | 'copy';
  timestamp: Date;
  ipAddress?: string;
  userAgent?: string;
  success: boolean;
  errorMessage?: string;
  duration?: number; // in milliseconds
  fileSize?: number;
  fileType?: string;
}

export interface DocumentAccessStats {
  totalAccesses: number;
  successfulAccesses: number;
  failedAccesses: number;
  uniqueUsers: number;
  accessByType: Record<string, number>;
  accessByTime: Record<string, number>;
}

export interface AccessQueryParams {
  userId?: string;
  fileId?: string;
  accessType?: string;
  startDate?: Date;
  endDate?: Date;
  limit?: number;
  offset?: number;
}

export class DocumentAccessAuditService extends BaseService {
  private cache: CacheService | null = null;
  private accessLogs: DocumentAccessLog[] = [];
  private readonly MAX_LOG_HISTORY = 10000;
  private userAccessMap = new Map<string, Set<string>>(); // userId -> Set of fileIds
  private fileAccessMap = new Map<string, Set<string>>(); // fileId -> Set of userIds

  constructor(config: BotConfig) {
    super('DocumentAccessAuditService', config);
  }

  /**
   * Initialize the service with dependencies
   */
  initializeServices(cache?: CacheService): void {
    this.cache = cache || null;
    
    logger.info('DocumentAccessAuditService initialized', {
      component: 'DocumentAccessAuditService'
    });
  }

  /**
   * Log a document access event
   */
  async logAccess(accessLog: Omit<DocumentAccessLog, 'id' | 'timestamp'>): Promise<void> {
    try {
      const logEntry: DocumentAccessLog = {
        id: this.generateId(),
        timestamp: new Date(),
        ...accessLog
      };

      // Add to in-memory storage
      this.accessLogs.push(logEntry);
      
      // Maintain history limit
      if (this.accessLogs.length > this.MAX_LOG_HISTORY) {
        this.accessLogs = this.accessLogs.slice(-this.MAX_LOG_HISTORY);
      }

      // Update user-file access mapping
      if (!this.userAccessMap.has(logEntry.userId)) {
        this.userAccessMap.set(logEntry.userId, new Set());
      }
      this.userAccessMap.get(logEntry.userId)?.add(logEntry.fileId);

      // Update file-user access mapping
      if (!this.fileAccessMap.has(logEntry.fileId)) {
        this.fileAccessMap.set(logEntry.fileId, new Set());
      }
      this.fileAccessMap.get(logEntry.fileId)?.add(logEntry.userId);

      // Log the access event
      logger.security(`Document ${logEntry.accessType} access`, logEntry.userId, {
        component: 'DocumentAccessAuditService',
        fileId: logEntry.fileId,
        fileName: logEntry.fileName,
        success: logEntry.success,
        errorMessage: logEntry.errorMessage,
        duration: logEntry.duration,
        fileSize: logEntry.fileSize,
        fileType: logEntry.fileType,
        severity: logEntry.success ? 'low' : 'medium'
      });

      // Cache the log entry if cache is available
      if (this.cache) {
        const cacheKey = `document_access:${logEntry.id}`;
        await this.cache.set(cacheKey, logEntry, 86400); // Cache for 24 hours
      }
    } catch (error) {
      logger.error('Failed to log document access', {
        component: 'DocumentAccessAuditService',
        error: error instanceof Error ? error.message : String(error)
      });
    }
  }

  /**
   * Get access logs based on query parameters
   */
  async getAccessLogs(params: AccessQueryParams = {}): Promise<DocumentAccessLog[]> {
    try {
      let filteredLogs = [...this.accessLogs].reverse(); // Most recent first

      // Apply filters
      if (params.userId) {
        filteredLogs = filteredLogs.filter(log => log.userId === params.userId);
      }

      if (params.fileId) {
        filteredLogs = filteredLogs.filter(log => log.fileId === params.fileId);
      }

      if (params.accessType) {
        filteredLogs = filteredLogs.filter(log => log.accessType === params.accessType);
      }

      if (params.startDate) {
        filteredLogs = filteredLogs.filter(log => log.timestamp >= params.startDate);
      }

      if (params.endDate) {
        filteredLogs = filteredLogs.filter(log => log.timestamp <= params.endDate);
      }

      // Apply pagination
      const limit = params.limit || 50;
      const offset = params.offset || 0;
      filteredLogs = filteredLogs.slice(offset, offset + limit);

      // Try to fetch from cache if available
      if (this.cache) {
        const cachedLogs: DocumentAccessLog[] = [];
        for (const log of filteredLogs) {
          const cacheKey = `document_access:${log.id}`;
          const cachedLog = await this.cache.get<DocumentAccessLog>(cacheKey);
          if (cachedLog) {
            cachedLogs.push(cachedLog);
          }
        }
        // If we have all logs in cache, return them
        if (cachedLogs.length === filteredLogs.length) {
          return cachedLogs;
        }
      }

      return filteredLogs;
    } catch (error) {
      logger.error('Failed to retrieve access logs', {
        component: 'DocumentAccessAuditService',
        error: error instanceof Error ? error.message : String(error)
      });
      return [];
    }
  }

  /**
   * Get statistics about document access
   */
  async getAccessStats(params: AccessQueryParams = {}): Promise<DocumentAccessStats> {
    try {
      const logs = await this.getAccessLogs(params);
      
      const stats: DocumentAccessStats = {
        totalAccesses: logs.length,
        successfulAccesses: logs.filter(log => log.success).length,
        failedAccesses: logs.filter(log => !log.success).length,
        uniqueUsers: new Set(logs.map(log => log.userId)).size,
        accessByType: {},
        accessByTime: {}
      };

      // Calculate access by type
      for (const log of logs) {
        stats.accessByType[log.accessType] = (stats.accessByType[log.accessType] || 0) + 1;
      }

      // Calculate access by time (hour of day)
      for (const log of logs) {
        const hour = log.timestamp.getHours().toString();
        stats.accessByTime[hour] = (stats.accessByTime[hour] || 0) + 1;
      }

      return stats;
    } catch (error) {
      logger.error('Failed to calculate access stats', {
        component: 'DocumentAccessAuditService',
        error: error instanceof Error ? error.message : String(error)
      });
      
      return {
        totalAccesses: 0,
        successfulAccesses: 0,
        failedAccesses: 0,
        uniqueUsers: 0,
        accessByType: {},
        accessByTime: {}
      };
    }
  }

  /**
   * Get files accessed by a specific user
   */
  getFilesAccessedByUser(userId: string): string[] {
    const fileIds = this.userAccessMap.get(userId);
    return fileIds ? Array.from(fileIds) : [];
  }

  /**
   * Get users who accessed a specific file
   */
  getUsersWhoAccessedFile(fileId: string): string[] {
    const userIds = this.fileAccessMap.get(fileId);
    return userIds ? Array.from(userIds) : [];
  }

  /**
   * Generate a unique ID
   */
  private generateId(): string {
    return Date.now().toString(36) + Math.random().toString(36).substr(2, 9);
  }

  /**
   * Export access logs to a file (simulated)
   */
  async exportAccessLogs(params: AccessQueryParams = {}): Promise<string> {
    try {
      const logs = await this.getAccessLogs(params);
      
      // In a real implementation, this would create an actual file
      // For now, we'll return a JSON string representation
      const exportData = {
        exportedAt: new Date().toISOString(),
        queryParams: params,
        logs: logs.map(log => ({
          ...log,
          timestamp: log.timestamp.toISOString()
        }))
      };
      
      logger.info('Document access logs exported', {
        component: 'DocumentAccessAuditService',
        logCount: logs.length
      });
      
      return JSON.stringify(exportData, null, 2);
    } catch (error) {
      logger.error('Failed to export access logs', {
        component: 'DocumentAccessAuditService',
        error: error instanceof Error ? error.message : String(error)
      });
      throw error;
    }
  }

  // === BaseServiceClass required methods ===
  
  protected async onInitialize(): Promise<void> {
    logger.info('DocumentAccessAuditService initialized', {
      component: 'DocumentAccessAuditService'
    });
  }

  protected async onShutdown(): Promise<void> {
    logger.info('DocumentAccessAuditService shut down', {
      component: 'DocumentAccessAuditService'
    });
  }

  protected async onHealthCheck(): Promise<any> {
    return {
      healthy: true,
      service: this.name,
      logCount: this.accessLogs.length,
      userCount: this.userAccessMap.size,
      fileCount: this.fileAccessMap.size
    };
  }

  protected onGetStats(): any {
    return {
      logCount: this.accessLogs.length,
      userCount: this.userAccessMap.size,
      fileCount: this.fileAccessMap.size
    };
  }
}