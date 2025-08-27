import { BaseService } from '@/core/BaseService';
import type { BotConfig } from '@/types';
import type { DriveFile } from '@/types/drive';
import logger from '@/utils/logger';

export interface DocumentAccessAuditRecord {
  id: string;
  fileId: string;
  fileName: string;
  userId: string;
  userName: string;
  action: 'view' | 'download' | 'search' | 'analyze' | 'export' | 'compare' | 'summarize' | 'delete';
  timestamp: Date;
  ipAddress?: string;
  userAgent?: string;
  sessionId: string;
  metadata?: Record<string, any>;
}

export interface AccessSummary {
  totalAccesses: number;
  uniqueUsers: number;
  popularDocuments: { fileId: string; fileName: string; accessCount: number }[];
  actionDistribution: Record<string, number>;
  timeSeries: { date: string; count: number }[];
}

export interface ComplianceReport {
  reportId: string;
  generatedAt: Date;
  periodStart: Date;
  periodEnd: Date;
  summary: AccessSummary;
  detailedRecords: DocumentAccessAuditRecord[];
  sensitiveDocumentAccesses: DocumentAccessAuditRecord[];
  flaggedActivities: DocumentAccessAuditRecord[];
}

export class DocumentAccessAuditService extends BaseService {
  private auditRecords: DocumentAccessAuditRecord[] = [];
  private readonly MAX_RECORDS = 50000; // Keep last 50,000 records in memory
  private readonly SENSITIVE_KEYWORDS = ['confidential', 'secret', 'private', 'internal', 'restricted'];
  
  constructor(config: BotConfig) {
    super('DocumentAccessAuditService', config);
  }

  /**
   * Log a document access event
   */
  logDocumentAccess(args: {
    file: DriveFile;
    userId: string;
    userName: string;
    action: DocumentAccessAuditRecord['action'];
    sessionId: string;
    ipAddress?: string;
    userAgent?: string;
    metadata?: Record<string, any>;
  }): void {
    try {
      const record: DocumentAccessAuditRecord = {
        id: this.generateId(),
        fileId: args.file.id,
        fileName: args.file.name || 'Untitled',
        userId: args.userId,
        userName: args.userName,
        action: args.action,
        timestamp: new Date(),
        ipAddress: args.ipAddress,
        userAgent: args.userAgent,
        sessionId: args.sessionId,
        metadata: args.metadata
      };

      // Add record to audit log
      this.auditRecords.push(record);

      // Maintain record limit
      if (this.auditRecords.length > this.MAX_RECORDS) {
        this.auditRecords = this.auditRecords.slice(-this.MAX_RECORDS);
      }

      // Log the access
      logger.info('Document access logged', {
        component: 'DocumentAccessAuditService',
        fileId: args.file.id,
        fileName: args.file.name,
        userId: args.userId,
        action: args.action,
        sessionId: args.sessionId
      });

      // Check for sensitive document access
      if (this.isSensitiveDocument(args.file)) {
        logger.warn('Sensitive document accessed', {
          component: 'DocumentAccessAuditService',
          fileId: args.file.id,
          fileName: args.file.name,
          userId: args.userId,
          action: args.action
        });
      }
    } catch (error) {
      logger.error('Error logging document access', {
        component: 'DocumentAccessAuditService',
        fileId: args.file?.id,
        userId: args.userId,
        error: error instanceof Error ? error.message : String(error)
      });
    }
  }

  /**
   * Generate an access summary report
   */
  generateAccessSummary(options?: {
    startDate?: Date;
    endDate?: Date;
    userId?: string;
    fileId?: string;
  }): AccessSummary {
    try {
      let filteredRecords = this.auditRecords;

      // Apply date filters
      if (options?.startDate) {
        filteredRecords = filteredRecords.filter(record => record.timestamp >= options.startDate!);
      }
      
      if (options?.endDate) {
        filteredRecords = filteredRecords.filter(record => record.timestamp <= options.endDate!);
      }
      
      // Apply user filter
      if (options?.userId) {
        filteredRecords = filteredRecords.filter(record => record.userId === options.userId);
      }
      
      // Apply file filter
      if (options?.fileId) {
        filteredRecords = filteredRecords.filter(record => record.fileId === options.fileId);
      }

      // Calculate summary statistics
      const totalAccesses = filteredRecords.length;
      const uniqueUsers = new Set(filteredRecords.map(record => record.userId)).size;
      
      // Calculate popular documents
      const documentAccessCount: Record<string, { fileName: string; count: number }> = {};
      filteredRecords.forEach(record => {
        if (!documentAccessCount[record.fileId]) {
          documentAccessCount[record.fileId] = { fileName: record.fileName, count: 0 };
        }
        documentAccessCount[record.fileId].count++;
      });
      
      const popularDocuments = Object.entries(documentAccessCount)
        .map(([fileId, data]) => ({
          fileId,
          fileName: data.fileName,
          accessCount: data.count
        }))
        .sort((a, b) => b.accessCount - a.accessCount)
        .slice(0, 10); // Top 10 most accessed documents
      
      // Calculate action distribution
      const actionDistribution: Record<string, number> = {};
      filteredRecords.forEach(record => {
        actionDistribution[record.action] = (actionDistribution[record.action] || 0) + 1;
      });
      
      // Generate time series data (grouped by day)
      const timeSeries: { date: string; count: number }[] = [];
      const dateCounts: Record<string, number> = {};
      
      filteredRecords.forEach(record => {
        const date = record.timestamp.toISOString().split('T')[0]; // YYYY-MM-DD
        dateCounts[date] = (dateCounts[date] || 0) + 1;
      });
      
      Object.entries(dateCounts).forEach(([date, count]) => {
        timeSeries.push({ date, count });
      });
      
      // Sort time series by date
      timeSeries.sort((a, b) => a.date.localeCompare(b.date));

      return {
        totalAccesses,
        uniqueUsers,
        popularDocuments,
        actionDistribution,
        timeSeries
      };
    } catch (error) {
      logger.error('Error generating access summary', {
        component: 'DocumentAccessAuditService',
        error: error instanceof Error ? error.message : String(error)
      });
      
      // Return empty summary on error
      return {
        totalAccesses: 0,
        uniqueUsers: 0,
        popularDocuments: [],
        actionDistribution: {},
        timeSeries: []
      };
    }
  }

  /**
   * Generate a compliance report
   */
  async generateComplianceReport(options?: {
    startDate?: Date;
    endDate?: Date;
  }): Promise<ComplianceReport> {
    try {
      const startDate = options?.startDate || new Date(Date.now() - 30 * 24 * 60 * 60 * 1000); // Last 30 days
      const endDate = options?.endDate || new Date();
      
      // Get filtered records for the period
      const filteredRecords = this.auditRecords.filter(
        record => record.timestamp >= startDate && record.timestamp <= endDate
      );
      
      // Generate access summary
      const summary = this.generateAccessSummary({ startDate, endDate });
      
      // Identify sensitive document accesses
      const sensitiveDocumentAccesses = filteredRecords.filter(record => 
        this.SENSITIVE_KEYWORDS.some(keyword => 
          record.fileName.toLowerCase().includes(keyword)
        )
      );
      
      // Flag potentially problematic activities
      const flaggedActivities: DocumentAccessAuditRecord[] = [];
      
      // Flag access outside business hours (9 AM - 6 PM)
      const afterHoursAccesses = filteredRecords.filter(record => {
        const hour = record.timestamp.getHours();
        return hour < 9 || hour >= 18;
      });
      
      flaggedActivities.push(...afterHoursAccesses);
      
      // Flag high frequency access by single user
      const userAccessCounts: Record<string, number> = {};
      filteredRecords.forEach(record => {
        userAccessCounts[record.userId] = (userAccessCounts[record.userId] || 0) + 1;
      });
      
      const highFrequencyUsers = Object.entries(userAccessCounts)
        .filter(([_, count]) => count > 100) // More than 100 accesses in period
        .map(([userId]) => userId);
      
      const highFrequencyAccesses = filteredRecords.filter(record => 
        highFrequencyUsers.includes(record.userId)
      );
      
      flaggedActivities.push(...highFrequencyAccesses);
      
      // Remove duplicates from flagged activities
      const uniqueFlaggedActivities = Array.from(
        new Map(flaggedActivities.map(item => [item.id, item])).values()
      );

      const report: ComplianceReport = {
        reportId: this.generateId(),
        generatedAt: new Date(),
        periodStart: startDate,
        periodEnd: endDate,
        summary,
        detailedRecords: filteredRecords,
        sensitiveDocumentAccesses,
        flaggedActivities: uniqueFlaggedActivities
      };
      
      logger.info('Compliance report generated', {
        component: 'DocumentAccessAuditService',
        reportId: report.reportId,
        periodStart: startDate.toISOString(),
        periodEnd: endDate.toISOString(),
        totalRecords: filteredRecords.length
      });
      
      return report;
    } catch (error) {
      logger.error('Error generating compliance report', {
        component: 'DocumentAccessAuditService',
        error: error instanceof Error ? error.message : String(error)
      });
      
      throw error;
    }
  }

  /**
   * Get audit records with filtering and pagination
   */
  getAuditRecords(options?: {
    page?: number;
    limit?: number;
    userId?: string;
    fileId?: string;
    action?: DocumentAccessAuditRecord['action'];
    startDate?: Date;
    endDate?: Date;
  }): { records: DocumentAccessAuditRecord[]; total: number } {
    try {
      let filteredRecords = [...this.auditRecords];
      
      // Apply filters
      if (options?.userId) {
        filteredRecords = filteredRecords.filter(record => record.userId === options.userId);
      }
      
      if (options?.fileId) {
        filteredRecords = filteredRecords.filter(record => record.fileId === options.fileId);
      }
      
      if (options?.action) {
        filteredRecords = filteredRecords.filter(record => record.action === options.action);
      }
      
      if (options?.startDate) {
        filteredRecords = filteredRecords.filter(record => record.timestamp >= options.startDate);
      }
      
      if (options?.endDate) {
        filteredRecords = filteredRecords.filter(record => record.timestamp <= options.endDate);
      }
      
      // Sort by timestamp (newest first)
      filteredRecords.sort((a, b) => b.timestamp.getTime() - a.timestamp.getTime());
      
      // Apply pagination
      const page = options?.page || 1;
      const limit = Math.min(options?.limit || 50, 100); // Max 100 per page
      const startIndex = (page - 1) * limit;
      const paginatedRecords = filteredRecords.slice(startIndex, startIndex + limit);
      
      return {
        records: paginatedRecords,
        total: filteredRecords.length
      };
    } catch (error) {
      logger.error('Error retrieving audit records', {
        component: 'DocumentAccessAuditService',
        error: error instanceof Error ? error.message : String(error)
      });
      
      return {
        records: [],
        total: 0
      };
    }
  }

  /**
   * Check if a document is sensitive based on its name
   */
  private isSensitiveDocument(file: DriveFile): boolean {
    const fileName = (file.name || '').toLowerCase();
    return this.SENSITIVE_KEYWORDS.some(keyword => fileName.includes(keyword));
  }

  /**
   * Generate a unique ID for audit records
   */
  private generateId(): string {
    return `audit-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
  }

  /**
   * Export audit records to JSON
   */
  exportAuditRecords(format: 'json' | 'csv' = 'json'): string {
    try {
      if (format === 'json') {
        return JSON.stringify(this.auditRecords, null, 2);
      } else {
        // CSV format
        const headers = ['ID', 'File ID', 'File Name', 'User ID', 'User Name', 'Action', 'Timestamp', 'Session ID'];
        const csvRows = [headers.join(',')];
        
        this.auditRecords.forEach(record => {
          const row = [
            record.id,
            record.fileId,
            `"${record.fileName.replace(/"/g, '""')}"`,
            record.userId,
            `"${record.userName.replace(/"/g, '""')}"`,
            record.action,
            record.timestamp.toISOString(),
            record.sessionId
          ];
          csvRows.push(row.join(','));
        });
        
        return csvRows.join('\n');
      }
    } catch (error) {
      logger.error('Error exporting audit records', {
        component: 'DocumentAccessAuditService',
        format,
        error: error instanceof Error ? error.message : String(error)
      });
      
      throw error;
    }
  }

  /**
   * Clear audit records
   */
  clearAuditRecords(): void {
    this.auditRecords = [];
    logger.info('Audit records cleared', {
      component: 'DocumentAccessAuditService'
    });
  }

  /**
   * Get service statistics
   */
  getStats(): {
    totalRecords: number;
    dateRange: { oldest: Date | null; newest: Date | null };
    sensitiveDocumentAccesses: number;
  } {
    try {
      const totalRecords = this.auditRecords.length;
      
      let oldest: Date | null = null;
      let newest: Date | null = null;
      
      if (totalRecords > 0) {
        oldest = new Date(Math.min(...this.auditRecords.map(r => r.timestamp.getTime())));
        newest = new Date(Math.max(...this.auditRecords.map(r => r.timestamp.getTime())));
      }
      
      const sensitiveDocumentAccesses = this.auditRecords.filter(record => 
        this.isSensitiveDocument({ id: record.fileId, name: record.fileName } as DriveFile)
      ).length;
      
      return {
        totalRecords,
        dateRange: { oldest, newest },
        sensitiveDocumentAccesses
      };
    } catch (error) {
      logger.error('Error getting audit service stats', {
        component: 'DocumentAccessAuditService',
        error: error instanceof Error ? error.message : String(error)
      });
      
      return {
        totalRecords: 0,
        dateRange: { oldest: null, newest: null },
        sensitiveDocumentAccesses: 0
      };
    }
  }
}