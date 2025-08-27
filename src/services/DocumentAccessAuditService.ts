import { BaseService } from '@/core/BaseService';
import type { BotConfig } from '@/types';
import type { DriveFile } from '@/types/drive';
import logger from '@/utils/logger';

// Enhanced interfaces for comprehensive document interaction logging
export interface DocumentInteractionEvent {
  id: string;
  fileId: string;
  fileName: string;
  userId: string;
  userName: string;
  action: 'view' | 'download' | 'search' | 'analyze' | 'export' | 'compare' | 'summarize' | 'delete' | 'edit' | 'share' | 'comment' | 'annotate' | 'copy' | 'move' | 'rename';
  timestamp: Date;
  ipAddress?: string;
  userAgent?: string;
  sessionId: string;
  // Enhanced fields for comprehensive logging
  contentAccessPattern?: ContentAccessPattern;
  modificationDetails?: ModificationDetails;
  documentContext?: DocumentContext;
  securityContext?: SecurityContext;
  performanceMetrics?: PerformanceMetrics;
  metadata?: Record<string, any>;
}

export interface ContentAccessPattern {
  accessedSections: { start: number; end: number }[];
  totalAccessedBytes: number;
  accessDuration?: number; // in milliseconds
  scrollPattern?: 'top-to-bottom' | 'bottom-to-top' | 'random' | 'focused';
  searchTerms?: string[];
  highlightedText?: string[];
}

export interface ModificationDetails {
  modificationType: 'insert' | 'delete' | 'replace' | 'format' | 'move';
  modifiedSections: { start: number; end: number }[];
  originalContent?: string;
  newContent?: string;
  modificationSize?: number; // in bytes
}

export interface DocumentContext {
  documentType: 'document' | 'spreadsheet' | 'presentation' | 'pdf' | 'image' | 'other';
  documentSize?: number; // in bytes
  pageCount?: number;
  wordCount?: number;
  language?: string;
  tags?: string[];
  folderPath?: string[];
}

export interface SecurityContext {
  encryptionStatus: 'encrypted' | 'unencrypted' | 'partially-encrypted';
  accessLevel: 'public' | 'internal' | 'confidential' | 'restricted';
  twoFactorVerified: boolean;
  sessionSecurityLevel: 'standard' | 'enhanced' | 'admin';
}

export interface PerformanceMetrics {
  responseTime?: number; // in milliseconds
  processingTime?: number; // in milliseconds
  bandwidthUsed?: number; // in bytes
  cacheHit?: boolean;
}

export interface DocumentAccessAuditRecord extends DocumentInteractionEvent {
  // Inherit all fields from DocumentInteractionEvent
}

export interface AccessSummary {
  totalAccesses: number;
  uniqueUsers: number;
  popularDocuments: { fileId: string; fileName: string; accessCount: number }[];
  actionDistribution: Record<string, number>;
  timeSeries: { date: string; count: number }[];
  // Enhanced summary fields
  contentAccessPatterns: {
    mostCommonScrollPattern: string;
    averageAccessDuration: number;
    commonSearchTerms: { term: string; count: number }[];
  };
  securityMetrics: {
    encryptedAccesses: number;
    unencryptedAccesses: number;
    twoFactorVerifiedAccesses: number;
  };
  performanceMetrics: {
    averageResponseTime: number;
    averageProcessingTime: number;
    totalBandwidthUsed: number;
  };
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
  // Enhanced compliance fields
  securityIncidents: SecurityIncident[];
  accessTrends: AccessTrend[];
  riskAssessment: RiskAssessment;
}

export interface SecurityIncident {
  id: string;
  type: 'unauthorized_access' | 'data_exfiltration' | 'suspicious_activity' | 'policy_violation';
  severity: 'low' | 'medium' | 'high' | 'critical';
  timestamp: Date;
  userId?: string;
  fileId?: string;
  description: string;
  evidence: any[];
  resolved: boolean;
  resolutionNotes?: string;
}

export interface AccessTrend {
  period: string; // e.g., "2023-01", "2023-Q1"
  totalAccesses: number;
  uniqueUsers: number;
  sensitiveDocumentAccesses: number;
  flaggedActivities: number;
  growthRate: number; // percentage
}

export interface RiskAssessment {
  overallRiskLevel: 'low' | 'medium' | 'high' | 'critical';
  riskFactors: {
    factor: string;
    level: 'low' | 'medium' | 'high' | 'critical';
    description: string;
    recommendations: string[];
  }[];
  complianceScore: number; // 0-100
}

export class DocumentAccessAuditService extends BaseService {
  private auditRecords: DocumentAccessAuditRecord[] = [];
  private securityIncidents: SecurityIncident[] = [];
  private readonly MAX_RECORDS = 50000; // Keep last 50,000 records in memory
  private readonly SENSITIVE_KEYWORDS = ['confidential', 'secret', 'private', 'internal', 'restricted', 'classified', 'proprietary'];
  
  constructor(config: BotConfig) {
    super('DocumentAccessAuditService', config);
  }

  /**
   * Log a comprehensive document interaction event
   */
  logDocumentInteraction(args: {
    file: DriveFile;
    userId: string;
    userName: string;
    action: DocumentInteractionEvent['action'];
    sessionId: string;
    ipAddress?: string;
    userAgent?: string;
    contentAccessPattern?: ContentAccessPattern;
    modificationDetails?: ModificationDetails;
    documentContext?: DocumentContext;
    securityContext?: SecurityContext;
    performanceMetrics?: PerformanceMetrics;
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
        contentAccessPattern: args.contentAccessPattern,
        modificationDetails: args.modificationDetails,
        documentContext: args.documentContext,
        securityContext: args.securityContext,
        performanceMetrics: args.performanceMetrics,
        metadata: args.metadata
      };

      // Add record to audit log
      this.auditRecords.push(record);

      // Maintain record limit
      if (this.auditRecords.length > this.MAX_RECORDS) {
        this.auditRecords = this.auditRecords.slice(-this.MAX_RECORDS);
      }

      // Log the access
      logger.info('Document interaction logged', {
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

      // Check for potential security incidents
      this.checkForSecurityIncidents(record);
    } catch (error) {
      logger.error('Error logging document interaction', {
        component: 'DocumentAccessAuditService',
        fileId: args.file?.id,
        userId: args.userId,
        error: error instanceof Error ? error.message : String(error)
      });
    }
  }

  /**
   * Log a document access event (backward compatibility)
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
    // Delegate to the new comprehensive logging method
    this.logDocumentInteraction(args);
  }

  /**
   * Check for potential security incidents based on access patterns
   */
  private checkForSecurityIncidents(record: DocumentAccessAuditRecord): void {
    const incidents: SecurityIncident[] = [];

    // Check for bulk downloads
    if (record.action === 'download' && record.contentAccessPattern?.totalAccessedBytes) {
      if (record.contentAccessPattern.totalAccessedBytes > 100 * 1024 * 1024) { // 100MB
        incidents.push({
          id: this.generateId(),
          type: 'data_exfiltration',
          severity: 'high',
          timestamp: record.timestamp,
          userId: record.userId,
          fileId: record.fileId,
          description: `Large download detected: ${record.contentAccessPattern.totalAccessedBytes} bytes`,
          evidence: [record],
          resolved: false
        });
      }
    }

    // Check for access outside business hours
    const hour = record.timestamp.getHours();
    if (hour < 6 || hour >= 22) {
      incidents.push({
        id: this.generateId(),
        type: 'suspicious_activity',
        severity: 'medium',
        timestamp: record.timestamp,
        userId: record.userId,
        fileId: record.fileId,
        description: `Access outside business hours: ${hour}:00`,
        evidence: [record],
        resolved: false
      });
    }

    // Check for access to sensitive documents without 2FA
    if (this.isSensitiveDocument({ id: record.fileId, name: record.fileName } as DriveFile) && 
        record.securityContext && !record.securityContext.twoFactorVerified) {
      incidents.push({
        id: this.generateId(),
        type: 'policy_violation',
        severity: 'high',
        timestamp: record.timestamp,
        userId: record.userId,
        fileId: record.fileId,
        description: 'Access to sensitive document without two-factor authentication',
        evidence: [record],
        resolved: false
      });
    }

    // Add incidents to the list
    this.securityIncidents.push(...incidents);

    // Log security incidents
    incidents.forEach(incident => {
      logger.warn('Security incident detected', {
        component: 'DocumentAccessAuditService',
        incidentId: incident.id,
        type: incident.type,
        severity: incident.severity,
        userId: incident.userId,
        fileId: incident.fileId,
        description: incident.description
      });
    });
  }

  /**
   * Generate an enhanced access summary report
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

      // Enhanced content access patterns
      const scrollPatterns: Record<string, number> = {};
      let totalAccessDuration = 0;
      let accessWithDurationCount = 0;
      const searchTermsMap: Record<string, number> = {};
      
      filteredRecords.forEach(record => {
        if (record.contentAccessPattern) {
          // Scroll patterns
          if (record.contentAccessPattern.scrollPattern) {
            scrollPatterns[record.contentAccessPattern.scrollPattern] = 
              (scrollPatterns[record.contentAccessPattern.scrollPattern] || 0) + 1;
          }
          
          // Access duration
          if (record.contentAccessPattern.accessDuration) {
            totalAccessDuration += record.contentAccessPattern.accessDuration;
            accessWithDurationCount++;
          }
          
          // Search terms
          if (record.contentAccessPattern.searchTerms) {
            record.contentAccessPattern.searchTerms.forEach(term => {
              searchTermsMap[term] = (searchTermsMap[term] || 0) + 1;
            });
          }
        }
      });
      
      const mostCommonScrollPattern = Object.entries(scrollPatterns)
        .sort((a, b) => b[1] - a[1])[0]?.[0] || 'unknown';
      
      const averageAccessDuration = accessWithDurationCount > 0 ? 
        totalAccessDuration / accessWithDurationCount : 0;
      
      const commonSearchTerms = Object.entries(searchTermsMap)
        .map(([term, count]) => ({ term, count }))
        .sort((a, b) => b.count - a.count)
        .slice(0, 10);

      // Enhanced security metrics
      let encryptedAccesses = 0;
      let unencryptedAccesses = 0;
      let twoFactorVerifiedAccesses = 0;
      
      filteredRecords.forEach(record => {
        if (record.securityContext) {
          if (record.securityContext.encryptionStatus === 'encrypted') {
            encryptedAccesses++;
          } else {
            unencryptedAccesses++;
          }
          
          if (record.securityContext.twoFactorVerified) {
            twoFactorVerifiedAccesses++;
          }
        }
      });

      // Enhanced performance metrics
      let totalResponseTime = 0;
      let totalProcessingTime = 0;
      let totalBandwidth = 0;
      let responseTimeCount = 0;
      let processingTimeCount = 0;
      
      filteredRecords.forEach(record => {
        if (record.performanceMetrics) {
          if (record.performanceMetrics.responseTime !== undefined) {
            totalResponseTime += record.performanceMetrics.responseTime;
            responseTimeCount++;
          }
          
          if (record.performanceMetrics.processingTime !== undefined) {
            totalProcessingTime += record.performanceMetrics.processingTime;
            processingTimeCount++;
          }
          
          if (record.performanceMetrics.bandwidthUsed !== undefined) {
            totalBandwidth += record.performanceMetrics.bandwidthUsed;
          }
        }
      });
      
      const averageResponseTime = responseTimeCount > 0 ? totalResponseTime / responseTimeCount : 0;
      const averageProcessingTime = processingTimeCount > 0 ? totalProcessingTime / processingTimeCount : 0;

      return {
        totalAccesses,
        uniqueUsers,
        popularDocuments,
        actionDistribution,
        timeSeries,
        contentAccessPatterns: {
          mostCommonScrollPattern,
          averageAccessDuration,
          commonSearchTerms
        },
        securityMetrics: {
          encryptedAccesses,
          unencryptedAccesses,
          twoFactorVerifiedAccesses
        },
        performanceMetrics: {
          averageResponseTime,
          averageProcessingTime,
          totalBandwidthUsed: totalBandwidth
        }
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
        timeSeries: [],
        contentAccessPatterns: {
          mostCommonScrollPattern: 'unknown',
          averageAccessDuration: 0,
          commonSearchTerms: []
        },
        securityMetrics: {
          encryptedAccesses: 0,
          unencryptedAccesses: 0,
          twoFactorVerifiedAccesses: 0
        },
        performanceMetrics: {
          averageResponseTime: 0,
          averageProcessingTime: 0,
          totalBandwidthUsed: 0
        }
      };
    }
  }

  /**
   * Generate an enhanced compliance report
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

      // Get security incidents for the period
      const periodIncidents = this.securityIncidents.filter(
        incident => incident.timestamp >= startDate && incident.timestamp <= endDate
      );

      // Generate access trends
      const accessTrends = this.generateAccessTrends(filteredRecords, startDate, endDate);

      // Generate risk assessment
      const riskAssessment = this.generateRiskAssessment(filteredRecords, periodIncidents);

      const report: ComplianceReport = {
        reportId: this.generateId(),
        generatedAt: new Date(),
        periodStart: startDate,
        periodEnd: endDate,
        summary,
        detailedRecords: filteredRecords,
        sensitiveDocumentAccesses,
        flaggedActivities: uniqueFlaggedActivities,
        securityIncidents: periodIncidents,
        accessTrends,
        riskAssessment
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
   * Generate access trends for compliance reporting
   */
  private generateAccessTrends(records: DocumentAccessAuditRecord[], startDate: Date, endDate: Date): AccessTrend[] {
    // Group records by month
    const monthlyData: Record<string, {
      totalAccesses: number;
      uniqueUsers: Set<string>;
      sensitiveAccesses: number;
      flaggedActivities: number;
    }> = {};
    
    records.forEach(record => {
      const month = record.timestamp.toISOString().substring(0, 7); // YYYY-MM
      if (!monthlyData[month]) {
        monthlyData[month] = {
          totalAccesses: 0,
          uniqueUsers: new Set(),
          sensitiveAccesses: 0,
          flaggedActivities: 0
        };
      }
      
      monthlyData[month].totalAccesses++;
      monthlyData[month].uniqueUsers.add(record.userId);
      
      if (this.isSensitiveDocument({ id: record.fileId, name: record.fileName } as DriveFile)) {
        monthlyData[month].sensitiveAccesses++;
      }
    });
    
    // Convert to AccessTrend objects
    const trends: AccessTrend[] = [];
    const months = Object.keys(monthlyData).sort();
    
    months.forEach((month, index) => {
      const data = monthlyData[month];
      const previousMonth = index > 0 ? months[index - 1] : null;
      const previousData = previousMonth ? monthlyData[previousMonth] : null;
      
      let growthRate = 0;
      if (previousData) {
        const difference = data.totalAccesses - previousData.totalAccesses;
        growthRate = previousData.totalAccesses > 0 ? (difference / previousData.totalAccesses) * 100 : 0;
      }
      
      trends.push({
        period: month,
        totalAccesses: data.totalAccesses,
        uniqueUsers: data.uniqueUsers.size,
        sensitiveDocumentAccesses: data.sensitiveAccesses,
        flaggedActivities: data.flaggedActivities || 0,
        growthRate
      });
    });
    
    return trends;
  }

  /**
   * Generate risk assessment for compliance reporting
   */
  private generateRiskAssessment(records: DocumentAccessAuditRecord[], incidents: SecurityIncident[]): RiskAssessment {
    const riskFactors = [];
    
    // Check for high number of security incidents
    if (incidents.length > 10) {
      riskFactors.push({
        factor: 'Security Incidents',
        level: 'high' as const,
        description: `High number of security incidents detected (${incidents.length})`,
        recommendations: [
          'Review and strengthen access controls',
          'Implement additional monitoring',
          'Conduct security training for users'
        ]
      });
    } else if (incidents.length > 0) {
      riskFactors.push({
        factor: 'Security Incidents',
        level: 'medium' as const,
        description: `Some security incidents detected (${incidents.length})`,
        recommendations: [
          'Monitor for similar activities',
          'Review access logs regularly'
        ]
      });
    }
    
    // Check for unencrypted accesses
    const unencryptedAccesses = records.filter(record => 
      record.securityContext?.encryptionStatus !== 'encrypted'
    ).length;
    
    if (unencryptedAccesses > records.length * 0.5) { // More than 50% unencrypted
      riskFactors.push({
        factor: 'Encryption',
        level: 'high' as const,
        description: 'High percentage of unencrypted document accesses',
        recommendations: [
          'Enforce encryption for all document accesses',
          'Audit encryption policies'
        ]
      });
    }
    
    // Check for accesses without 2FA
    const non2FAAccesses = records.filter(record => 
      !record.securityContext?.twoFactorVerified
    ).length;
    
    if (non2FAAccesses > records.length * 0.3) { // More than 30% without 2FA
      riskFactors.push({
        factor: 'Authentication',
        level: 'medium' as const,
        description: 'Significant number of accesses without two-factor authentication',
        recommendations: [
          'Enforce two-factor authentication for all users',
          'Implement step-up authentication for sensitive documents'
        ]
      });
    }
    
    // Determine overall risk level
    const highRiskFactors = riskFactors.filter(f => f.level === 'high' || f.level === 'critical').length;
    const mediumRiskFactors = riskFactors.filter(f => f.level === 'medium').length;
    
    let overallRiskLevel: 'low' | 'medium' | 'high' | 'critical' = 'low';
    if (highRiskFactors > 2) {
      overallRiskLevel = 'critical';
    } else if (highRiskFactors > 0 || mediumRiskFactors > 3) {
      overallRiskLevel = 'high';
    } else if (mediumRiskFactors > 0) {
      overallRiskLevel = 'medium';
    }
    
    // Calculate compliance score (simplified)
    const maxScore = 100;
    const penaltyPerHighRisk = 25;
    const penaltyPerMediumRisk = 10;
    const complianceScore = Math.max(0, maxScore - (highRiskFactors * penaltyPerHighRisk) - (mediumRiskFactors * penaltyPerMediumRisk));
    
    return {
      overallRiskLevel,
      riskFactors,
      complianceScore
    };
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
    securityIncidents: number;
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
        sensitiveDocumentAccesses,
        securityIncidents: this.securityIncidents.length
      };
    } catch (error) {
      logger.error('Error getting audit service stats', {
        component: 'DocumentAccessAuditService',
        error: error instanceof Error ? error.message : String(error)
      });
      
      return {
        totalRecords: 0,
        dateRange: { oldest: null, newest: null },
        sensitiveDocumentAccesses: 0,
        securityIncidents: 0
      };
    }
  }
}