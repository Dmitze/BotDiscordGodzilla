/**
 * Enhanced unit tests for DocumentAccessAuditService with comprehensive document interaction logging
 */

import { describe, it, expect, beforeEach } from '@jest/globals';
import { DocumentAccessAuditService } from '../../../services/DocumentAccessAuditService';
import { createMockConfig, createMockDriveFile } from '../../utils/testHelpers';

describe('DocumentAccessAuditService - Enhanced Features', () => {
  let auditService: DocumentAccessAuditService;
  let mockConfig: any;

  beforeEach(() => {
    mockConfig = createMockConfig();
    auditService = new DocumentAccessAuditService(mockConfig);
  });

  it('should log comprehensive document interactions with all enhanced fields', () => {
    const mockFile = createMockDriveFile('enhanced-test-file', 'Enhanced Test Document.txt');
    
    // Log document interaction with enhanced fields
    auditService.logDocumentInteraction({
      file: mockFile,
      userId: 'enhanced-user-id',
      userName: 'Enhanced Test User',
      action: 'view',
      sessionId: 'enhanced-session-id',
      contentAccessPattern: {
        accessedSections: [{ start: 0, end: 100 }],
        totalAccessedBytes: 1000,
        accessDuration: 5000,
        scrollPattern: 'top-to-bottom',
        searchTerms: ['test', 'document'],
        highlightedText: ['important', 'section']
      },
      documentContext: {
        documentType: 'document',
        documentSize: 2048,
        wordCount: 100,
        language: 'en',
        tags: ['test', 'document']
      },
      securityContext: {
        encryptionStatus: 'encrypted',
        accessLevel: 'internal',
        twoFactorVerified: true,
        sessionSecurityLevel: 'enhanced'
      },
      performanceMetrics: {
        responseTime: 100,
        processingTime: 50,
        bandwidthUsed: 2048,
        cacheHit: true
      },
      metadata: {
        customField: 'customValue',
        version: '1.0'
      }
    });
    
    // Verify service can retrieve records with enhanced fields
    const { records, total } = auditService.getAuditRecords();
    expect(total).toBe(1);
    expect(records[0].fileId).toBe('enhanced-test-file');
    expect(records[0].userId).toBe('enhanced-user-id');
    expect(records[0].action).toBe('view');
    
    // Check enhanced fields
    expect(records[0].contentAccessPattern).toBeDefined();
    expect(records[0].contentAccessPattern?.accessedSections).toHaveLength(1);
    expect(records[0].contentAccessPattern?.totalAccessedBytes).toBe(1000);
    expect(records[0].contentAccessPattern?.accessDuration).toBe(5000);
    expect(records[0].contentAccessPattern?.scrollPattern).toBe('top-to-bottom');
    expect(records[0].contentAccessPattern?.searchTerms).toContain('test');
    
    expect(records[0].documentContext).toBeDefined();
    expect(records[0].documentContext?.documentType).toBe('document');
    expect(records[0].documentContext?.documentSize).toBe(2048);
    expect(records[0].documentContext?.wordCount).toBe(100);
    expect(records[0].documentContext?.language).toBe('en');
    expect(records[0].documentContext?.tags).toContain('test');
    
    expect(records[0].securityContext).toBeDefined();
    expect(records[0].securityContext?.encryptionStatus).toBe('encrypted');
    expect(records[0].securityContext?.accessLevel).toBe('internal');
    expect(records[0].securityContext?.twoFactorVerified).toBe(true);
    expect(records[0].securityContext?.sessionSecurityLevel).toBe('enhanced');
    
    expect(records[0].performanceMetrics).toBeDefined();
    expect(records[0].performanceMetrics?.responseTime).toBe(100);
    expect(records[0].performanceMetrics?.processingTime).toBe(50);
    expect(records[0].performanceMetrics?.bandwidthUsed).toBe(2048);
    expect(records[0].performanceMetrics?.cacheHit).toBe(true);
    
    expect(records[0].metadata).toBeDefined();
    expect(records[0].metadata?.customField).toBe('customValue');
    expect(records[0].metadata?.version).toBe('1.0');
  });

  it('should generate enhanced access summary with comprehensive statistics', () => {
    const mockFile1 = createMockDriveFile('file-1', 'Document 1.txt');
    const mockFile2 = createMockDriveFile('file-2', 'Document 2.txt');
    
    // Log multiple accesses with enhanced fields
    auditService.logDocumentInteraction({
      file: mockFile1,
      userId: 'user-1',
      userName: 'User One',
      action: 'view',
      sessionId: 'session-1',
      contentAccessPattern: {
        accessedSections: [{ start: 0, end: 100 }],
        totalAccessedBytes: 1000,
        accessDuration: 5000,
        scrollPattern: 'top-to-bottom',
        searchTerms: ['test']
      },
      securityContext: {
        encryptionStatus: 'encrypted',
        accessLevel: 'internal',
        twoFactorVerified: true,
        sessionSecurityLevel: 'standard'
      },
      performanceMetrics: {
        responseTime: 100,
        processingTime: 50
      }
    });
    
    auditService.logDocumentInteraction({
      file: mockFile1,
      userId: 'user-1',
      userName: 'User One',
      action: 'download',
      sessionId: 'session-2',
      contentAccessPattern: {
        accessedSections: [{ start: 0, end: 200 }],
        totalAccessedBytes: 2000,
        accessDuration: 10000,
        scrollPattern: 'top-to-bottom'
      },
      securityContext: {
        encryptionStatus: 'unencrypted',
        accessLevel: 'internal',
        twoFactorVerified: false,
        sessionSecurityLevel: 'standard'
      }
    });
    
    auditService.logDocumentInteraction({
      file: mockFile2,
      userId: 'user-2',
      userName: 'User Two',
      action: 'view',
      sessionId: 'session-3',
      contentAccessPattern: {
        accessedSections: [{ start: 0, end: 150 }],
        totalAccessedBytes: 1500,
        accessDuration: 7500,
        scrollPattern: 'random',
        searchTerms: ['document']
      },
      securityContext: {
        encryptionStatus: 'encrypted',
        accessLevel: 'internal',
        twoFactorVerified: true,
        sessionSecurityLevel: 'standard'
      },
      performanceMetrics: {
        responseTime: 150,
        processingTime: 75
      }
    });
    
    // Generate enhanced summary
    const summary = auditService.generateAccessSummary();
    
    expect(summary.totalAccesses).toBe(3);
    expect(summary.uniqueUsers).toBe(2);
    expect(summary.popularDocuments).toHaveLength(2);
    expect(summary.actionDistribution).toHaveProperty('view', 2);
    expect(summary.actionDistribution).toHaveProperty('download', 1);
    
    // Check enhanced content access patterns
    expect(summary.contentAccessPatterns).toBeDefined();
    expect(summary.contentAccessPatterns.mostCommonScrollPattern).toBe('top-to-bottom');
    expect(summary.contentAccessPatterns.averageAccessDuration).toBeGreaterThan(0);
    expect(summary.contentAccessPatterns.commonSearchTerms).toHaveLength(2);
    
    // Check enhanced security metrics
    expect(summary.securityMetrics).toBeDefined();
    expect(summary.securityMetrics.encryptedAccesses).toBe(2);
    expect(summary.securityMetrics.unencryptedAccesses).toBe(1);
    expect(summary.securityMetrics.twoFactorVerifiedAccesses).toBe(2);
    
    // Check enhanced performance metrics
    expect(summary.performanceMetrics).toBeDefined();
    expect(summary.performanceMetrics.averageResponseTime).toBeGreaterThan(0);
    expect(summary.performanceMetrics.averageProcessingTime).toBeGreaterThan(0);
    expect(summary.performanceMetrics.totalBandwidthUsed).toBe(0); // Not provided in test data
  });

  it('should detect and log security incidents', () => {
    const mockFile = createMockDriveFile('sensitive-file', 'Confidential Report.txt');
    
    // Log access that should trigger a security incident (large download)
    auditService.logDocumentInteraction({
      file: mockFile,
      userId: 'user-1',
      userName: 'User One',
      action: 'download',
      sessionId: 'session-1',
      contentAccessPattern: {
        accessedSections: [{ start: 0, end: 1000 }],
        totalAccessedBytes: 150 * 1024 * 1024, // 150MB - should trigger incident
        accessDuration: 5000
      },
      securityContext: {
        encryptionStatus: 'unencrypted',
        accessLevel: 'confidential',
        twoFactorVerified: false,
        sessionSecurityLevel: 'standard'
      }
    });
    
    // Check that security incidents were logged
    const stats = auditService.getStats();
    expect(stats.securityIncidents).toBeGreaterThan(0);
  });

  it('should generate enhanced compliance reports with risk assessment', async () => {
    const mockFile = createMockDriveFile('compliance-test', 'Confidential Document.txt');
    
    // Log accesses with various patterns
    auditService.logDocumentInteraction({
      file: mockFile,
      userId: 'user-1',
      userName: 'User One',
      action: 'view',
      sessionId: 'session-1',
      securityContext: {
        encryptionStatus: 'unencrypted',
        accessLevel: 'confidential',
        twoFactorVerified: false,
        sessionSecurityLevel: 'standard'
      }
    });
    
    // Generate enhanced compliance report
    const report = await auditService.generateComplianceReport();
    
    expect(report).toBeDefined();
    expect(report.reportId).toBeDefined();
    expect(report.summary).toBeDefined();
    expect(report.detailedRecords).toHaveLength(1);
    expect(report.accessTrends).toBeDefined();
    expect(report.riskAssessment).toBeDefined();
    expect(report.riskAssessment.complianceScore).toBeDefined();
  });

  it('should maintain backward compatibility with old logDocumentAccess method', () => {
    const mockFile = createMockDriveFile('compatibility-test', 'Compatibility Test.txt');
    
    // Use old method
    auditService.logDocumentAccess({
      file: mockFile,
      userId: 'compat-user',
      userName: 'Compatibility User',
      action: 'view',
      sessionId: 'compat-session'
    });
    
    // Verify record was logged
    const { records, total } = auditService.getAuditRecords();
    expect(total).toBe(1);
    expect(records[0].fileId).toBe('compatibility-test');
    expect(records[0].userId).toBe('compat-user');
    expect(records[0].action).toBe('view');
  });
});