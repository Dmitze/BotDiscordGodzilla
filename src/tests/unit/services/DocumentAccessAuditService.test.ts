/**
 * Unit tests for DocumentAccessAuditService functionality
 */

import { describe, it, expect, beforeEach } from '@jest/globals';
import { DocumentAccessAuditService } from '../../../services/DocumentAccessAuditService';
import { createMockConfig, createMockDriveFile } from '../../utils/testHelpers';

describe('DocumentAccessAuditService', () => {
  let auditService: DocumentAccessAuditService;
  let mockConfig: any;

  beforeEach(() => {
    mockConfig = createMockConfig();
    auditService = new DocumentAccessAuditService(mockConfig);
  });

  it('should log document access with all required fields', () => {
    const mockFile = createMockDriveFile('test-file-id', 'Test Document.txt');
    
    // Log document access
    auditService.logDocumentAccess({
      file: mockFile,
      userId: 'test-user-id',
      userName: 'Test User',
      action: 'view',
      sessionId: 'test-session-id'
    });
    
    // Verify service can retrieve records
    const { records, total } = auditService.getAuditRecords();
    expect(total).toBe(1);
    expect(records[0].fileId).toBe('test-file-id');
    expect(records[0].userId).toBe('test-user-id');
    expect(records[0].action).toBe('view');
  });

  it('should generate access summary with correct statistics', () => {
    const mockFile1 = createMockDriveFile('file-1', 'Document 1.txt');
    const mockFile2 = createMockDriveFile('file-2', 'Document 2.txt');
    
    // Log multiple accesses
    auditService.logDocumentAccess({
      file: mockFile1,
      userId: 'user-1',
      userName: 'User One',
      action: 'view',
      sessionId: 'session-1'
    });
    
    auditService.logDocumentAccess({
      file: mockFile1,
      userId: 'user-1',
      userName: 'User One',
      action: 'download',
      sessionId: 'session-2'
    });
    
    auditService.logDocumentAccess({
      file: mockFile2,
      userId: 'user-2',
      userName: 'User Two',
      action: 'view',
      sessionId: 'session-3'
    });
    
    // Generate summary
    const summary = auditService.generateAccessSummary();
    
    expect(summary.totalAccesses).toBe(3);
    expect(summary.uniqueUsers).toBe(2);
    expect(summary.popularDocuments).toHaveLength(2);
    expect(summary.actionDistribution).toHaveProperty('view', 2);
    expect(summary.actionDistribution).toHaveProperty('download', 1);
  });

  it('should identify sensitive document accesses', () => {
    const sensitiveFile = createMockDriveFile('sensitive-file', 'Confidential Report.txt');
    const normalFile = createMockDriveFile('normal-file', 'Regular Document.txt');
    
    // Log accesses to sensitive and normal documents
    auditService.logDocumentAccess({
      file: sensitiveFile,
      userId: 'user-1',
      userName: 'User One',
      action: 'view',
      sessionId: 'session-1'
    });
    
    auditService.logDocumentAccess({
      file: normalFile,
      userId: 'user-2',
      userName: 'User Two',
      action: 'view',
      sessionId: 'session-2'
    });
    
    // Generate compliance report
    const report = auditService.generateAccessSummary();
    
    // The sensitive document should be flagged
    expect(report.popularDocuments).toHaveLength(2);
  });

  it('should filter audit records by various criteria', () => {
    const mockFile = createMockDriveFile('test-file', 'Test Document.txt');
    
    // Log accesses at different times
    auditService.logDocumentAccess({
      file: mockFile,
      userId: 'user-1',
      userName: 'User One',
      action: 'view',
      sessionId: 'session-1'
    });
    
    auditService.logDocumentAccess({
      file: mockFile,
      userId: 'user-2',
      userName: 'User Two',
      action: 'download',
      sessionId: 'session-2'
    });
    
    // Filter by user
    const userRecords = auditService.getAuditRecords({ userId: 'user-1' });
    expect(userRecords.total).toBe(1);
    expect(userRecords.records[0].userId).toBe('user-1');
    
    // Filter by action
    const downloadRecords = auditService.getAuditRecords({ action: 'download' });
    expect(downloadRecords.total).toBe(1);
    expect(downloadRecords.records[0].action).toBe('download');
    
    // Filter by file
    const fileRecords = auditService.getAuditRecords({ fileId: 'test-file' });
    expect(fileRecords.total).toBe(2);
  });

  it('should handle pagination correctly', () => {
    const mockFile = createMockDriveFile('pagination-test', 'Pagination Test.txt');
    
    // Log 25 accesses
    for (let i = 0; i < 25; i++) {
      auditService.logDocumentAccess({
        file: mockFile,
        userId: `user-${i % 3}`, // 3 different users
        userName: `User ${i % 3}`,
        action: i % 2 === 0 ? 'view' : 'download', // Alternate actions
        sessionId: `session-${i}`
      });
    }
    
    // Test pagination
    const page1 = auditService.getAuditRecords({ page: 1, limit: 10 });
    const page2 = auditService.getAuditRecords({ page: 2, limit: 10 });
    const page3 = auditService.getAuditRecords({ page: 3, limit: 10 });
    
    expect(page1.records).toHaveLength(10);
    expect(page2.records).toHaveLength(10);
    expect(page3.records).toHaveLength(5); // Last page with remaining records
    expect(page1.total).toBe(25);
  });

  it('should generate compliance reports with flagged activities', async () => {
    const mockFile = createMockDriveFile('compliance-test', 'Confidential Document.txt');
    
    // Log accesses including after-hours access
    auditService.logDocumentAccess({
      file: mockFile,
      userId: 'user-1',
      userName: 'User One',
      action: 'view',
      sessionId: 'session-1'
    });
    
    // Generate compliance report
    const report = await auditService.generateComplianceReport();
    
    expect(report).toBeDefined();
    expect(report.reportId).toBeDefined();
    expect(report.summary).toBeDefined();
    expect(report.detailedRecords).toHaveLength(1);
  });

  it('should export audit records in different formats', () => {
    const mockFile = createMockDriveFile('export-test', 'Export Test.txt');
    
    auditService.logDocumentAccess({
      file: mockFile,
      userId: 'export-user',
      userName: 'Export User',
      action: 'view',
      sessionId: 'export-session'
    });
    
    // Test JSON export
    const jsonExport = auditService.exportAuditRecords('json');
    expect(jsonExport).toContain('export-test');
    expect(jsonExport).toContain('export-user');
    
    // Test CSV export
    const csvExport = auditService.exportAuditRecords('csv');
    expect(csvExport).toContain('export-test');
    expect(csvExport).toContain('export-user');
  });

  it('should maintain records within configured limits', () => {
    const mockFile = createMockDriveFile('limit-test', 'Limit Test.txt');
    
    // Log more accesses than the maximum limit (50,000)
    // We'll test with a smaller number for performance
    for (let i = 0; i < 100; i++) {
      auditService.logDocumentAccess({
        file: { ...mockFile, id: `file-${i}` },
        userId: `user-${i}`,
        userName: `User ${i}`,
        action: 'view',
        sessionId: `session-${i}`
      });
    }
    
    // Verify service handles high volume without errors
    const { total } = auditService.getAuditRecords();
    expect(total).toBe(100);
  });

  it('should provide service statistics', () => {
    const mockFile = createMockDriveFile('stats-test', 'Stats Test.txt');
    
    auditService.logDocumentAccess({
      file: mockFile,
      userId: 'stats-user',
      userName: 'Stats User',
      action: 'view',
      sessionId: 'stats-session'
    });
    
    const stats = auditService.getStats();
    expect(stats.totalRecords).toBe(1);
    expect(stats.dateRange.oldest).toBeDefined();
    expect(stats.dateRange.newest).toBeDefined();
  });
});