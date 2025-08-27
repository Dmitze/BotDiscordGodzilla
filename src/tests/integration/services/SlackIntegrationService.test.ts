/**
 * Integration test for Slack Integration with Document Audit Service
 */

import { describe, it, expect, beforeEach, jest } from '@jest/globals';
import { DocumentAccessAuditService } from '../../../services/DocumentAccessAuditService';
import { SlackIntegrationService } from '../../../services/SlackIntegrationService';
import { createMockConfig, createMockDriveFile } from '../../utils/testHelpers';

// Mock the Slack WebClient
jest.mock('@slack/web-api', () => {
  return {
    WebClient: jest.fn().mockImplementation(() => {
      return {
        chat: {
          postMessage: jest.fn().mockResolvedValue({ ok: true })
        }
      };
    })
  };
});

describe('Slack Integration with Document Audit', () => {
  let auditService: DocumentAccessAuditService;
  let slackService: SlackIntegrationService;
  let mockConfig: any;

  beforeEach(() => {
    mockConfig = createMockConfig();
    // Add Slack integration config
    mockConfig.integrations = {
      slack: {
        enabled: true,
        botToken: 'xoxb-test-token',
        channelId: 'C1234567890'
      }
    };
    
    auditService = new DocumentAccessAuditService(mockConfig);
    slackService = new SlackIntegrationService(mockConfig);
  });

  it('should send Slack notification when document is accessed', async () => {
    const mockFile = createMockDriveFile('test-file', 'Confidential Document.txt');
    
    // Log document access
    auditService.logDocumentInteraction({
      file: mockFile,
      userId: 'user-123',
      userName: 'Test User',
      action: 'view',
      sessionId: 'session-abc',
      contentAccessPattern: {
        accessedSections: [{ start: 0, end: 100 }],
        totalAccessedBytes: 1000,
        accessDuration: 5000,
        scrollPattern: 'top-to-bottom'
      },
      securityContext: {
        encryptionStatus: 'encrypted',
        accessLevel: 'confidential',
        twoFactorVerified: true,
        sessionSecurityLevel: 'enhanced'
      }
    });
    
    // Send notification to Slack
    const event = {
      fileId: mockFile.id,
      fileName: mockFile.name || 'Untitled',
      userId: 'user-123',
      userName: 'Test User',
      action: 'view' as const,
      timestamp: new Date()
    };
    
    const result = await slackService.sendDocumentNotification(event);
    
    expect(result).toBe(true);
    
    // Verify audit records
    const { records } = auditService.getAuditRecords();
    expect(records.length).toBe(1);
    expect(records[0].fileId).toBe('test-file');
    expect(records[0].action).toBe('view');
  });

  it('should detect sensitive documents and send appropriate Slack notification', async () => {
    const mockFile = createMockDriveFile('confidential-file', 'Top Secret Report.pdf');
    
    // Log document access
    auditService.logDocumentInteraction({
      file: mockFile,
      userId: 'admin-456',
      userName: 'Admin User',
      action: 'download',
      sessionId: 'session-def',
      contentAccessPattern: {
        accessedSections: [{ start: 0, end: 5000 }],
        totalAccessedBytes: 5000000, // 5MB
        accessDuration: 10000,
        scrollPattern: 'random'
      },
      securityContext: {
        encryptionStatus: 'encrypted',
        accessLevel: 'confidential',
        twoFactorVerified: true,
        sessionSecurityLevel: 'admin'
      }
    });
    
    // Send notification to Slack
    const event = {
      fileId: mockFile.id,
      fileName: mockFile.name || 'Untitled',
      userId: 'admin-456',
      userName: 'Admin User',
      action: 'download' as const,
      timestamp: new Date()
    };
    
    const result = await slackService.sendDocumentNotification(event);
    
    expect(result).toBe(true);
    
    // Verify the notification was sent for a sensitive document
    const stats = slackService.getStats();
    expect(stats.totalNotificationsSent).toBe(1);
  });

  it('should handle batch notifications for multiple document events', async () => {
    const mockFile1 = createMockDriveFile('file-1', 'Document 1.txt');
    const mockFile2 = createMockDriveFile('file-2', 'Document 2.txt');
    const mockFile3 = createMockDriveFile('file-3', 'Confidential Document.txt');
    
    // Log multiple document accesses
    auditService.logDocumentInteraction({
      file: mockFile1,
      userId: 'user-1',
      userName: 'User One',
      action: 'view',
      sessionId: 'session-1'
    });
    
    auditService.logDocumentInteraction({
      file: mockFile2,
      userId: 'user-2',
      userName: 'User Two',
      action: 'edit',
      sessionId: 'session-2'
    });
    
    auditService.logDocumentInteraction({
      file: mockFile3,
      userId: 'user-3',
      userName: 'User Three',
      action: 'download',
      sessionId: 'session-3'
    });
    
    // Create events for batch notification
    const events = [
      {
        fileId: mockFile1.id,
        fileName: mockFile1.name || 'Untitled',
        userId: 'user-1',
        userName: 'User One',
        action: 'view' as const,
        timestamp: new Date()
      },
      {
        fileId: mockFile2.id,
        fileName: mockFile2.name || 'Untitled',
        userId: 'user-2',
        userName: 'User Two',
        action: 'edit' as const,
        timestamp: new Date()
      },
      {
        fileId: mockFile3.id,
        fileName: mockFile3.name || 'Untitled',
        userId: 'user-3',
        userName: 'User Three',
        action: 'download' as const,
        timestamp: new Date()
      }
    ];
    
    const results = await slackService.sendBatchNotifications(events);
    
    expect(results.length).toBe(3);
    expect(results.every(r => r === true)).toBe(true);
    
    // Verify audit records
    const { records } = auditService.getAuditRecords();
    expect(records.length).toBe(3);
  });
});