/**
 * Unit tests for SlackIntegrationService
 */

import { describe, it, expect, beforeEach, jest } from '@jest/globals';
import { SlackIntegrationService } from '../../../services/SlackIntegrationService';
import { createMockConfig } from '../../utils/testHelpers';

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

describe('SlackIntegrationService', () => {
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
    slackService = new SlackIntegrationService(mockConfig);
  });

  describe('constructor', () => {
    it('should initialize successfully with valid config', () => {
      expect(slackService).toBeDefined();
    });

    it('should handle missing Slack config gracefully', () => {
      const configWithoutSlack = createMockConfig();
      configWithoutSlack.integrations = {};
      
      const service = new SlackIntegrationService(configWithoutSlack);
      expect(service).toBeDefined();
    });
  });

  describe('sendDocumentNotification', () => {
    it('should send document notification successfully', async () => {
      const event = {
        fileId: 'test-file-id',
        fileName: 'Test Document.txt',
        userId: 'user-123',
        userName: 'Test User',
        action: 'created' as const,
        timestamp: new Date()
      };

      const result = await slackService.sendDocumentNotification(event);
      
      expect(result).toBe(true);
    });

    it('should handle Slack client errors gracefully', async () => {
      // Mock a Slack API error
      const mockSlackClient = (slackService as any).slackClient;
      mockSlackClient.chat.postMessage.mockRejectedValue(new Error('Slack API error'));

      const event = {
        fileId: 'test-file-id',
        fileName: 'Test Document.txt',
        userId: 'user-123',
        userName: 'Test User',
        action: 'created' as const,
        timestamp: new Date()
      };

      const result = await slackService.sendDocumentNotification(event);
      
      expect(result).toBe(false);
    });

    it('should skip notification if Slack is not configured', async () => {
      // Create service without Slack config
      const configWithoutSlack = createMockConfig();
      configWithoutSlack.integrations = {};
      const serviceWithoutSlack = new SlackIntegrationService(configWithoutSlack);

      const event = {
        fileId: 'test-file-id',
        fileName: 'Test Document.txt',
        userId: 'user-123',
        userName: 'Test User',
        action: 'created' as const,
        timestamp: new Date()
      };

      const result = await serviceWithoutSlack.sendDocumentNotification(event);
      
      expect(result).toBe(false);
    });
  });

  describe('createDefaultMessage', () => {
    it('should create appropriate message for different actions', () => {
      const event = {
        fileId: 'test-file-id',
        fileName: 'Test Document.txt',
        userId: 'user-123',
        userName: 'Test User',
        action: 'created' as const,
        timestamp: new Date()
      };

      // Use reflection to test private method
      const message = (slackService as any).createDefaultMessage(event);
      
      expect(message.blocks).toBeDefined();
      expect(message.blocks.length).toBe(2);
      expect(message.blocks[0].type).toBe('section');
    });

    it('should add sensitivity warning for sensitive documents', () => {
      const event = {
        fileId: 'test-file-id',
        fileName: 'Confidential Report.txt',
        userId: 'user-123',
        userName: 'Test User',
        action: 'created' as const,
        timestamp: new Date()
      };

      const message = (slackService as any).createDefaultMessage(event);
      
      expect(JSON.stringify(message.blocks)).toContain('Sensitive Document');
    });
  });

  describe('isSensitiveDocument', () => {
    it('should detect sensitive documents by filename', () => {
      const sensitiveFileName = 'confidential-report.txt';
      const result = (slackService as any).isSensitiveDocument(sensitiveFileName);
      
      expect(result).toBe(true);
    });

    it('should return false for non-sensitive documents', () => {
      const regularFileName = 'regular-document.txt';
      const result = (slackService as any).isSensitiveDocument(regularFileName);
      
      expect(result).toBe(false);
    });
  });

  describe('getStats', () => {
    it('should return initial stats', () => {
      const stats = slackService.getStats();
      
      expect(stats.totalNotificationsSent).toBe(0);
      expect(stats.totalNotificationsFailed).toBe(0);
      expect(stats.averageResponseTime).toBe(0);
    });
  });

  describe('sendBatchNotifications', () => {
    it('should send batch notifications', async () => {
      const events = [
        {
          fileId: 'file-1',
          fileName: 'Document 1.txt',
          userId: 'user-1',
          userName: 'User One',
          action: 'created' as const,
          timestamp: new Date()
        },
        {
          fileId: 'file-2',
          fileName: 'Document 2.txt',
          userId: 'user-2',
          userName: 'User Two',
          action: 'updated' as const,
          timestamp: new Date()
        }
      ];

      const results = await slackService.sendBatchNotifications(events);
      
      expect(results.length).toBe(2);
      expect(results[0]).toBe(true);
      expect(results[1]).toBe(true);
    });
  });

  describe('sendCustomMessage', () => {
    it('should send custom message successfully', async () => {
      const message = {
        text: 'Custom test message'
      };

      const result = await slackService.sendCustomMessage(message);
      
      expect(result).toBe(true);
    });
  });

  describe('isConfigured', () => {
    it('should return true when properly configured', () => {
      const result = slackService.isConfigured();
      
      expect(result).toBe(true);
    });

    it('should return false when not configured', () => {
      const configWithoutSlack = createMockConfig();
      configWithoutSlack.integrations = {};
      const serviceWithoutSlack = new SlackIntegrationService(configWithoutSlack);
      
      const result = serviceWithoutSlack.isConfigured();
      
      expect(result).toBe(false);
    });
  });

  describe('testConnection', () => {
    it('should test connection successfully', async () => {
      const result = await slackService.testConnection();
      
      expect(result).toBe(true);
    });
  });
});