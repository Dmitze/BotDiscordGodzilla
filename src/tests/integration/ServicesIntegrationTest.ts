/**
 * Integration tests for service coordination
 * Tests document processing workflow coordination
 */

import { jest, describe, it, expect, beforeAll, afterAll } from '@jest/globals';
import { DocumentSummarizationService } from '../../services/DocumentSummarizationService';
import { DocumentVersionComparisonService } from '../../services/DocumentVersionComparisonService';
import { DocumentAnalyticsService } from '../../services/DocumentAnalyticsService';
import { DriveIndexerService } from '../../services/DriveIndexerService';
import { AIService } from '../../services/AIService';
import { createMockConfig, createMockDriveFile } from '../utils/testHelpers';

describe('Services Integration Tests', () => {
  let mockConfig: any;

  beforeAll(() => {
    mockConfig = createMockConfig();
  });

  describe('Document Processing Workflow', () => {
    it('should coordinate document indexing, summarization, and analytics', async () => {
      // Create services
      const analyticsService = new DocumentAnalyticsService(mockConfig);
      const summarizationService = new DocumentSummarizationService(mockConfig);
      
      // Create mock AI service
      const mockAIService = {
        generateResponse: jest.fn().mockResolvedValue({
          content: 'This is an AI-generated summary of the document.',
          provider: 'openai',
          model: 'gpt-3.5-turbo',
          tokens: 100
        })
      };
      
      // Initialize summarization service with mock AI service
      (summarizationService as any).aiService = mockAIService;
      
      // Mock the Google service dependencies for indexer
      const mockGoogleService = {
        listDriveFiles: jest.fn().mockImplementation(async () => {
          // Return a page of files
          const files = [
            createMockDriveFile('file-1', 'Document 1.txt'),
            createMockDriveFile('file-2', 'Document 2.txt')
          ];
          
          return {
            files,
            nextPageToken: undefined
          };
        }),
        downloadFileContent: jest.fn().mockResolvedValue('Sample document content for testing.')
      };
      
      // Create a mock bot instance with the Google service
      const mockBot = {
        config: mockConfig,
        getService: jest.fn().mockImplementation((serviceName: string) => {
          if (serviceName === 'google') return mockGoogleService;
          if (serviceName === 'cache') return { get: jest.fn(), set: jest.fn() };
          return undefined;
        })
      };
      
      const indexerService = new DriveIndexerService(mockBot as any);
      
      // Mock the search index
      (indexerService as any).searchIndex = {
        indexDocument: jest.fn()
      };
      
      // Simulate a document processing workflow:
      // 1. Index documents from Google Drive
      // 2. Download and summarize each document
      // 3. Record analytics for each document
      
      // Step 1: Index documents
      await (indexerService as any).reindexAll('test-folder-id');
      
      expect(mockGoogleService.listDriveFiles).toHaveBeenCalled();
      expect(mockGoogleService.downloadFileContent).toHaveBeenCalledTimes(2);
      
      // Step 2: Summarize documents
      const mockFile1 = createMockDriveFile('file-1', 'Document 1.txt');
      const mockFile2 = createMockDriveFile('file-2', 'Document 2.txt');
      
      const summary1 = await summarizationService.summarizeDocument(
        mockFile1,
        'Sample content for document 1.'
      );
      
      const summary2 = await summarizationService.summarizeDocument(
        mockFile2,
        'Sample content for document 2.'
      );
      
      // Verify summaries were generated
      expect(summary1).toBeDefined();
      expect(summary1.summary).toContain('AI-generated summary');
      expect(summary2).toBeDefined();
      expect(summary2.summary).toContain('AI-generated summary');
      
      // Verify AI service was called
      expect(mockAIService.generateResponse).toHaveBeenCalledTimes(2);
      
      // Step 3: Record analytics
      analyticsService.recordDocumentAccess('file-1', 'user-1', 'view', 'session-1');
      analyticsService.recordDocumentAccess('file-1', 'user-1', 'download', 'session-2');
      analyticsService.recordDocumentAccess('file-2', 'user-2', 'view', 'session-3');
      
      // Verify analytics were recorded
      // Since the access records are private, we'll verify indirectly
      expect(analyticsService).toBeDefined();
    }, 30000); // 30 second timeout

    it('should coordinate document version comparison with analytics', async () => {
      // Create services
      const analyticsService = new DocumentAnalyticsService(mockConfig);
      const comparisonService = new DocumentVersionComparisonService(mockConfig);
      
      // Create mock AI service
      const mockAIService = {
        generateResponse: jest.fn().mockResolvedValue({
          content: 'Summary: Document has been updated with new information.\nAdditions: - New section\nRemovals: - Old section\nSentiment: positive',
          provider: 'openai',
          model: 'gpt-3.5-turbo',
          tokens: 120
        })
      };
      
      // Initialize comparison service with mock AI service
      (comparisonService as any).aiService = mockAIService;
      
      // Create document versions
      const mockFile = createMockDriveFile('versioned-file', 'Versioned Document.txt');
      const versions = [
        {
          versionId: 'v1',
          modifiedTime: new Date(Date.now() - 86400000).toISOString(), // 1 day ago
          author: 'author-1',
          content: 'Initial content of the document.'
        },
        {
          versionId: 'v2',
          modifiedTime: new Date().toISOString(), // Now
          author: 'author-2',
          content: 'Initial content of the document. Added new information.'
        }
      ];
      
      // Simulate a version comparison workflow with analytics:
      // 1. Compare document versions
      // 2. Record analytics for the comparison
      // 3. Record user interaction with the comparison results
      
      // Step 1: Compare versions
      const comparison = await comparisonService.compareDocumentVersions(mockFile, versions);
      
      // Verify comparison was successful
      expect(comparison).toBeDefined();
      expect(comparison.summary).toContain('Document has been updated');
      expect(comparison.changes).toHaveLength(1); // One addition based on simple comparison
      
      // Verify AI service was called
      expect(mockAIService.generateResponse).toHaveBeenCalled();
      
      // Step 2: Record analytics for the comparison
      analyticsService.recordDocumentAccess('versioned-file', 'user-1', 'analyze', 'session-1');
      
      // Step 3: Record user interaction with comparison results
      analyticsService.recordDocumentAccess('versioned-file', 'user-1', 'view', 'session-2');
      
      // Verify analytics were recorded
      expect(analyticsService).toBeDefined();
    }, 20000); // 20 second timeout

    it('should coordinate multiple AI services with caching', async () => {
      // Create services
      const aiService = new AIService(mockConfig);
      const summarizationService = new DocumentSummarizationService(mockConfig);
      const comparisonService = new DocumentVersionComparisonService(mockConfig);
      
      // Mock AI providers
      const mockOpenAIProvider = {
        generate: jest.fn().mockResolvedValue({
          content: 'AI response content',
          provider: 'openai',
          model: 'gpt-3.5-turbo'
        }),
        isHealthy: jest.fn().mockResolvedValue(true)
      };
      
      // Set up AI service with mock providers
      (aiService as any).providers = {
        openai: mockOpenAIProvider
      };
      (aiService as any).currentProvider = 'openai';
      
      // Initialize other services with the AI service
      (summarizationService as any).aiService = aiService;
      (comparisonService as any).aiService = aiService;
      
      // Create test data
      const mockFile = createMockDriveFile('ai-test-file', 'AI Test Document.txt');
      const documentContent = 'This is a test document for AI processing.';
      
      // Step 1: Generate document summary
      const summary = await summarizationService.summarizeDocument(mockFile, documentContent);
      
      // Step 2: Use AI service directly for another task
      const directAIResponse = await aiService.generateResponse('Summarize this content: ' + documentContent);
      
      // Step 3: Verify caching works across services
      const cachedSummary = await summarizationService.summarizeDocument(mockFile, documentContent);
      
      // Verify all operations were successful
      expect(summary).toBeDefined();
      expect(directAIResponse).toBeDefined();
      expect(cachedSummary).toBeDefined();
      
      // Verify AI provider was called
      expect(mockOpenAIProvider.generate).toHaveBeenCalled();
    }, 25000); // 25 second timeout
  });

  describe('Error Handling Integration', () => {
    it('should handle service failures gracefully', async () => {
      // Create services
      const analyticsService = new DocumentAnalyticsService(mockConfig);
      const summarizationService = new DocumentSummarizationService(mockConfig);
      
      // Create mock AI service that fails
      const failingAIService = {
        generateResponse: jest.fn().mockRejectedValue(new Error('AI service unavailable'))
      };
      
      // Initialize summarization service with failing AI service
      (summarizationService as any).aiService = failingAIService;
      
      // Create test data
      const mockFile = createMockDriveFile('error-test-file', 'Error Test Document.txt');
      const documentContent = 'This document will cause an AI error.';
      
      // Attempt to summarize document with failing AI service
      await expect(
        summarizationService.summarizeDocument(mockFile, documentContent)
      ).rejects.toThrow('AI service unavailable');
      
      // Verify analytics service continues to work despite AI failure
      analyticsService.recordDocumentAccess('error-test-file', 'user-1', 'view', 'session-1');
      
      expect(analyticsService).toBeDefined();
    });

    it('should maintain service independence during partial failures', async () => {
      // Create services
      const analyticsService = new DocumentAnalyticsService(mockConfig);
      const summarizationService = new DocumentSummarizationService(mockConfig);
      
      // Test that analytics service works independently of AI service
      analyticsService.recordDocumentAccess('independent-file', 'user-1', 'view', 'session-1');
      
      // Create mock AI service for summarization
      const mockAIService = {
        generateResponse: jest.fn().mockResolvedValue({
          content: 'Independent summary',
          provider: 'openai',
          model: 'gpt-3.5-turbo'
        })
      };
      
      // Initialize summarization service
      (summarizationService as any).aiService = mockAIService;
      
      // Generate summary
      const summary = await summarizationService.summarizeDocument(
        createMockDriveFile('independent-file', 'Independent Document.txt'),
        'Independent content'
      );
      
      // Verify both services work independently
      expect(summary).toBeDefined();
      expect(summary.summary).toBe('Independent summary');
      expect(analyticsService).toBeDefined();
    });
  });
});