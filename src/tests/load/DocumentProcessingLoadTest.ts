/**
 * Load tests for document processing services
 * Tests performance with large document datasets
 */

import { jest, describe, it, expect, beforeAll, afterAll } from '@jest/globals';
import { DocumentSummarizationService } from '../../services/DocumentSummarizationService';
import { DocumentVersionComparisonService } from '../../services/DocumentVersionComparisonService';
import { DriveIndexerService } from '../../services/DriveIndexerService';
import { createMockConfig, createMockDriveFile } from '../utils/testHelpers';

describe('Document Processing Load Tests', () => {
  let mockConfig: any;

  beforeAll(() => {
    mockConfig = createMockConfig();
  });

  describe('Document Summarization Load Test', () => {
    it('should handle concurrent document summarization requests', async () => {
      const summarizationService = new DocumentSummarizationService(mockConfig);
      
      // Create mock AI service
      const mockAIService = {
        generateResponse: jest.fn().mockResolvedValue({
          content: 'This is a sample summary of the document content.',
          provider: 'openai',
          model: 'gpt-3.5-turbo',
          tokens: 100
        })
      };
      
      // Initialize service with mock AI service
      (summarizationService as any).aiService = mockAIService;
      
      const largeDocumentContent = 'Sample document content. '.repeat(1000); // Large content
      const mockFile = createMockDriveFile('test-file-id', 'Test Document.txt');
      
      const startTime = Date.now();
      
      // Process 100 documents concurrently
      const promises = Array(100).fill(null).map((_, i) => 
        summarizationService.summarizeDocument(
          { ...mockFile, id: `file-${i}` },
          largeDocumentContent
        )
      );
      
      const results = await Promise.all(promises);
      const duration = Date.now() - startTime;
      
      // Verify all summaries were generated
      expect(results).toHaveLength(100);
      expect(results[0].summary).toContain('sample summary');
      
      // Should complete within reasonable time
      expect(duration).toBeLessThan(30000); // 30 seconds
      
      // Verify caching works
      const cachedSummary = await summarizationService.summarizeDocument(
        { ...mockFile, id: 'file-0' },
        largeDocumentContent
      );
      
      expect(cachedSummary).toBeDefined();
    }, 60000); // 60 second timeout

    it('should maintain performance with large document content', async () => {
      const summarizationService = new DocumentSummarizationService(mockConfig);
      
      // Create mock AI service
      const mockAIService = {
        generateResponse: jest.fn().mockResolvedValue({
          content: 'This is a summary of the very large document content.',
          provider: 'openai',
          model: 'gpt-3.5-turbo',
          tokens: 150
        })
      };
      
      // Initialize service with mock AI service
      (summarizationService as any).aiService = mockAIService;
      
      // Very large document content (10,000 repetitions)
      const veryLargeDocumentContent = 'This is a sentence in a very large document. '.repeat(10000);
      const mockFile = createMockDriveFile('large-file-id', 'Very Large Document.txt');
      
      const startTime = Date.now();
      
      // Process the large document
      const summary = await summarizationService.summarizeDocument(
        mockFile,
        veryLargeDocumentContent
      );
      
      const duration = Date.now() - startTime;
      
      // Verify summary was generated
      expect(summary).toBeDefined();
      expect(summary.summary).toContain('summary');
      expect(summary.wordCount).toBeGreaterThan(10000);
      
      // Should complete within reasonable time
      expect(duration).toBeLessThan(45000); // 45 seconds
    }, 90000); // 90 second timeout
  });

  describe('Document Version Comparison Load Test', () => {
    it('should handle concurrent document version comparisons', async () => {
      const comparisonService = new DocumentVersionComparisonService(mockConfig);
      
      // Create mock AI service
      const mockAIService = {
        generateResponse: jest.fn().mockResolvedValue({
          content: 'Summary: Document has minor changes.\nAdditions: - New section added\nRemovals: - Old section removed\nSentiment: neutral',
          provider: 'openai',
          model: 'gpt-3.5-turbo',
          tokens: 120
        })
      };
      
      // Initialize service with mock AI service
      (comparisonService as any).aiService = mockAIService;
      
      const mockFile = createMockDriveFile('comparison-file-id', 'Comparison Document.txt');
      
      // Create multiple document versions
      const versions = Array(5).fill(null).map((_, i) => ({
        versionId: `version-${i}`,
        modifiedTime: new Date(Date.now() - (i * 86400000)).toISOString(), // One day apart
        author: `author-${i}`,
        content: `Content for version ${i}. `.repeat(100)
      }));
      
      const startTime = Date.now();
      
      // Compare versions for 50 documents concurrently
      const promises = Array(50).fill(null).map((_, i) => 
        comparisonService.compareDocumentVersions(
          { ...mockFile, id: `doc-${i}` },
          versions
        )
      );
      
      const results = await Promise.all(promises);
      const duration = Date.now() - startTime;
      
      // Verify all comparisons were completed
      expect(results).toHaveLength(50);
      expect(results[0].summary).toContain('Document has minor changes');
      
      // Should complete within reasonable time
      expect(duration).toBeLessThan(40000); // 40 seconds
    }, 80000); // 80 second timeout

    it('should handle complex document version comparisons', async () => {
      const comparisonService = new DocumentVersionComparisonService(mockConfig);
      
      // Create mock AI service
      const mockAIService = {
        generateResponse: jest.fn().mockResolvedValue({
          content: 'Summary: Major changes detected with significant additions and removals.\nAdditions: - New feature section\n- Updated guidelines\nRemovals: - Deprecated functions\n- Old examples\nSentiment: positive',
          provider: 'openai',
          model: 'gpt-3.5-turbo',
          tokens: 150
        })
      };
      
      // Initialize service with mock AI service
      (comparisonService as any).aiService = mockAIService;
      
      const mockFile = createMockDriveFile('complex-file-id', 'Complex Document.txt');
      
      // Create complex document versions with significant differences
      const versions = [
        {
          versionId: 'v1',
          modifiedTime: new Date(Date.now() - (30 * 86400000)).toISOString(), // 30 days ago
          author: 'author-1',
          content: 'Initial content. '.repeat(500) // Large initial content
        },
        {
          versionId: 'v2',
          modifiedTime: new Date(Date.now() - (15 * 86400000)).toISOString(), // 15 days ago
          author: 'author-2',
          content: 'Initial content. '.repeat(300) + ' New content added. '.repeat(300) // Mixed content
        },
        {
          versionId: 'v3',
          modifiedTime: new Date().toISOString(), // Now
          author: 'author-3',
          content: 'New content added. '.repeat(600) + ' Final content. '.repeat(200) // Mostly new content
        }
      ];
      
      const startTime = Date.now();
      
      // Perform complex comparison
      const comparison = await comparisonService.compareDocumentVersions(mockFile, versions);
      const duration = Date.now() - startTime;
      
      // Verify comparison results
      expect(comparison).toBeDefined();
      expect(comparison.summary).toContain('Major changes detected');
      expect(comparison.changes).toHaveLength(2); // Based on simple text comparison
      
      // Should complete within reasonable time
      expect(duration).toBeLessThan(30000); // 30 seconds
    }, 60000); // 60 second timeout
  });

  describe('Drive Indexer Load Test', () => {
    it('should handle indexing of large document sets', async () => {
      // Mock the Google service dependencies
      const mockGoogleService = {
        listDriveFiles: jest.fn().mockImplementation(async () => {
          // Return a page of files
          const files = Array(100).fill(null).map((_, i) => ({
            id: `file-${i}`,
            name: `Document ${i}`,
            mimeType: 'text/plain',
            modifiedTime: new Date().toISOString()
          }));
          
          return {
            files,
            nextPageToken: undefined // No more pages for simplicity
          };
        }),
        downloadFileContent: jest.fn().mockResolvedValue('Sample content for indexing')
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
      
      const startTime = Date.now();
      
      // Index a large set of documents
      await (indexerService as any).reindexAll('test-folder-id');
      
      const duration = Date.now() - startTime;
      
      // Verify the Google service was called
      expect(mockGoogleService.listDriveFiles).toHaveBeenCalled();
      expect(mockGoogleService.downloadFileContent).toHaveBeenCalledTimes(100);
      
      // Should complete within reasonable time
      expect(duration).toBeLessThan(60000); // 60 seconds
    }, 120000); // 2 minute timeout
  });
});