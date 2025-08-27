/**
 * Unit tests for DocumentAnalysisService
 */

import { describe, it, expect, beforeEach, jest } from '@jest/globals';
import { DocumentAnalysisService } from '../../../services/DocumentAnalysisService';
import { createMockConfig } from '../../utils/testHelpers';

// Mock the AIService and GoogleService
const mockAIService = {
  analyzeDocumentStructure: jest.fn().mockResolvedValue({ content: 'Structure analysis result' }),
  summarizeDocumentContent: jest.fn().mockResolvedValue({ content: 'Summary result' }),
  extractActionItems: jest.fn().mockResolvedValue({ content: 'Action items result' }),
  generateQnA: jest.fn().mockResolvedValue({ content: 'Q&A result' }),
  checkCompliance: jest.fn().mockResolvedValue({ content: 'Compliance result' }),
  translateDocument: jest.fn().mockResolvedValue({ content: 'Translation result' }),
  assessDocumentQuality: jest.fn().mockResolvedValue({ content: 'Quality assessment result' }),
  analyzeStakeholders: jest.fn().mockResolvedValue({ content: 'Stakeholders analysis result' }),
  analyzeBudget: jest.fn().mockResolvedValue({ content: 'Budget analysis result' }),
  assessRisks: jest.fn().mockResolvedValue({ content: 'Risk assessment result' }),
  segmentAudience: jest.fn().mockResolvedValue({ content: 'Audience segmentation result' }),
  analyzeVersionChanges: jest.fn().mockResolvedValue({ content: 'Version changes result' }),
  predictDocumentPerformance: jest.fn().mockResolvedValue({ content: 'Performance prediction result' }),
};

const mockGoogleService = {
  extractTextForChat: jest.fn().mockResolvedValue({ text: 'Document content' }),
  getDriveFileMetadata: jest.fn().mockResolvedValue({ mimeType: 'application/vnd.google-apps.document' }),
};

describe('DocumentAnalysisService', () => {
  let documentAnalysisService: DocumentAnalysisService;
  let mockConfig: any;

  beforeEach(() => {
    mockConfig = createMockConfig();
    documentAnalysisService = new DocumentAnalysisService(mockConfig);
    (documentAnalysisService as any).initializeServices(mockAIService, mockGoogleService);
  });

  describe('constructor', () => {
    it('should initialize successfully with valid config', () => {
      expect(documentAnalysisService).toBeDefined();
    });
  });

  describe('analyzeDocument', () => {
    it('should perform full document analysis successfully', async () => {
      const mockFile = {
        id: 'test-file-id',
        name: 'Test Document.txt',
      };

      const result = await documentAnalysisService.analyzeDocument(mockFile);

      expect(result).toBeDefined();
      expect(result.fileId).toBe('test-file-id');
      expect(result.fileName).toBe('Test Document.txt');
      expect(mockGoogleService.extractTextForChat).toHaveBeenCalledWith('test-file-id');
    });

    it('should handle service initialization errors gracefully', async () => {
      const serviceWithoutDependencies = new DocumentAnalysisService(mockConfig);
      
      const mockFile = {
        id: 'test-file-id',
        name: 'Test Document.txt',
      };

      await expect(serviceWithoutDependencies.analyzeDocument(mockFile))
        .rejects
        .toThrow('Google service not initialized');
    });
  });

  describe('compareDocumentVersions', () => {
    it('should compare two document versions successfully', async () => {
      const oldFile = {
        id: 'old-file-id',
        name: 'Old Document.txt',
      };

      const newFile = {
        id: 'new-file-id',
        name: 'New Document.txt',
      };

      const result = await documentAnalysisService.compareDocumentVersions(oldFile, newFile);

      expect(result).toBeDefined();
      expect(mockGoogleService.extractTextForChat).toHaveBeenCalledWith('old-file-id');
      expect(mockGoogleService.extractTextForChat).toHaveBeenCalledWith('new-file-id');
      expect(mockAIService.analyzeVersionChanges).toHaveBeenCalled();
    });
  });

  describe('predictPerformance', () => {
    it('should predict document performance successfully', async () => {
      const mockFile = {
        id: 'test-file-id',
        name: 'Test Document.txt',
      };

      const result = await documentAnalysisService.predictPerformance(mockFile);

      expect(result).toBeDefined();
      expect(mockGoogleService.extractTextForChat).toHaveBeenCalledWith('test-file-id');
      expect(mockAIService.predictDocumentPerformance).toHaveBeenCalled();
    });
  });

  describe('cache management', () => {
    it('should cache analysis results', async () => {
      const mockFile = {
        id: 'test-file-id',
        name: 'Test Document.txt',
      };

      // First analysis
      await documentAnalysisService.analyzeDocument(mockFile);
      
      // Second analysis should use cache
      const cachedResult = documentAnalysisService.getAnalysis('test-file-id');
      
      expect(cachedResult).toBeDefined();
      expect(cachedResult?.fileId).toBe('test-file-id');
    });

    it('should clear cached analysis', () => {
      const mockFile = {
        id: 'test-file-id',
        name: 'Test Document.txt',
      };

      // Add to cache
      (documentAnalysisService as any).cacheAnalysis('test-file-id', {
        fileId: 'test-file-id',
        fileName: 'Test Document.txt',
        generatedAt: new Date()
      });

      // Verify it's in cache
      const cachedResult = documentAnalysisService.getAnalysis('test-file-id');
      expect(cachedResult).toBeDefined();

      // Clear cache
      documentAnalysisService.clearAnalysis('test-file-id');

      // Verify it's no longer in cache
      const clearedResult = documentAnalysisService.getAnalysis('test-file-id');
      expect(clearedResult).toBeUndefined();
    });

    it('should clear all cached analyses', () => {
      // Add multiple items to cache
      (documentAnalysisService as any).cacheAnalysis('file-1', {
        fileId: 'file-1',
        fileName: 'Document 1.txt',
        generatedAt: new Date()
      });

      (documentAnalysisService as any).cacheAnalysis('file-2', {
        fileId: 'file-2',
        fileName: 'Document 2.txt',
        generatedAt: new Date()
      });

      // Verify items are in cache
      expect(documentAnalysisService.getAnalysis('file-1')).toBeDefined();
      expect(documentAnalysisService.getAnalysis('file-2')).toBeDefined();

      // Clear all cache
      documentAnalysisService.clearAllAnalyses();

      // Verify cache is empty
      expect(documentAnalysisService.getAnalysis('file-1')).toBeUndefined();
      expect(documentAnalysisService.getAnalysis('file-2')).toBeUndefined();
    });
  });
});