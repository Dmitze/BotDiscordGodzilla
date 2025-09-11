/**
 * Unit tests for DocumentAnalysisCommand
 */

import { describe, it, expect, beforeEach, jest } from '@jest/globals';
import { DocumentAnalysisCommand } from '../../../commands/DocumentAnalysisCommand';
import { createMockConfig } from '../../utils/testHelpers';

// Mock Discord.js interaction
const mockInteraction = {
  deferReply: jest.fn().mockResolvedValue(undefined),
  editReply: jest.fn().mockResolvedValue(undefined),
  options: {
    getString: jest.fn().mockImplementation((name) => {
      if (name === 'file') return 'test-document';
      if (name === 'type') return 'full';
      return null;
    }),
  },
  user: {
    id: 'user-123',
  },
};

// Mock services
const mockGoogleService = {
  searchFiles: jest.fn().mockResolvedValue([{
    id: 'test-file-id',
    name: 'Test Document.txt',
  }]),
  getDriveFile: jest.fn().mockResolvedValue({
    id: 'test-file-id',
    name: 'Test Document.txt',
  }),
};

const mockDocumentAnalysisService = {
  analyzeDocument: jest.fn().mockResolvedValue({
    fileId: 'test-file-id',
    fileName: 'Test Document.txt',
    generatedAt: new Date(),
  }),
};

const mockAnalyticsService = {
  trackCommandUsage: jest.fn().mockResolvedValue(undefined),
};

describe('DocumentAnalysisCommand', () => {
  let documentAnalysisCommand: DocumentAnalysisCommand;
  let mockConfig: any;

  beforeEach(() => {
    mockConfig = createMockConfig();
    documentAnalysisCommand = new DocumentAnalysisCommand(
      mockConfig, 
      mockGoogleService, 
      mockDocumentAnalysisService
    );
  });

  describe('constructor', () => {
    it('should initialize successfully with valid config and services', () => {
      expect(documentAnalysisCommand).toBeDefined();
      expect(documentAnalysisCommand.getName()).toBe('analyze-doc');
    });
  });

  describe('execute', () => {
    it('should execute successfully with valid inputs', async () => {
      const mockServices = {
        analytics: mockAnalyticsService,
      };

      await documentAnalysisCommand.execute({
        interaction: mockInteraction as any,
        services: mockServices,
      });

      expect(mockInteraction.deferReply).toHaveBeenCalled();
      expect(mockGoogleService.searchFiles).toHaveBeenCalledWith(`name contains 'test-document'`);
      expect(mockDocumentAnalysisService.analyzeDocument).toHaveBeenCalledWith(
        expect.objectContaining({ id: 'test-file-id' }),
        expect.objectContaining({ includeStructure: true, includeSummary: true, includeActionItems: true, includeCompliance: true, includeQuality: true })
      );
      expect(mockInteraction.editReply).toHaveBeenCalled();
    });

    it('should handle missing file parameter', async () => {
      const interactionWithoutFile = {
        ...mockInteraction,
        options: {
          getString: jest.fn().mockImplementation((name) => {
            if (name === 'file') return null;
            if (name === 'type') return 'full';
            return null;
          }),
        },
      };

      await documentAnalysisCommand.execute({
        interaction: interactionWithoutFile as any,
        services: {},
      });

      expect(interactionWithoutFile.deferReply).toHaveBeenCalled();
      expect(interactionWithoutFile.editReply).toHaveBeenCalledWith({
        content: 'document.analysis.error.no_file',
      });
    });

    it('should handle missing Google service', async () => {
      const commandWithoutGoogleService = new DocumentAnalysisCommand(
        mockConfig, 
        undefined, 
        mockDocumentAnalysisService
      );

      await commandWithoutGoogleService.execute({
        interaction: mockInteraction as any,
        services: {},
      });

      expect(mockInteraction.deferReply).toHaveBeenCalled();
      expect(mockInteraction.editReply).toHaveBeenCalledWith({
        content: 'document.analysis.error.google_service',
      });
    });

    it('should handle missing Document Analysis service', async () => {
      const commandWithoutAnalysisService = new DocumentAnalysisCommand(
        mockConfig, 
        mockGoogleService, 
        undefined
      );

      await commandWithoutAnalysisService.execute({
        interaction: mockInteraction as any,
        services: {},
      });

      expect(mockInteraction.deferReply).toHaveBeenCalled();
      expect(mockInteraction.editReply).toHaveBeenCalledWith({
        content: 'document.analysis.error.analysis_service',
      });
    });

    it('should handle file not found', async () => {
      const googleServiceWithoutResults = {
        searchFiles: jest.fn().mockResolvedValue([]),
      };

      const commandWithNoFileResults = new DocumentAnalysisCommand(
        mockConfig, 
        googleServiceWithoutResults as any, 
        mockDocumentAnalysisService
      );

      await commandWithNoFileResults.execute({
        interaction: mockInteraction as any,
        services: {},
      });

      expect(mockInteraction.deferReply).toHaveBeenCalled();
      expect(googleServiceWithoutResults.searchFiles).toHaveBeenCalledWith(`name contains 'test-document'`);
      expect(mockInteraction.editReply).toHaveBeenCalledWith({
        content: 'document.analysis.error.file_not_found',
      });
    });

    it('should handle execution errors gracefully', async () => {
      const failingGoogleService = {
        searchFiles: jest.fn().mockRejectedValue(new Error('Search failed')),
      };

      const commandWithFailingService = new DocumentAnalysisCommand(
        mockConfig, 
        failingGoogleService as any, 
        mockDocumentAnalysisService
      );

      await commandWithFailingService.execute({
        interaction: mockInteraction as any,
        services: {},
      });

      expect(mockInteraction.deferReply).toHaveBeenCalled();
      expect(mockInteraction.editReply).toHaveBeenCalledWith({
        content: 'document.analysis.error.analysis_failed',
      });
    });
  });

  describe('analysis type handling', () => {
    it('should handle structure analysis type', async () => {
      const interactionWithStructureType = {
        ...mockInteraction,
        options: {
          getString: jest.fn().mockImplementation((name) => {
            if (name === 'file') return 'test-document';
            if (name === 'type') return 'structure';
            return null;
          }),
        },
      };

      const mockServices = {
        analytics: mockAnalyticsService,
      };

      await documentAnalysisCommand.execute({
        interaction: interactionWithStructureType as any,
        services: mockServices,
      });

      expect(interactionWithStructureType.deferReply).toHaveBeenCalled();
      expect(mockDocumentAnalysisService.analyzeDocument).toHaveBeenCalledWith(
        expect.objectContaining({ id: 'test-file-id' }),
        expect.objectContaining({ includeStructure: true })
      );
    });

    it('should handle summary analysis type', async () => {
      const interactionWithSummaryType = {
        ...mockInteraction,
        options: {
          getString: jest.fn().mockImplementation((name) => {
            if (name === 'file') return 'test-document';
            if (name === 'type') return 'summary';
            return null;
          }),
        },
      };

      const mockServices = {
        analytics: mockAnalyticsService,
      };

      await documentAnalysisCommand.execute({
        interaction: interactionWithSummaryType as any,
        services: mockServices,
      });

      expect(interactionWithSummaryType.deferReply).toHaveBeenCalled();
      expect(mockDocumentAnalysisService.analyzeDocument).toHaveBeenCalledWith(
        expect.objectContaining({ id: 'test-file-id' }),
        expect.objectContaining({ includeSummary: true })
      );
    });

    it('should handle action items analysis type', async () => {
      const interactionWithActionsType = {
        ...mockInteraction,
        options: {
          getString: jest.fn().mockImplementation((name) => {
            if (name === 'file') return 'test-document';
            if (name === 'type') return 'actions';
            return null;
          }),
        },
      };

      const mockServices = {
        analytics: mockAnalyticsService,
      };

      await documentAnalysisCommand.execute({
        interaction: interactionWithActionsType as any,
        services: mockServices,
      });

      expect(interactionWithActionsType.deferReply).toHaveBeenCalled();
      expect(mockDocumentAnalysisService.analyzeDocument).toHaveBeenCalledWith(
        expect.objectContaining({ id: 'test-file-id' }),
        expect.objectContaining({ includeActionItems: true })
      );
    });

    it('should handle compliance analysis type', async () => {
      const interactionWithComplianceType = {
        ...mockInteraction,
        options: {
          getString: jest.fn().mockImplementation((name) => {
            if (name === 'file') return 'test-document';
            if (name === 'type') return 'compliance';
            return null;
          }),
        },
      };

      const mockServices = {
        analytics: mockAnalyticsService,
      };

      await documentAnalysisCommand.execute({
        interaction: interactionWithComplianceType as any,
        services: mockServices,
      });

      expect(interactionWithComplianceType.deferReply).toHaveBeenCalled();
      expect(mockDocumentAnalysisService.analyzeDocument).toHaveBeenCalledWith(
        expect.objectContaining({ id: 'test-file-id' }),
        expect.objectContaining({ includeCompliance: true })
      );
    });

    it('should handle quality analysis type', async () => {
      const interactionWithQualityType = {
        ...mockInteraction,
        options: {
          getString: jest.fn().mockImplementation((name) => {
            if (name === 'file') return 'test-document';
            if (name === 'type') return 'quality';
            return null;
          }),
        },
      };

      const mockServices = {
        analytics: mockAnalyticsService,
      };

      await documentAnalysisCommand.execute({
        interaction: interactionWithQualityType as any,
        services: mockServices,
      });

      expect(interactionWithQualityType.deferReply).toHaveBeenCalled();
      expect(mockDocumentAnalysisService.analyzeDocument).toHaveBeenCalledWith(
        expect.objectContaining({ id: 'test-file-id' }),
        expect.objectContaining({ includeQuality: true })
      );
    });
  });
});