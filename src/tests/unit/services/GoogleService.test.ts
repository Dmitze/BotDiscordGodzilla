/**
 * Unit тесты для GoogleService
 */

import { jest, describe, it, expect, beforeEach } from '@jest/globals';
import { GoogleService } from '../../../services/GoogleService';
import { createMockConfig } from '../../utils/testHelpers';

// Моки для Google APIs
jest.mock('googleapis', () => ({
  google: {
    sheets: jest.fn(() => ({
      spreadsheets: {
        values: {
          get: jest.fn(),
        },
      },
    })),
    drive: jest.fn(() => ({
      files: {
        list: jest.fn(),
        get: jest.fn(),
      },
    })),
  },
}));

describe('GoogleService', () => {
  let googleService: GoogleService;
  let mockConfig: any;

  beforeEach(() => {
    mockConfig = createMockConfig();
    googleService = new GoogleService(mockConfig);
  });

  describe('constructor', () => {
    it('should create GoogleService instance', () => {
      expect(googleService).toBeInstanceOf(GoogleService);
    });

    it('should have correct service name', () => {
      expect(googleService.getName()).toBe('GoogleService');
    });
  });

  describe('initialization', () => {
    it('should initialize successfully', async () => {
      await expect(googleService.initialize()).resolves.not.toThrow();
    });

    it('should handle initialization error', async () => {
      // Мокаем ошибку инициализации
      jest.spyOn(googleService as any, 'authenticate').mockRejectedValue(new Error('Auth error'));

      await expect(googleService.initialize()).rejects.toThrow('Auth error');
    });
  });

  describe('searchData', () => {
    beforeEach(async () => {
      await googleService.initialize();
    });

    it('should search data successfully', async () => {
      const mockData = [
        ['ID', 'Name', 'Value'],
        ['1', 'Test 1', '100'],
        ['2', 'Test 2', '200'],
      ];

      // Мокаем Google Sheets API
      const mockSheetsApi = {
        spreadsheets: {
          values: {
            get: jest.fn().mockResolvedValue({
              data: {
                values: mockData,
              },
            }),
          },
        },
      };

      (googleService as any).sheetsApi = mockSheetsApi;

      const result = await googleService.searchData('test', 10);

      expect(result).toEqual(mockData);
      expect(mockSheetsApi.spreadsheets.values.get).toHaveBeenCalled();
    });

    it('should handle empty search results', async () => {
      const mockSheetsApi = {
        spreadsheets: {
          values: {
            get: jest.fn().mockResolvedValue({
              data: {
                values: [],
              },
            }),
          },
        },
      };

      (googleService as any).sheetsApi = mockSheetsApi;

      const result = await googleService.searchData('nonexistent', 10);

      expect(result).toEqual([]);
    });

    it('should handle API error', async () => {
      const mockSheetsApi = {
        spreadsheets: {
          values: {
            get: jest.fn().mockRejectedValue(new Error('API error')),
          },
        },
      };

      (googleService as any).sheetsApi = mockSheetsApi;

      await expect(googleService.searchData('test', 10)).rejects.toThrow('API error');
    });
  });

  describe('searchDocuments', () => {
    beforeEach(async () => {
      await googleService.initialize();
    });

    it('should search documents successfully', async () => {
      const mockFiles = [
        { id: '1', name: 'Document 1', mimeType: 'application/pdf' },
        { id: '2', name: 'Document 2', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
      ];

      const mockDriveApi = {
        files: {
          list: jest.fn().mockResolvedValue({
            data: {
              files: mockFiles,
            },
          }),
        },
      };

      (googleService as any).driveApi = mockDriveApi;

      const result = await googleService.searchDocuments('test');

      expect(result).toEqual(mockFiles);
      expect(mockDriveApi.files.list).toHaveBeenCalled();
    });

    it('should handle empty document results', async () => {
      const mockDriveApi = {
        files: {
          list: jest.fn().mockResolvedValue({
            data: {
              files: [],
            },
          }),
        },
      };

      (googleService as any).driveApi = mockDriveApi;

      const result = await googleService.searchDocuments('nonexistent');

      expect(result).toEqual([]);
    });
  });

  describe('health check', () => {
    it('should return healthy status when initialized', async () => {
      await googleService.initialize();
      
      const health = await googleService.getHealthStatus();
      
      expect(health.healthy).toBe(true);
      expect(health.service).toBe('GoogleService');
    });

    it('should return unhealthy status when not initialized', async () => {
      const health = await googleService.getHealthStatus();
      
      expect(health.healthy).toBe(false);
      expect(health.service).toBe('GoogleService');
    });
  });
}); 