import { GoogleSheetsService } from '../GoogleSheetsService';
import type { BotConfig } from '@/types';

describe('GoogleSheetsService', () => {
  let googleSheetsService: GoogleSheetsService;
  let mockConfig: BotConfig;

  beforeEach(() => {
    mockConfig = {
      discord: {
        token: 'test-token',
        clientId: 'test-client-id',
        prefix: '!',
        intents: [],
      },
      google: {
        spreadsheetId: 'test-spreadsheet-id',
        driveFolderId: 'test-folder-id',
        credentials: {
          client_email: 'test@example.com',
          private_key: 'test-private-key',
          project_id: 'test-project-id',
        },
      },
      ai: {
        provider: 'none',
        openai: {
          apiKey: 'test-openai-key',
          model: 'gpt-3.5-turbo',
          maxTokens: 1000,
          temperature: 0.7,
        },
        ollama: {
          host: 'http://localhost:11434',
          model: 'llama2',
        },
      },
      cache: {
        redis: {
          host: 'localhost',
          port: 6379,
          password: '',
          database: 0,
        },
        ttl: 3600,
      },
      metrics: {
        enabled: true,
        port: 9090,
        path: '/metrics',
      },
      security: {
        rateLimitWindow: 60000,
        rateLimitMax: 100,
        adminRole: 'admin',
        botUserRole: 'bot-user',
      },
      performance: {
        cacheTTL: 300,
        maxSearchResults: 50,
        maxAnalysisRows: 1000,
        requestTimeout: 30000,
        maxRetries: 3,
      },
      logging: {
        level: 'info',
        maxFiles: 5,
        maxSize: '10m',
        directory: './logs',
      },
      drive: {
        allowedMime: ['*'],
        ttlListSec: 300,
        ttlTextSec: 300,
        maxResults: 1000,
        rateQps: 5,
        rateBurst: 10,
      },
      features: {
        defaultLocale: 'uk',
      },
    } as unknown as BotConfig;

    googleSheetsService = new GoogleSheetsService(mockConfig);
  });

  describe('constructor', () => {
    it('should create GoogleSheetsService with correct configuration', () => {
      expect(googleSheetsService).toBeInstanceOf(GoogleSheetsService);
    });
  });

  describe('utility methods', () => {
    it('should get document type name correctly', () => {
      const typeName = (googleSheetsService as any).getDocumentTypeName('orders');
      expect(typeName).toBe('Накази');
    });

    it('should parse range correctly', () => {
      const [sheetName, cellRange] = (googleSheetsService as any).parseRange('Sheet1!A1:B2');
      expect(sheetName).toBe('Sheet1');
      expect(cellRange).toBe('A1:B2');
    });

    it('should handle range without cell reference', () => {
      const [sheetName, cellRange] = (googleSheetsService as any).parseRange('Sheet1');
      expect(sheetName).toBe('Sheet1');
      expect(cellRange).toBe('');
    });
  });

  describe('searchData', () => {
    it('should return test data in test environment', async () => {
      const result = await googleSheetsService.searchData('test', 10);
      expect(Array.isArray(result)).toBe(true);
      expect(result.length).toBeGreaterThan(0);
      expect(result[0]).toEqual(['id', 'query', 'timestamp']);
    });
  });
});