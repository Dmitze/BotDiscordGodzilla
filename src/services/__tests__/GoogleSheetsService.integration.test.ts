import { Bot } from '@/core/Bot';
import type { BotConfig } from '@/types';

describe('GoogleSheetsService Integration', () => {
  let bot: Bot;
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

    bot = new Bot(mockConfig);
  });

  describe('Service Registration', () => {
    it('should register GoogleSheetsService', async () => {
      await bot.initialize();
      const googleService = bot.getService('google');
      expect(googleService).toBeDefined();
      expect(googleService.constructor.name).toBe('GoogleSheetsService');
    });

    it('should have required methods', async () => {
      await bot.initialize();
      const googleService = bot.getService('google');
      
      // Check that all required methods exist
      expect(typeof (googleService as any).listSheets).toBe('function');
      expect(typeof (googleService as any).getSheetData).toBe('function');
      expect(typeof (googleService as any).writeSheetData).toBe('function');
      expect(typeof (googleService as any).searchData).toBe('function');
      expect(typeof (googleService as any).extractTextForChat).toBe('function');
    });
  });

  describe('Health Check', () => {
    it('should return health status', async () => {
      await bot.initialize();
      const googleService = bot.getService('google');
      
      // In test mode, the service should be healthy even without real credentials
      const health = await (googleService as any).healthCheck();
      expect(health).toBeDefined();
      // In test mode, it might be healthy or not depending on configuration
    });
  });
});