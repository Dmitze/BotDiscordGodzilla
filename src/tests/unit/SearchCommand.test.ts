import { SearchCommand } from '../../commands/SearchCommand';
import type { BotConfig } from '../../types';

describe('SearchCommand', () => {
  let searchCommand: SearchCommand;

  const mockConfig: BotConfig = {
    token: 'test-token',
    clientId: 'test-client-id',
    google: {
      spreadsheetId: 'test-spreadsheet-id',
      credentials: {
        project_id: 'test-project',
        private_key: 'test-private-key',
        client_email: 'test@example.com',
        client_id: 'test-client-id',
        auth_uri: 'https://accounts.google.com/o/oauth2/auth',
        token_uri: 'https://oauth2.googleapis.com/token',
        auth_provider_x509_cert_url: 'https://www.googleapis.com/oauth2/v1/certs',
        client_x509_cert_url: 'https://www.googleapis.com/robot/v1/metadata/x509/test%40example.com',
      },
    },
    ai: {
      openai: {
        apiKey: 'test-openai-key',
        model: 'gpt-3.5-turbo',
        maxTokens: 1000,
        temperature: 0.7,
      },
      ollama: {
        model: 'llama2',
      },
    },
    cache: {
      redis: {
        host: 'localhost',
        port: 6379,
        password: '',
        db: 0,
      },
      ttl: 3600,
    },
    metrics: {
      enabled: true,
      port: 9090,
      path: '/metrics',
    },
  };

  beforeEach(() => {
    searchCommand = new SearchCommand(mockConfig);
  });

  describe('constructor', () => {
    it('should create SearchCommand with correct configuration', () => {
      expect(searchCommand).toBeInstanceOf(SearchCommand);
    });
  });

  describe('utility methods', () => {
    it('should generate cache key correctly', () => {
      const params = {
        query: 'тест',
        documentType: 'all',
        limit: 20,
      };

      const cacheKey = (searchCommand as any).generateCacheKey(params);
      expect(cacheKey).toMatch(/^search:/);
      expect(cacheKey).toContain('base64');
    });

    it('should parse date correctly', () => {
      const date = (searchCommand as any).parseDate('01.01.2024');
      expect(date).toBeInstanceOf(Date);
      expect(date?.getFullYear()).toBe(2024);
    });
  });
});