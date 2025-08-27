import { DocumentAccessAuditService } from '../DocumentAccessAuditService';
import type { BotConfig } from '@/types';

// Mock logger
jest.mock('@/utils/logger', () => ({
  __esModule: true,
  default: {
    info: jest.fn(),
    error: jest.fn(),
    warn: jest.fn(),
    debug: jest.fn(),
    security: jest.fn(),
    command: jest.fn(),
    commandError: jest.fn(),
    apiRequest: jest.fn(),
    apiError: jest.fn(),
    performance: jest.fn(),
    system: jest.fn(),
    getStats: jest.fn(),
    getLogBuffer: jest.fn(),
    cleanup: jest.fn(),
    isHealthy: jest.fn(),
    log: jest.fn(),
    logStructured: jest.fn(),
    startStructuredTimer: jest.fn()
  }
}));

describe('DocumentAccessAuditService', () => {
  let service: DocumentAccessAuditService;
  const mockConfig: BotConfig = {
    discord: {
      token: 'test-token',
      clientId: 'test-client-id',
      prefix: '!',
      intents: []
    },
    google: {
      spreadsheetId: 'test-spreadsheet-id',
      driveFolderId: 'test-folder-id',
      apiKey: 'test-api-key',
      applicationCredentials: 'test-credentials',
      appScriptUrl: 'test-script-url',
      sheetName: 'test-sheet'
    },
    ai: {
      provider: 'openai',
      openai: {
        apiKey: 'test-openai-key',
        model: 'gpt-3.5-turbo',
        maxTokens: 1000,
        temperature: 0.7
      },
      ollama: {
        host: 'http://localhost:11434',
        model: 'llama2'
      }
    },
    redis: {
      host: 'localhost',
      port: 6379,
      database: 0,
      enabled: false
    },
    metrics: {
      enabled: false,
      port: 9090,
      path: '/metrics'
    },
    security: {
      rateLimitWindow: 60000,
      rateLimitMax: 10,
      adminRole: 'admin',
      botUserRole: 'bot-user'
    },
    performance: {
      cacheTTL: 300,
      maxSearchResults: 50,
      maxAnalysisRows: 1000,
      requestTimeout: 30000,
      maxRetries: 3
    },
    logging: {
      level: 'info',
      maxFiles: 5,
      maxSize: '10m',
      directory: './logs'
    },
    drive: {
      folderId: 'test-folder-id'
    },
    features: {
      defaultLocale: 'en'
    }
  };

  beforeEach(() => {
    service = new DocumentAccessAuditService(mockConfig);
  });

  afterEach(() => {
    jest.clearAllMocks();
  });

  test('should initialize correctly', () => {
    expect(service).toBeInstanceOf(DocumentAccessAuditService);
  });

  test('should log document access', async () => {
    const accessLog = {
      userId: 'user123',
      userName: 'Test User',
      fileId: 'file456',
      fileName: 'test-document.pdf',
      accessType: 'view' as const,
      success: true,
      fileSize: 1024,
      fileType: 'application/pdf'
    };

    await service.logAccess(accessLog);

    // Verify the log was stored
    const logs = await service.getAccessLogs();
    expect(logs).toHaveLength(1);
    expect(logs[0].userId).toBe('user123');
    expect(logs[0].fileId).toBe('file456');
    expect(logs[0].accessType).toBe('view');
  });

  test('should retrieve access logs with filters', async () => {
    // Add multiple access logs
    await service.logAccess({
      userId: 'user1',
      userName: 'User One',
      fileId: 'file1',
      fileName: 'document1.pdf',
      accessType: 'view',
      success: true
    });

    await service.logAccess({
      userId: 'user1',
      userName: 'User One',
      fileId: 'file2',
      fileName: 'document2.pdf',
      accessType: 'edit',
      success: true
    });

    await service.logAccess({
      userId: 'user2',
      userName: 'User Two',
      fileId: 'file1',
      fileName: 'document1.pdf',
      accessType: 'view',
      success: false,
      errorMessage: 'Permission denied'
    });

    // Test filtering by userId
    const user1Logs = await service.getAccessLogs({ userId: 'user1' });
    expect(user1Logs).toHaveLength(2);
    expect(user1Logs.every(log => log.userId === 'user1')).toBe(true);

    // Test filtering by fileId
    const file1Logs = await service.getAccessLogs({ fileId: 'file1' });
    expect(file1Logs).toHaveLength(2);
    expect(file1Logs.every(log => log.fileId === 'file1')).toBe(true);

    // Test filtering by accessType
    const viewLogs = await service.getAccessLogs({ accessType: 'view' });
    expect(viewLogs).toHaveLength(2);
    expect(viewLogs.every(log => log.accessType === 'view')).toBe(true);
  });

  test('should calculate access statistics', async () => {
    // Add test data
    await service.logAccess({
      userId: 'user1',
      userName: 'User One',
      fileId: 'file1',
      fileName: 'document1.pdf',
      accessType: 'view',
      success: true
    });

    await service.logAccess({
      userId: 'user1',
      userName: 'User One',
      fileId: 'file2',
      fileName: 'document2.pdf',
      accessType: 'edit',
      success: true
    });

    await service.logAccess({
      userId: 'user2',
      userName: 'User Two',
      fileId: 'file1',
      fileName: 'document1.pdf',
      accessType: 'view',
      success: false,
      errorMessage: 'Permission denied'
    });

    const stats = await service.getAccessStats();
    
    expect(stats.totalAccesses).toBe(3);
    expect(stats.successfulAccesses).toBe(2);
    expect(stats.failedAccesses).toBe(1);
    expect(stats.uniqueUsers).toBe(2);
    expect(stats.accessByType['view']).toBe(2);
    expect(stats.accessByType['edit']).toBe(1);
  });

  test('should track user-file access relationships', async () => {
    await service.logAccess({
      userId: 'user1',
      userName: 'User One',
      fileId: 'file1',
      fileName: 'document1.pdf',
      accessType: 'view',
      success: true
    });

    await service.logAccess({
      userId: 'user1',
      userName: 'User One',
      fileId: 'file2',
      fileName: 'document2.pdf',
      accessType: 'view',
      success: true
    });

    await service.logAccess({
      userId: 'user2',
      userName: 'User Two',
      fileId: 'file1',
      fileName: 'document1.pdf',
      accessType: 'view',
      success: true
    });

    // Test files accessed by user
    const user1Files = service.getFilesAccessedByUser('user1');
    expect(user1Files).toHaveLength(2);
    expect(user1Files).toContain('file1');
    expect(user1Files).toContain('file2');

    // Test users who accessed file
    const file1Users = service.getUsersWhoAccessedFile('file1');
    expect(file1Users).toHaveLength(2);
    expect(file1Users).toContain('user1');
    expect(file1Users).toContain('user2');
  });

  test('should limit log history', async () => {
    // Add more logs than the maximum history (10000)
    // For testing purposes, we'll add just a few more than the default limit of getAccessLogs (50)
    for (let i = 0; i < 60; i++) {
      await service.logAccess({
        userId: `user${i}`,
        userName: `User ${i}`,
        fileId: `file${i}`,
        fileName: `document${i}.pdf`,
        accessType: 'view',
        success: true
      });
    }

    // Test with default limit (50)
    const logs = await service.getAccessLogs();
    expect(logs).toHaveLength(50);
    
    // Test with custom limit
    const limitedLogs = await service.getAccessLogs({ limit: 10 });
    expect(limitedLogs).toHaveLength(10);
    
    // Verify these are the most recent logs
    expect(limitedLogs[0].userId).toBe('user59');
    expect(limitedLogs[9].userId).toBe('user50');
  });

  test('should export access logs', async () => {
    await service.logAccess({
      userId: 'user1',
      userName: 'User One',
      fileId: 'file1',
      fileName: 'document1.pdf',
      accessType: 'view',
      success: true
    });

    const exportedData = await service.exportAccessLogs();
    expect(exportedData).toContain('exportedAt');
    expect(exportedData).toContain('logs');
    
    const parsedData = JSON.parse(exportedData);
    expect(parsedData.logs).toHaveLength(1);
    expect(parsedData.logs[0].userId).toBe('user1');
  });
});