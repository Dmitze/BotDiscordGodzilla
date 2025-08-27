import { DocumentAuditCommand } from '../DocumentAuditCommand';
import { DocumentAccessAuditService } from '@/services/DocumentAccessAuditService';
import type { BotConfig } from '@/types';

// Mock Discord interaction
const createMockInteraction = (subcommand: string, options: Record<string, any> = {}) => ({
  options: {
    getSubcommand: () => subcommand,
    getString: (name: string) => options[name] || null,
    getInteger: (name: string) => options[name] || null
  },
  deferReply: jest.fn(),
  reply: jest.fn(),
  editReply: jest.fn()
});

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

describe('DocumentAuditCommand', () => {
  let command: DocumentAuditCommand;
  let auditService: DocumentAccessAuditService;
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
    command = new DocumentAuditCommand(mockConfig);
    auditService = new DocumentAccessAuditService(mockConfig);
    command.initializeServices(auditService);
  });

  afterEach(() => {
    jest.clearAllMocks();
  });

  test('should initialize correctly', () => {
    expect(command).toBeInstanceOf(DocumentAuditCommand);
  });

  test('should handle view logs subcommand', async () => {
    const interaction = createMockInteraction('view', { limit: 5 });
    
    // Add some test data
    await auditService.logAccess({
      userId: 'user123',
      userName: 'Test User',
      fileId: 'file456',
      fileName: 'test-document.pdf',
      accessType: 'view',
      success: true
    });

    await command.execute(interaction as any);

    expect(interaction.deferReply).toHaveBeenCalled();
    expect(interaction.editReply).toHaveBeenCalled();
    
    const editReplyCall = (interaction.editReply as jest.Mock).mock.calls[0][0];
    expect(editReplyCall.content).toContain('Document Access Logs');
    expect(editReplyCall.content).toContain('test-document.pdf');
  });

  test('should handle stats subcommand', async () => {
    const interaction = createMockInteraction('stats');
    
    // Add some test data
    await auditService.logAccess({
      userId: 'user123',
      userName: 'Test User',
      fileId: 'file456',
      fileName: 'test-document.pdf',
      accessType: 'view',
      success: true
    });

    await command.execute(interaction as any);

    expect(interaction.deferReply).toHaveBeenCalled();
    expect(interaction.editReply).toHaveBeenCalled();
    
    const editReplyCall = (interaction.editReply as jest.Mock).mock.calls[0][0];
    expect(editReplyCall.content).toContain('Document Access Statistics');
    expect(editReplyCall.content).toContain('Total Accesses: 1');
  });

  test('should handle export subcommand', async () => {
    const interaction = createMockInteraction('export');
    
    // Add some test data
    await auditService.logAccess({
      userId: 'user123',
      userName: 'Test User',
      fileId: 'file456',
      fileName: 'test-document.pdf',
      accessType: 'view',
      success: true
    });

    await command.execute(interaction as any);

    expect(interaction.deferReply).toHaveBeenCalled();
    expect(interaction.editReply).toHaveBeenCalled();
    
    const editReplyCall = (interaction.editReply as jest.Mock).mock.calls[0][0];
    expect(editReplyCall.content).toContain('Document access logs exported successfully');
  });

  test('should handle error when audit service is not initialized', async () => {
    const commandWithoutService = new DocumentAuditCommand(mockConfig);
    const interaction = createMockInteraction('view');
    
    await commandWithoutService.onExecute({ interaction: interaction as any });
    
    expect(interaction.reply).toHaveBeenCalled();
    const replyCall = (interaction.reply as jest.Mock).mock.calls[0][0];
    expect(replyCall.content).toContain('Audit service not initialized');
    expect(replyCall.ephemeral).toBe(true);
  });

  test('should handle unknown subcommand', async () => {
    // Mock the getSubcommand to return an unknown value
    const interaction = {
      options: {
        getSubcommand: () => 'unknown'
      },
      deferReply: jest.fn(),
      reply: jest.fn(),
      editReply: jest.fn()
    };
    
    await command.execute(interaction as any);
    
    expect(interaction.reply).toHaveBeenCalled();
    const replyCall = (interaction.reply as jest.Mock).mock.calls[0][0];
    expect(replyCall.content).toContain('Unknown subcommand');
    expect(replyCall.ephemeral).toBe(true);
  });

  test('should handle errors gracefully', async () => {
    const interaction = createMockInteraction('view');
    
    // Mock the audit service to throw an error
    jest.spyOn(auditService, 'getAccessLogs').mockRejectedValue(new Error('Test error'));
    
    await command.execute(interaction as any);
    
    expect(interaction.deferReply).toHaveBeenCalled();
    expect(interaction.editReply).toHaveBeenCalled();
    
    const editReplyCall = (interaction.editReply as jest.Mock).mock.calls[0][0];
    expect(editReplyCall.content).toContain('Failed to retrieve document access logs');
  });
});