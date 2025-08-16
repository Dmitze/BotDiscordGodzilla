/**
 * Утилиты для тестирования
 */

import { jest, expect } from '@jest/globals';

/**
 * Создание мок конфигурации для тестов
 */
export function createMockConfig(): any {
  return {
    discord: {
      token: 'test_token',
      clientId: 'test_client_id',
      guildId: 'test_guild_id',
      prefix: '!',
      intents: ['Guilds', 'GuildMessages', 'MessageContent'],
    },
    google: {
      spreadsheetId: 'test_spreadsheet_id',
      driveFolderId: 'test_drive_folder_id',
      apiKey: 'test_api_key',
      credentials: {
        project_id: 'test_project',
        private_key: 'test_private_key',
        client_email: 'test@test.com',
      },
      applicationCredentials: './test-credentials.json',
      appScriptUrl: 'https://script.google.com/macros/s/test/exec',
      sheetName: 'TestSheet',
    },
    ai: {
      provider: 'openai',
      openai: {
        apiKey: 'test_openai_key',
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
      enabled: false,
      host: 'localhost',
      port: 6379,
      password: '',
      db: 1,
    },
    security: {
      rateLimitEnabled: false,
      rateLimitWindow: 900000,
      rateLimitMax: 100,
      adminRole: 'TestAdmin',
      botUserRole: 'TestBotUser',
      sheetsAccessRole: 'TestSheetsAccess',
      aiAccessRole: 'TestAIAccess',
      exportAccessRole: 'TestExportAccess',
      securityLogLevel: 'error',
    },
    files: {
      maxFileSize: 1048576,
      exportMaxFileSize: 2097152,
      tempDir: './test-tmp',
      fileCleanupInterval: 3600000,
      downloadTimeout: 5000,
      tempFileTtl: 30000,
      includeMetadata: false,
    },
    metrics: {
      enabled: true,
      port: 9090,
      path: '/metrics',
    },
  };
}

/**
 * Создание мок Discord взаимодействия
 */
export function createMockInteraction() {
  const interaction: any = {
    commandName: 'test',
    user: {
      id: 'test_user_id',
      username: 'test_user',
      tag: 'test_user#1234',
    },
    guild: {
      id: 'test_guild_id',
      name: 'Test Guild',
    },
    channel: {
      id: 'test_channel_id',
      name: 'test-channel',
    },
    options: {
      getString: jest.fn(),
      getInteger: jest.fn(),
      getBoolean: jest.fn(),
      getSubcommand: jest.fn(),
    },
    reply: jest.fn(),
    editReply: jest.fn(),
    followUp: jest.fn(),
    deferReply: jest.fn().mockImplementation(() => {
      interaction.deferred = true;
      return Promise.resolve();
    }),
    replied: false,
    deferred: false,
    isCommand: () => true,
    client: {
      serviceContainer: {
        get: jest.fn(),
      },
    },
  };
  return interaction;
}

/**
 * Создание мок Google Sheets данных
 */
export function createMockSheetData() {
  return [
    ['ID', 'Назва', 'Ціна', 'Кількість', 'Дата'],
    ['1', 'Товар 1', '100', '10', '2024-01-01'],
    ['2', 'Товар 2', '200', '5', '2024-01-02'],
    ['3', 'Товар 3', '150', '15', '2024-01-03'],
  ];
}

/**
 * Ожидание асинхронной операции
 */
export function wait(ms: number): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * Очистка моков
 */
export function clearMocks() {
  jest.clearAllMocks();
}

/**
 * Проверка вызова функции
 */
export function expectFunctionCalled(fn: jest.Mock, times: number = 1) {
  expect(fn).toHaveBeenCalledTimes(times);
}

/**
 * Проверка вызова функции с параметрами
 */
export function expectFunctionCalledWith(fn: jest.Mock, ...args: any[]) {
  expect(fn).toHaveBeenCalledWith(...args);
}
