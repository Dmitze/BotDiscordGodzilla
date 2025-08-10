"use strict";
/**
 * Утилиты для тестирования
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.createMockConfig = createMockConfig;
exports.createMockInteraction = createMockInteraction;
exports.createMockSheetData = createMockSheetData;
exports.wait = wait;
exports.clearMocks = clearMocks;
exports.expectFunctionCalled = expectFunctionCalled;
exports.expectFunctionCalledWith = expectFunctionCalledWith;
const globals_1 = require("@jest/globals");
/**
 * Создание мок конфигурации для тестов
 */
function createMockConfig() {
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
function createMockInteraction() {
    return {
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
            getString: globals_1.jest.fn(),
            getInteger: globals_1.jest.fn(),
            getBoolean: globals_1.jest.fn(),
            getSubcommand: globals_1.jest.fn(),
        },
        reply: globals_1.jest.fn(),
        editReply: globals_1.jest.fn(),
        followUp: globals_1.jest.fn(),
        deferReply: globals_1.jest.fn(),
        replied: false,
        deferred: false,
        isCommand: () => true,
        client: {
            serviceContainer: {
                get: globals_1.jest.fn(),
            },
        },
    };
}
/**
 * Создание мок Google Sheets данных
 */
function createMockSheetData() {
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
function wait(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}
/**
 * Очистка моков
 */
function clearMocks() {
    globals_1.jest.clearAllMocks();
}
/**
 * Проверка вызова функции
 */
function expectFunctionCalled(fn, times = 1) {
    (0, globals_1.expect)(fn).toHaveBeenCalledTimes(times);
}
/**
 * Проверка вызова функции с параметрами
 */
function expectFunctionCalledWith(fn, ...args) {
    (0, globals_1.expect)(fn).toHaveBeenCalledWith(...args);
}
//# sourceMappingURL=testHelpers.js.map