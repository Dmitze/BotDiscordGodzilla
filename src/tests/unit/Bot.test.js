/**
 * Unit тести для основного класу Bot
 * Оновлено: 28.07.2025
 */

const { jest } = require('@jest/globals');
const Bot = require('../../core/Bot');

// Мокаємо залежності
jest.mock('discord.js', () => ({
  Client: jest.fn().mockImplementation(() => ({
    on: jest.fn(),
    once: jest.fn(),
    login: jest.fn(),
    destroy: jest.fn(),
    user: { tag: 'TestBot#1234' },
    isReady: true,
    uptime: 1000,
    guilds: { cache: { size: 5 } },
    users: { cache: { size: 100 } },
  })),
  GatewayIntentBits: {
    Guilds: 1,
    GuildMessages: 2,
    MessageContent: 4,
    GuildMessageReactions: 8,
  },
  Collection: jest.fn(),
}));

jest.mock('../../utils/logger', () => ({
  info: jest.fn(),
  error: jest.fn(),
  warn: jest.fn(),
  debug: jest.fn(),
}));

jest.mock('../../config/Config', () => {
  return jest.fn().mockImplementation(() => ({
    discord: { token: 'test-token' },
    redis: { enabled: false },
    isMetricsEnabled: () => false,
  }));
});

describe('Bot Class', () => {
  let bot;
  let mockCommandManager;
  let mockEventManager;
  let mockServiceManager;
  let mockErrorHandler;

  beforeEach(() => {
    // Мокаємо менеджери
    mockCommandManager = {
      initialize: jest.fn(),
      getCommand: jest.fn(),
      getAllCommands: jest.fn(),
      getStats: jest.fn(),
    };

    mockEventManager = {
      initialize: jest.fn(),
      registerEvent: jest.fn(),
      removeEvent: jest.fn(),
      getRegisteredEvents: jest.fn(),
    };

    mockServiceManager = {
      initialize: jest.fn(),
      getService: jest.fn(),
      hasService: jest.fn(),
      getAllServices: jest.fn(),
      getStats: jest.fn(),
      shutdown: jest.fn(),
    };

    mockErrorHandler = {
      handle: jest.fn(),
    };

    // Мокаємо конструктори менеджерів
    jest.doMock('../../core/CommandManager', () => {
      return jest.fn().mockImplementation(() => mockCommandManager);
    });

    jest.doMock('../../core/EventManager', () => {
      return jest.fn().mockImplementation(() => mockEventManager);
    });

    jest.doMock('../../core/ServiceManager', () => {
      return jest.fn().mockImplementation(() => mockServiceManager);
    });

    jest.doMock('../../core/ErrorHandler', () => {
      return jest.fn().mockImplementation(() => mockErrorHandler);
    });

    bot = new Bot();
  });

  afterEach(() => {
    jest.clearAllMocks();
  });

  describe('Constructor', () => {
    test('should initialize bot with correct properties', () => {
      expect(bot.config).toBeDefined();
      expect(bot.client).toBeNull();
      expect(bot.commands).toBeDefined();
      expect(bot.services).toBeDefined();
      expect(bot.isReady).toBe(false);
    });
  });

  describe('initialize', () => {
    test('should initialize bot successfully', async () => {
      // Мокаємо успішну ініціалізацію
      mockCommandManager.initialize.mockResolvedValue();
      mockEventManager.initialize.mockResolvedValue();
      mockServiceManager.initialize.mockResolvedValue();

      await bot.initialize();

      expect(mockCommandManager.initialize).toHaveBeenCalled();
      expect(mockEventManager.initialize).toHaveBeenCalled();
      expect(mockServiceManager.initialize).toHaveBeenCalled();
      expect(bot.isReady).toBe(true);
    });

    test('should handle initialization errors', async () => {
      const error = new Error('Initialization failed');
      mockCommandManager.initialize.mockRejectedValue(error);

      await expect(bot.initialize()).rejects.toThrow('Initialization failed');
      expect(bot.isReady).toBe(false);
    });
  });

  describe('initializeManagers', () => {
    test('should initialize all managers', async () => {
      mockCommandManager.initialize.mockResolvedValue();
      mockEventManager.initialize.mockResolvedValue();
      mockServiceManager.initialize.mockResolvedValue();

      await bot.initializeManagers();

      expect(bot.errorHandler).toBeDefined();
      expect(bot.serviceManager).toBeDefined();
      expect(bot.commandManager).toBeDefined();
      expect(bot.eventManager).toBeDefined();
    });
  });

  describe('connect', () => {
    test('should connect to Discord successfully', async () => {
      const mockClient = {
        once: jest.fn((event, callback) => {
          if (event === 'ready') {
            callback();
          }
        }),
        on: jest.fn(),
        login: jest.fn(),
      };

      bot.client = mockClient;

      await bot.connect();

      expect(mockClient.once).toHaveBeenCalledWith('ready', expect.any(Function));
      expect(mockClient.login).toHaveBeenCalledWith(bot.config.discord.token);
    });

    test('should handle connection errors', async () => {
      const mockClient = {
        once: jest.fn(),
        on: jest.fn((event, callback) => {
          if (event === 'error') {
            callback(new Error('Connection failed'));
          }
        }),
        login: jest.fn(),
      };

      bot.client = mockClient;

      await expect(bot.connect()).rejects.toThrow('Connection failed');
    });
  });

  describe('startServices', () => {
    test('should start services when metrics are enabled', async () => {
      bot.config.isMetricsEnabled = () => true;
      bot.config.redis = { enabled: true };

      await bot.startServices();

      expect(mockServiceManager.startMetrics).toHaveBeenCalled();
      expect(mockServiceManager.startCache).toHaveBeenCalled();
      expect(mockServiceManager.startScheduler).toHaveBeenCalled();
    });

    test('should handle service start errors', async () => {
      const error = new Error('Service start failed');
      mockServiceManager.startMetrics.mockRejectedValue(error);

      bot.config.isMetricsEnabled = () => true;

      await bot.startServices();

      // Повинно продовжити роботу навіть при помилці
      expect(mockServiceManager.startScheduler).toHaveBeenCalled();
    });
  });

  describe('getCommand', () => {
    test('should return command by name', () => {
      const mockCommand = { name: 'test', execute: jest.fn() };
      mockCommandManager.getCommand.mockReturnValue(mockCommand);

      const result = bot.getCommand('test');

      expect(result).toBe(mockCommand);
      expect(mockCommandManager.getCommand).toHaveBeenCalledWith('test');
    });
  });

  describe('getService', () => {
    test('should return service by name', () => {
      const mockService = { name: 'test', isActive: () => true };
      mockServiceManager.getService.mockReturnValue(mockService);

      const result = bot.getService('test');

      expect(result).toBe(mockService);
      expect(mockServiceManager.getService).toHaveBeenCalledWith('test');
    });
  });

  describe('handleError', () => {
    test('should handle errors through ErrorHandler', () => {
      const error = new Error('Test error');
      const context = { command: 'test' };

      bot.handleError(error, context);

      expect(mockErrorHandler.handle).toHaveBeenCalledWith(error, context);
    });
  });

  describe('shutdown', () => {
    test('should shutdown bot gracefully', async () => {
      bot.client = { destroy: jest.fn() };
      mockServiceManager.shutdown.mockResolvedValue();

      await bot.shutdown();

      expect(mockServiceManager.shutdown).toHaveBeenCalled();
      expect(bot.client.destroy).toHaveBeenCalled();
    });

    test('should handle shutdown errors', async () => {
      const error = new Error('Shutdown failed');
      mockServiceManager.shutdown.mockRejectedValue(error);

      await bot.shutdown();

      // Повинно продовжити роботу навіть при помилці
      expect(bot.client.destroy).toHaveBeenCalled();
    });
  });

  describe('getStats', () => {
    test('should return bot statistics', () => {
      bot.client = {
        uptime: 1000,
        guilds: { cache: { size: 5 } },
        users: { cache: { size: 100 } },
      };
      bot.commands = { size: 10 };
      bot.services = { ai: {}, google: {} };
      bot.isReady = true;

      const stats = bot.getStats();

      expect(stats).toEqual({
        uptime: 1000,
        guilds: 5,
        users: 100,
        commands: 10,
        services: 2,
        isReady: true,
      });
    });

    test('should handle missing client', () => {
      bot.client = null;

      const stats = bot.getStats();

      expect(stats.uptime).toBe(0);
      expect(stats.guilds).toBe(0);
      expect(stats.users).toBe(0);
    });
  });
});
