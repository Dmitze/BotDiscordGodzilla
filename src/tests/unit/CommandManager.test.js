/**
 * Unit тести для CommandManager
 * Оновлено: 28.07.2025
 */

const { jest } = require('@jest/globals');
const CommandManager = require('../../core/CommandManager');

// Мокаємо залежності
jest.mock('fs', () => ({
  promises: {
    readdir: jest.fn(),
  },
}));

jest.mock('path', () => ({
  join: jest.fn(),
}));

jest.mock('../../utils/logger', () => ({
  info: jest.fn(),
  error: jest.fn(),
  warn: jest.fn(),
  debug: jest.fn(),
}));

describe('CommandManager', () => {
  let commandManager;
  let mockBot;

  beforeEach(() => {
    mockBot = {
      client: {
        on: jest.fn(),
      },
      config: {
        discord: { token: 'test-token' },
      },
    };

    commandManager = new CommandManager(mockBot);
  });

  afterEach(() => {
    jest.clearAllMocks();
  });

  describe('Constructor', () => {
    test('should initialize with correct properties', () => {
      expect(commandManager.bot).toBe(mockBot);
      expect(commandManager.commands).toBeDefined();
      expect(commandManager.commandHandlers).toBeDefined();
      expect(commandManager.commandCategories).toBeDefined();
    });
  });

  describe('initialize', () => {
    test('should initialize successfully', async () => {
      const mockFs = require('fs');
      const mockPath = require('path');

      mockFs.promises.readdir.mockResolvedValue(['test-command.js']);
      mockPath.join.mockReturnValue('/test/path');

      // Мокаємо команду
      const mockCommand = {
        data: { name: 'test' },
        execute: jest.fn(),
      };

      jest.doMock('/test/path', () => mockCommand, { virtual: true });

      await commandManager.initialize();

      expect(mockFs.promises.readdir).toHaveBeenCalled();
      expect(commandManager.commands.size).toBeGreaterThan(0);
    });

    test('should handle initialization errors', async () => {
      const mockFs = require('fs');
      mockFs.promises.readdir.mockRejectedValue(new Error('Read error'));

      await expect(commandManager.initialize()).rejects.toThrow('Read error');
    });
  });

  describe('loadCommands', () => {
    test('should load valid commands', async () => {
      const mockFs = require('fs');
      const mockPath = require('path');

      mockFs.promises.readdir.mockResolvedValue(['valid-command.js']);
      mockPath.join.mockReturnValue('/test/path');

      const mockCommand = {
        data: { name: 'valid' },
        execute: jest.fn(),
      };

      jest.doMock('/test/path', () => mockCommand, { virtual: true });

      await commandManager.loadCommands();

      expect(commandManager.commands.has('valid')).toBe(true);
    });

    test('should skip invalid commands', async () => {
      const mockFs = require('fs');
      const mockPath = require('path');

      mockFs.promises.readdir.mockResolvedValue(['invalid-command.js']);
      mockPath.join.mockReturnValue('/test/path');

      // Мокаємо невалідну команду
      const mockInvalidCommand = {
        data: {}, // Відсутня назва
        execute: jest.fn(),
      };

      jest.doMock('/test/path', () => mockInvalidCommand, { virtual: true });

      await commandManager.loadCommands();

      expect(commandManager.commands.size).toBe(0);
    });
  });

  describe('validateCommand', () => {
    test('should validate correct command', () => {
      const validCommand = {
        data: { name: 'test' },
        execute: jest.fn(),
      };

      const result = commandManager.validateCommand(validCommand);

      expect(result).toBe(true);
    });

    test('should reject command without data', () => {
      const invalidCommand = {
        execute: jest.fn(),
      };

      const result = commandManager.validateCommand(invalidCommand);

      expect(result).toBe(false);
    });

    test('should reject command without execute', () => {
      const invalidCommand = {
        data: { name: 'test' },
      };

      const result = commandManager.validateCommand(invalidCommand);

      expect(result).toBe(false);
    });

    test('should reject command without name', () => {
      const invalidCommand = {
        data: {},
        execute: jest.fn(),
      };

      const result = commandManager.validateCommand(invalidCommand);

      expect(result).toBe(false);
    });
  });

  describe('getCommandCategory', () => {
    test('should categorize search commands', () => {
      const command = { data: { name: 'пошук' } };
      const category = commandManager.getCommandCategory(command);

      expect(category).toBe('search');
    });

    test('should categorize document commands', () => {
      const command = { data: { name: 'документи' } };
      const category = commandManager.getCommandCategory(command);

      expect(category).toBe('documents');
    });

    test('should categorize AI commands', () => {
      const command = { data: { name: 'ai' } };
      const category = commandManager.getCommandCategory(command);

      expect(category).toBe('ai');
    });

    test('should return general for unknown commands', () => {
      const command = { data: { name: 'unknown' } };
      const category = commandManager.getCommandCategory(command);

      expect(category).toBe('general');
    });
  });

  describe('handleCommand', () => {
    test('should handle valid command', async () => {
      const mockInteraction = {
        isChatInputCommand: () => true,
        commandName: 'test',
        user: { tag: 'testuser' },
        reply: jest.fn(),
      };

      const mockCommand = {
        execute: jest.fn().mockResolvedValue(),
      };

      commandManager.commands.set('test', mockCommand);

      await commandManager.handleCommand(mockInteraction);

      expect(mockCommand.execute).toHaveBeenCalledWith(mockInteraction, commandManager.bot);
    });

    test('should handle unknown command', async () => {
      const mockInteraction = {
        isChatInputCommand: () => true,
        commandName: 'unknown',
        reply: jest.fn(),
      };

      await commandManager.handleCommand(mockInteraction);

      expect(mockInteraction.reply).toHaveBeenCalledWith({
        content: '❌ Команда не знайдена',
        ephemeral: true,
      });
    });

    test('should handle command execution errors', async () => {
      const mockInteraction = {
        isChatInputCommand: () => true,
        commandName: 'test',
        user: { tag: 'testuser' },
        reply: jest.fn(),
        deferred: false,
      };

      const mockCommand = {
        execute: jest.fn().mockRejectedValue(new Error('Command error')),
      };

      commandManager.commands.set('test', mockCommand);

      await commandManager.handleCommand(mockInteraction);

      expect(mockInteraction.reply).toHaveBeenCalledWith({
        content: '❌ Помилка виконання команди. Спробуйте ще раз.',
        ephemeral: true,
      });
    });

    test('should handle permission errors', async () => {
      const mockInteraction = {
        isChatInputCommand: () => true,
        commandName: 'test',
        user: { tag: 'testuser' },
        reply: jest.fn(),
        guild: { id: '123' },
        member: {
          roles: { cache: { some: jest.fn().mockReturnValue(false) } },
        },
      };

      const mockCommand = {
        permissions: { roles: ['Admin'] },
        execute: jest.fn(),
      };

      commandManager.commands.set('test', mockCommand);

      await commandManager.handleCommand(mockInteraction);

      expect(mockInteraction.reply).toHaveBeenCalledWith({
        content: '❌ У вас немає прав для використання цієї команди',
        ephemeral: true,
      });
    });
  });

  describe('checkPermissions', () => {
    test('should check role permissions', () => {
      const mockInteraction = {
        guild: { id: '123' },
        member: {
          roles: {
            cache: {
              some: jest.fn().mockReturnValue(true),
            },
          },
        },
      };

      const permissions = { roles: ['Admin'] };

      const result = commandManager.checkPermissions(mockInteraction, permissions);

      expect(result).toBe(true);
    });

    test('should check Discord permissions', () => {
      const mockInteraction = {
        guild: { id: '123' },
        member: {
          permissions: {
            has: jest.fn().mockReturnValue(true),
          },
        },
      };

      const permissions = { permissions: ['SendMessages'] };

      const result = commandManager.checkPermissions(mockInteraction, permissions);

      expect(result).toBe(true);
    });

    test('should return false for missing guild', () => {
      const mockInteraction = {
        guild: null,
      };

      const permissions = { roles: ['Admin'] };

      const result = commandManager.checkPermissions(mockInteraction, permissions);

      expect(result).toBe(false);
    });

    test('should return false for missing member', () => {
      const mockInteraction = {
        guild: { id: '123' },
        member: null,
      };

      const permissions = { roles: ['Admin'] };

      const result = commandManager.checkPermissions(mockInteraction, permissions);

      expect(result).toBe(false);
    });
  });

  describe('getCommand', () => {
    test('should return command by name', () => {
      const mockCommand = { name: 'test', execute: jest.fn() };
      commandManager.commands.set('test', mockCommand);

      const result = commandManager.getCommand('test');

      expect(result).toBe(mockCommand);
    });

    test('should return undefined for unknown command', () => {
      const result = commandManager.getCommand('unknown');

      expect(result).toBeUndefined();
    });
  });

  describe('getAllCommands', () => {
    test('should return all commands', () => {
      const mockCommand1 = { name: 'test1', execute: jest.fn() };
      const mockCommand2 = { name: 'test2', execute: jest.fn() };

      commandManager.commands.set('test1', mockCommand1);
      commandManager.commands.set('test2', mockCommand2);

      const result = commandManager.getAllCommands();

      expect(result).toHaveLength(2);
      expect(result).toContain(mockCommand1);
      expect(result).toContain(mockCommand2);
    });
  });

  describe('getCommandsByCategory', () => {
    test('should return commands by category', () => {
      const mockCommand = { name: 'test', execute: jest.fn() };
      commandManager.commands.set('test', mockCommand);
      commandManager.commandCategories.set('search', ['test']);

      const result = commandManager.getCommandsByCategory('search');

      expect(result).toHaveLength(1);
      expect(result[0]).toBe(mockCommand);
    });

    test('should return empty array for unknown category', () => {
      const result = commandManager.getCommandsByCategory('unknown');

      expect(result).toHaveLength(0);
    });
  });

  describe('getCategories', () => {
    test('should return all categories', () => {
      commandManager.commandCategories.set('search', []);
      commandManager.commandCategories.set('documents', []);

      const result = commandManager.getCategories();

      expect(result).toHaveLength(2);
      expect(result).toContain('search');
      expect(result).toContain('documents');
    });
  });

  describe('getStats', () => {
    test('should return command statistics', () => {
      commandManager.commands.set('test1', {});
      commandManager.commands.set('test2', {});
      commandManager.commandCategories.set('search', ['test1']);
      commandManager.commandCategories.set('documents', ['test2']);

      const stats = commandManager.getStats();

      expect(stats.total).toBe(2);
      expect(stats.categories).toBe(2);
      expect(stats.byCategory.search).toBe(1);
      expect(stats.byCategory.documents).toBe(1);
    });
  });
});
