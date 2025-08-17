/**
 * Інтеграційні тести для команд
 * Оновлено: 28.07.2025
 */

const { describe, test, expect, beforeEach, afterEach } = require('@jest/globals');
const { SlashCommandBuilder } = require('discord.js');

// Мокаємо Discord.js
jest.mock('discord.js', () => ({
  SlashCommandBuilder: jest.fn().mockImplementation(() => ({
    setName: jest.fn().mockReturnThis(),
    setDescription: jest.fn().mockReturnThis(),
    addStringOption: jest.fn().mockReturnThis(),
    addIntegerOption: jest.fn().mockReturnThis(),
    addSubcommand: jest.fn().mockReturnThis(),
  })),
  EmbedBuilder: jest.fn().mockImplementation(() => ({
    setColor: jest.fn().mockReturnThis(),
    setTitle: jest.fn().mockReturnThis(),
    setDescription: jest.fn().mockReturnThis(),
    addFields: jest.fn().mockReturnThis(),
    setTimestamp: jest.fn().mockReturnThis(),
  })),
  ActionRowBuilder: jest.fn().mockImplementation(() => ({
    addComponents: jest.fn().mockReturnThis(),
  })),
  ButtonBuilder: jest.fn().mockImplementation(() => ({
    setCustomId: jest.fn().mockReturnThis(),
    setLabel: jest.fn().mockReturnThis(),
    setStyle: jest.fn().mockReturnThis(),
    setDisabled: jest.fn().mockReturnThis(),
  })),
  ButtonStyle: {
    Primary: 'primary',
    Secondary: 'secondary',
    Danger: 'danger',
  },
}));

// Мокаємо сервіси
jest.mock('../../services/AIService', () => {
  return jest.fn().mockImplementation(() => ({
    initialize: jest.fn().mockResolvedValue(),
    generateResponse: jest.fn().mockResolvedValue('AI response'),
    isActive: () => true,
  }));
});

jest.mock('../../services/GoogleService', () => {
  return jest.fn().mockImplementation(() => ({
    initialize: jest.fn().mockResolvedValue(),
    getSheetData: jest.fn().mockResolvedValue([
      ['Назва', 'Опис', 'Тип'],
      ['Тест 1', 'Опис 1', 'orders'],
      ['Тест 2', 'Опис 2', 'reports'],
    ]),
    isActive: () => true,
  }));
});

jest.mock('../../services/CacheService', () => {
  return jest.fn().mockImplementation(() => ({
    initialize: jest.fn().mockResolvedValue(),
    get: jest.fn().mockResolvedValue(null),
    set: jest.fn().mockResolvedValue(),
    isActive: () => true,
  }));
});

jest.mock('../../services/MetricsService', () => {
  return jest.fn().mockImplementation(() => ({
    initialize: jest.fn().mockResolvedValue(),
    incrementCommand: jest.fn(),
    measureCommandDuration: jest.fn(),
    isActive: () => true,
  }));
});

jest.mock('../../utils/logger', () => ({
  info: jest.fn(),
  error: jest.fn(),
  warn: jest.fn(),
  debug: jest.fn(),
}));

describe('Commands Integration Tests', () => {
  let mockBot;
  let mockInteraction;

  beforeEach(() => {
    // Створюємо моки сервісів ОДИН раз на тест і повертаємо ті ж самі інстанси
    const services = {
      ai: require('../../services/AIService')(),
      google: require('../../services/GoogleService')(),
      cache: require('../../services/CacheService')(),
      metrics: require('../../services/MetricsService')(),
    };

    // Створюємо мок бота
    mockBot = {
      getService: jest.fn(name => services[name]),
      handleError: jest.fn().mockResolvedValue({
        handled: true,
        message: 'Error handled',
      }),
    };

    // Створюємо мок interaction
    mockInteraction = {
      options: {
        getString: jest.fn(),
        getInteger: jest.fn(),
        getSubcommand: jest.fn(),
      },
      user: {
        tag: 'testuser#1234',
        id: '123456789',
      },
      deferReply: jest.fn().mockResolvedValue(),
      editReply: jest.fn().mockResolvedValue(),
      reply: jest.fn().mockResolvedValue(),
      deferred: false,
      replied: false,
    };
  });

  afterEach(() => {
    jest.clearAllMocks();
  });

  describe('Search Command', () => {
    let searchCommand;

    beforeEach(async () => {
      const SearchCommand = require('../../commands/search');
      searchCommand = SearchCommand;
    });

    test('should execute search command successfully', async () => {
      // Налаштовуємо моки
      mockInteraction.options.getString
        .mockReturnValueOnce('test query') // запит
        .mockReturnValueOnce('all') // тип_документа
        .mockReturnValueOnce(null) // дата_від
        .mockReturnValueOnce(null) // дата_до
        .mockReturnValueOnce(null) // підрозділ
        .mockReturnValueOnce('all') // пріоритет
        .mockReturnValueOnce(null); // ліміт

      mockInteraction.options.getInteger.mockReturnValueOnce(20); // ліміт

      // Виконуємо команду
      await searchCommand.execute(mockInteraction, mockBot);

      // Перевіряємо результати
      expect(mockInteraction.deferReply).toHaveBeenCalled();
      expect(mockInteraction.editReply).toHaveBeenCalled();

      const replyCall = mockInteraction.editReply.mock.calls[0][0];
      expect(replyCall.embeds).toBeDefined();
      expect(replyCall.embeds[0].data.title).toBe('🔍 Результати пошуку');
    });

    test('should handle search validation errors', async () => {
      // Налаштовуємо невалідний запит
      mockInteraction.options.getString
        .mockReturnValueOnce('') // порожній запит
        .mockReturnValueOnce('all');

      mockInteraction.options.getInteger.mockReturnValueOnce(20);

      // Виконуємо команду
      await searchCommand.execute(mockInteraction, mockBot);

      // Перевіряємо, що повернулася помилка валідації
      expect(mockInteraction.editReply).toHaveBeenCalledWith({
        content: expect.stringContaining('Помилка валідації'),
        ephemeral: true,
      });
    });

    test('should handle search service errors', async () => {
      // Налаштовуємо моки
      mockInteraction.options.getString
        .mockReturnValueOnce('test query')
        .mockReturnValueOnce('all');

      mockInteraction.options.getInteger.mockReturnValueOnce(20);

      // Мокаємо помилку сервісу
      const mockGoogleService = mockBot.getService('google');
      mockGoogleService.getSheetData.mockRejectedValue(new Error('Service error'));

      // Виконуємо команду
      await searchCommand.execute(mockInteraction, mockBot);

      // Перевіряємо, що помилка оброблена
      expect(mockBot.handleError).toHaveBeenCalled();
    });
  });

  describe('AI Assistant Command', () => {
    let aiCommand;

    beforeEach(async () => {
      const AICommand = require('../../commands/aiAssistant');
      aiCommand = AICommand;
    });

    test('should execute AI command successfully', async () => {
      // Налаштовуємо моки
      mockInteraction.options.getString
        .mockReturnValueOnce('Test AI query') // запит
        .mockReturnValueOnce(null) // контекст
        .mockReturnValueOnce('general'); // режим

      // Виконуємо команду
      await aiCommand.execute(mockInteraction, mockBot);

      // Перевіряємо результати
      expect(mockInteraction.deferReply).toHaveBeenCalled();
      expect(mockInteraction.editReply).toHaveBeenCalled();

      const replyCall = mockInteraction.editReply.mock.calls[0][0];
      expect(replyCall.embeds).toBeDefined();
      expect(replyCall.embeds[0].data.title).toBe('🤖 AI Відповідь');
    });

    test('should handle AI service errors', async () => {
      // Налаштовуємо моки
      mockInteraction.options.getString
        .mockReturnValueOnce('Test AI query')
        .mockReturnValueOnce(null)
        .mockReturnValueOnce('general');

      // Мокаємо помилку AI сервісу
      const mockAIService = mockBot.getService('ai');
      mockAIService.generateResponse.mockRejectedValue(new Error('AI service error'));

      // Виконуємо команду
      await aiCommand.execute(mockInteraction, mockBot);

      // Перевіряємо, що помилка оброблена
      expect(mockBot.handleError).toHaveBeenCalled();
    });

    test('should handle different AI modes', async () => {
      const modes = ['general', 'analysis', 'report', 'explanation', 'recommendations'];

      for (const mode of modes) {
        jest.clearAllMocks();

        mockInteraction.options.getString
          .mockReturnValueOnce('Test query')
          .mockReturnValueOnce(null)
          .mockReturnValueOnce(mode);

        await aiCommand.execute(mockInteraction, mockBot);

        expect(mockInteraction.editReply).toHaveBeenCalled();
      }
    });
  });

  describe('Documents Command', () => {
    let documentsCommand;

    beforeEach(async () => {
      const DocumentsCommand = require('../../commands/documents');
      documentsCommand = DocumentsCommand;
    });

    test('should execute documents command with search action', async () => {
      // Налаштовуємо моки
      mockInteraction.options.getSubcommand.mockReturnValue('особовий-склад');
      mockInteraction.options.getString
        .mockReturnValueOnce('search') // дія
        .mockReturnValueOnce('test query'); // запит

      // Виконуємо команду
      await documentsCommand.execute(mockInteraction, mockBot);

      // Перевіряємо результати
      expect(mockInteraction.deferReply).toHaveBeenCalled();
      expect(mockInteraction.editReply).toHaveBeenCalled();

      const replyCall = mockInteraction.editReply.mock.calls[0][0];
      expect(replyCall.embeds).toBeDefined();
      expect(replyCall.embeds[0].data.title).toContain('👥');
    });

    test('should handle different document subcommands', async () => {
      const subcommands = ['особовий-склад', 'техніка', 'матеріали', 'операції', 'накази'];

      for (const subcommand of subcommands) {
        jest.clearAllMocks();

        mockInteraction.options.getSubcommand.mockReturnValue(subcommand);
        mockInteraction.options.getString
          .mockReturnValueOnce('search')
          .mockReturnValueOnce('test query');

        await documentsCommand.execute(mockInteraction, mockBot);

        expect(mockInteraction.editReply).toHaveBeenCalled();
      }
    });

    test('should handle different document actions', async () => {
      const actions = ['search', 'add', 'update', 'report', 'status'];

      for (const action of actions) {
        jest.clearAllMocks();

        mockInteraction.options.getSubcommand.mockReturnValue('особовий-склад');
        mockInteraction.options.getString
          .mockReturnValueOnce(action)
          .mockReturnValueOnce('test query');

        await documentsCommand.execute(mockInteraction, mockBot);

        expect(mockInteraction.editReply).toHaveBeenCalled();
      }
    });

    test('should handle validation errors', async () => {
      // Налаштовуємо невалідні дані
      mockInteraction.options.getSubcommand.mockReturnValue('особовий-склад');
      mockInteraction.options.getString
        .mockReturnValueOnce('search')
        .mockReturnValueOnce('a'.repeat(1001)); // занадто довгий запит

      // Виконуємо команду
      await documentsCommand.execute(mockInteraction, mockBot);

      // Перевіряємо, що повернулася помилка валідації
      expect(mockInteraction.editReply).toHaveBeenCalledWith({
        content: expect.stringContaining('Помилка валідації'),
        ephemeral: true,
      });
    });
  });

  describe('Command Error Handling', () => {
    test('should handle general command errors', async () => {
      const searchCommand = require('../../commands/search');

      // Налаштовуємо моки
      mockInteraction.options.getString
        .mockReturnValueOnce('test query')
        .mockReturnValueOnce('all');

      mockInteraction.options.getInteger.mockReturnValueOnce(20);

      // Мокаємо помилку в deferReply
      mockInteraction.deferReply.mockRejectedValue(new Error('Network error'));

      // Виконуємо команду
      await searchCommand.execute(mockInteraction, mockBot);

      // Перевіряємо, що помилка оброблена
      expect(mockBot.handleError).toHaveBeenCalled();
    });

    test('should handle service unavailability', async () => {
      const searchCommand = require('../../commands/search');

      // Налаштовуємо моки
      mockInteraction.options.getString
        .mockReturnValueOnce('test query')
        .mockReturnValueOnce('all');

      mockInteraction.options.getInteger.mockReturnValueOnce(20);

      // Мокаємо недоступність сервісу
      mockBot.getService.mockReturnValue(null);

      // Виконуємо команду
      await searchCommand.execute(mockInteraction, mockBot);

      // Перевіряємо, що помилка оброблена
      expect(mockBot.handleError).toHaveBeenCalled();
    });
  });

  describe('Command Performance', () => {
    test('should measure command execution time', async () => {
      const searchCommand = require('../../commands/search');
      const mockMetricsService = mockBot.getService('metrics');

      // Налаштовуємо моки
      mockInteraction.options.getString
        .mockReturnValueOnce('test query')
        .mockReturnValueOnce('all');

      mockInteraction.options.getInteger.mockReturnValueOnce(20);

      // Виконуємо команду
      await searchCommand.execute(mockInteraction, mockBot);

      // Перевіряємо, що метрики оновлені
      expect(mockMetricsService.incrementCommand).toHaveBeenCalledWith('search', 'success');
      expect(mockMetricsService.measureCommandDuration).toHaveBeenCalledWith(
        'search',
        expect.any(Number)
      );
    });

    test('should handle slow command execution', async () => {
      const searchCommand = require('../../commands/search');

      // Налаштовуємо моки
      mockInteraction.options.getString
        .mockReturnValueOnce('test query')
        .mockReturnValueOnce('all');

      mockInteraction.options.getInteger.mockReturnValueOnce(20);

      // Мокаємо повільний сервіс
      const mockGoogleService = mockBot.getService('google');
      mockGoogleService.getSheetData.mockImplementation(
        () => new Promise(resolve => setTimeout(resolve, 100))
      );

      const startTime = Date.now();
      await searchCommand.execute(mockInteraction, mockBot);
      const endTime = Date.now();

      // Перевіряємо, що команда виконана за розумний час
      expect(endTime - startTime).toBeLessThan(200);
    });
  });

  describe('Command Caching', () => {
    test('should use cached results when available', async () => {
      const searchCommand = require('../../commands/search');
      const mockCacheService = mockBot.getService('cache');

      // Налаштовуємо моки
      mockInteraction.options.getString
        .mockReturnValueOnce('test query')
        .mockReturnValueOnce('all');

      mockInteraction.options.getInteger.mockReturnValueOnce(20);

      // Мокаємо кешований результат
      mockCacheService.get.mockResolvedValue([
        { name: 'Cached Result', description: 'From cache' },
      ]);

      // Виконуємо команду
      await searchCommand.execute(mockInteraction, mockBot);

      // Перевіряємо, що кеш використано
      expect(mockCacheService.get).toHaveBeenCalled();

      // Перевіряємо, що Google сервіс не викликано
      const mockGoogleService = mockBot.getService('google');
      expect(mockGoogleService.getSheetData).not.toHaveBeenCalled();
    });

    test('should cache new results', async () => {
      const searchCommand = require('../../commands/search');
      const mockCacheService = mockBot.getService('cache');

      // Налаштовуємо моки
      mockInteraction.options.getString
        .mockReturnValueOnce('test query')
        .mockReturnValueOnce('all');

      mockInteraction.options.getInteger.mockReturnValueOnce(20);

      // Мокаємо відсутність кешу
      mockCacheService.get.mockResolvedValue(null);

      // Виконуємо команду
      await searchCommand.execute(mockInteraction, mockBot);

      // Перевіряємо, що результат закешовано
      expect(mockCacheService.set).toHaveBeenCalled();
    });
  });
});
