/**
 * Unit тесты для SearchCommand
 */

import { jest, describe, it, expect, beforeEach } from '@jest/globals';
import { SearchCommand } from '../../../commands/SearchCommand';
import { createMockConfig, createMockInteraction } from '../../utils/testHelpers';

describe('SearchCommand', () => {
  let searchCommand: SearchCommand;
  let mockConfig: any;
  let mockInteraction: any;

  beforeEach(() => {
    mockConfig = createMockConfig();
    searchCommand = new SearchCommand(mockConfig);
    mockInteraction = createMockInteraction();
  });

  describe('constructor', () => {
    it('should create SearchCommand instance', () => {
      expect(searchCommand).toBeInstanceOf(SearchCommand);
    });

    it('should have correct name', () => {
      expect(searchCommand.getName()).toBe('пошук');
    });

    it('should have correct description', () => {
      expect(searchCommand.getDescription()).toBe('🔍 Гнучкий пошук по документах');
    });
  });

  describe('getData', () => {
    it('should return SlashCommandBuilder', () => {
      const data = searchCommand.getData();
      expect(data).toBeDefined();
      expect(data.name).toBe('пошук');
    });
  });

  describe('execute', () => {
    it('should handle basic search', async () => {
      // Настройка моков
      const mockGoogleService = {
        searchData: (jest.fn() as any).mockResolvedValue([['test', 'data']]),
      };
      const mockCacheService = {
        get: (jest.fn() as any).mockResolvedValue(null),
        set: (jest.fn() as any).mockResolvedValue(true),
      };

      mockInteraction.client.serviceContainer.get
        .mockReturnValueOnce(mockGoogleService)
        .mockReturnValueOnce(mockCacheService);

      mockInteraction.options.getString.mockReturnValue('тест');

      // Выполнение
      await searchCommand.execute({ interaction: mockInteraction } as any);

      // Проверки
      // В новій реалізації getString може отримувати другий параметр (required)
      expect(mockInteraction.options.getString).toHaveBeenCalledWith('запит', expect.anything());
      expect(mockGoogleService.searchData).toHaveBeenCalled();
      // Дозволяємо як reply, так і editReply залежно від гілки виконання
      expect(mockInteraction.reply.mock.calls.length + mockInteraction.editReply.mock.calls.length).toBeGreaterThan(0);
    });

    it('should handle empty results', async () => {
      // Настройка моков с пустыми результатами
      const mockGoogleService = {
        searchData: (jest.fn() as any).mockResolvedValue([]),
      };
      const mockCacheService = {
        get: (jest.fn() as any).mockResolvedValue(null),
        set: (jest.fn() as any).mockResolvedValue(true),
      };

      mockInteraction.client.serviceContainer.get
        .mockReturnValueOnce(mockGoogleService)
        .mockReturnValueOnce(mockCacheService);

      mockInteraction.options.getString.mockReturnValue('неіснуючий');

      // Выполнение
      await searchCommand.execute({ interaction: mockInteraction } as any);

      // Проверки: тепер відповідь може йти через editReply з embed-повідомленням
      expect(mockInteraction.reply.mock.calls.length + mockInteraction.editReply.mock.calls.length).toBeGreaterThan(0);
    });

    it('should handle service errors', async () => {
      // Настройка моков с ошибкой
      const mockGoogleService = {
        searchData: (jest.fn() as any).mockRejectedValue(new Error('Service error')),
      };
      const mockCacheService = {
        get: (jest.fn() as any).mockResolvedValue(null),
        set: jest.fn(),
      };

      mockInteraction.client.serviceContainer.get
        .mockReturnValueOnce(mockGoogleService)
        .mockReturnValueOnce(mockCacheService);

      mockInteraction.options.getString.mockReturnValue('тест');

      // Выполнение
      await searchCommand.execute({ interaction: mockInteraction } as any);

      // Проверки: помилка може бути надіслана як embed через reply або editReply
      expect(mockInteraction.reply.mock.calls.length + mockInteraction.editReply.mock.calls.length).toBeGreaterThan(0);
    });
  });
}); 