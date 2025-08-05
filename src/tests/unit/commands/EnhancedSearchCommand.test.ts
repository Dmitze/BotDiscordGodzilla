/**
 * Unit тесты для EnhancedSearchCommand
 */

import { jest, describe, it, expect, beforeEach } from '@jest/globals';
import { EnhancedSearchCommand } from '../../../commands/EnhancedSearchCommand';
import { createMockConfig, createMockInteraction } from '../../utils/testHelpers';

describe('EnhancedSearchCommand', () => {
  let enhancedSearchCommand: EnhancedSearchCommand;
  let mockConfig: any;
  let mockInteraction: any;

  beforeEach(() => {
    mockConfig = createMockConfig();
    enhancedSearchCommand = new EnhancedSearchCommand(mockConfig);
    mockInteraction = createMockInteraction();
  });

  describe('constructor', () => {
    it('should create EnhancedSearchCommand instance', () => {
      expect(enhancedSearchCommand).toBeInstanceOf(EnhancedSearchCommand);
    });

    it('should have correct name', () => {
      expect(enhancedSearchCommand.getName()).toBe('розширений_пошук');
    });

    it('should have correct description', () => {
      expect(enhancedSearchCommand.getDescription()).toBe('Розширений пошук з фільтрами та сортуванням');
    });
  });

  describe('getData', () => {
    it('should return SlashCommandBuilder', () => {
      const data = enhancedSearchCommand.getData();
      expect(data).toBeDefined();
      expect(data.name).toBe('розширений_пошук');
    });
  });

  describe('execute', () => {
    it('should handle basic search with filters', async () => {
      // Настройка моков
      const mockGoogleService = {
        enhancedSearch: jest.fn().mockResolvedValue([
          { id: '1', name: 'Item 1', price: 100, category: 'electronics' },
          { id: '2', name: 'Item 2', price: 200, category: 'electronics' },
        ]),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockGoogleService);
      mockInteraction.options.getString.mockReturnValue('electronics');
      mockInteraction.options.getInteger.mockReturnValue(100);
      mockInteraction.options.getInteger.mockReturnValueOnce(100).mockReturnValueOnce(500);

      // Выполнение
      await enhancedSearchCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.options.getString).toHaveBeenCalledWith('номенклатура');
      expect(mockInteraction.options.getInteger).toHaveBeenCalledWith('ціна_від');
      expect(mockInteraction.options.getInteger).toHaveBeenCalledWith('ціна_до');
      expect(mockGoogleService.enhancedSearch).toHaveBeenCalled();
      expect(mockInteraction.reply).toHaveBeenCalled();
    });

    it('should handle search with sorting', async () => {
      // Настройка моков
      const mockGoogleService = {
        enhancedSearch: jest.fn().mockResolvedValue([
          { id: '1', name: 'Item 1', price: 100 },
          { id: '2', name: 'Item 2', price: 200 },
        ]),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockGoogleService);
      mockInteraction.options.getString.mockReturnValue('test');
      mockInteraction.options.getString.mockReturnValueOnce('test').mockReturnValueOnce('price').mockReturnValueOnce('desc');

      // Выполнение
      await enhancedSearchCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.options.getString).toHaveBeenCalledWith('запит');
      expect(mockInteraction.options.getString).toHaveBeenCalledWith('сортування');
      expect(mockInteraction.options.getString).toHaveBeenCalledWith('порядок');
      expect(mockGoogleService.enhancedSearch).toHaveBeenCalled();
      expect(mockInteraction.reply).toHaveBeenCalled();
    });

    it('should handle search with date filters', async () => {
      // Настройка моков
      const mockGoogleService = {
        enhancedSearch: jest.fn().mockResolvedValue([
          { id: '1', name: 'Item 1', date: '2024-01-01' },
          { id: '2', name: 'Item 2', date: '2024-01-02' },
        ]),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockGoogleService);
      mockInteraction.options.getString.mockReturnValue('test');
      mockInteraction.options.getString.mockReturnValueOnce('test').mockReturnValueOnce('2024-01-01').mockReturnValueOnce('2024-01-31');

      // Выполнение
      await enhancedSearchCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.options.getString).toHaveBeenCalledWith('запит');
      expect(mockInteraction.options.getString).toHaveBeenCalledWith('дата_від');
      expect(mockInteraction.options.getString).toHaveBeenCalledWith('дата_до');
      expect(mockGoogleService.enhancedSearch).toHaveBeenCalled();
      expect(mockInteraction.reply).toHaveBeenCalled();
    });

    it('should handle search with limit', async () => {
      // Настройка моков
      const mockGoogleService = {
        enhancedSearch: jest.fn().mockResolvedValue([
          { id: '1', name: 'Item 1' },
          { id: '2', name: 'Item 2' },
        ]),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockGoogleService);
      mockInteraction.options.getString.mockReturnValue('test');
      mockInteraction.options.getInteger.mockReturnValue(10);

      // Выполнение
      await enhancedSearchCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.options.getString).toHaveBeenCalledWith('запит');
      expect(mockInteraction.options.getInteger).toHaveBeenCalledWith('ліміт');
      expect(mockGoogleService.enhancedSearch).toHaveBeenCalled();
      expect(mockInteraction.reply).toHaveBeenCalled();
    });

    it('should handle empty search query', async () => {
      mockInteraction.options.getString.mockReturnValue('');

      // Выполнение
      await enhancedSearchCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.reply).toHaveBeenCalledWith(
        expect.objectContaining({
          content: expect.stringContaining('Будь ласка, вкажіть запит'),
          ephemeral: true,
        })
      );
    });

    it('should handle service error', async () => {
      // Настройка моков с ошибкой
      const mockGoogleService = {
        enhancedSearch: jest.fn().mockRejectedValue(new Error('Search service error')),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockGoogleService);
      mockInteraction.options.getString.mockReturnValue('test');

      // Выполнение
      await enhancedSearchCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.reply).toHaveBeenCalledWith(
        expect.objectContaining({
          content: expect.stringContaining('Помилка'),
          ephemeral: true,
        })
      );
    });

    it('should handle empty search results', async () => {
      // Настройка моков с пустыми результатами
      const mockGoogleService = {
        enhancedSearch: jest.fn().mockResolvedValue([]),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockGoogleService);
      mockInteraction.options.getString.mockReturnValue('неіснуючий');

      // Выполнение
      await enhancedSearchCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.reply).toHaveBeenCalledWith(
        expect.objectContaining({
          content: expect.stringContaining('Результатів не знайдено'),
          ephemeral: true,
        })
      );
    });

    it('should handle complex filters', async () => {
      // Настройка моков для сложных фильтров
      const mockGoogleService = {
        enhancedSearch: jest.fn().mockResolvedValue([
          { id: '1', name: 'Item 1', price: 150, category: 'electronics', date: '2024-01-15' },
        ]),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockGoogleService);
      mockInteraction.options.getString.mockReturnValue('electronics');
      mockInteraction.options.getInteger.mockReturnValue(100);
      mockInteraction.options.getInteger.mockReturnValueOnce(100).mockReturnValueOnce(200);
      mockInteraction.options.getString.mockReturnValueOnce('electronics').mockReturnValueOnce('2024-01-01').mockReturnValueOnce('2024-01-31');

      // Выполнение
      await enhancedSearchCommand.execute(mockInteraction);

      // Проверки
      expect(mockGoogleService.enhancedSearch).toHaveBeenCalledWith(
        expect.objectContaining({
          query: 'electronics',
          priceFrom: 100,
          priceTo: 200,
          dateFrom: '2024-01-01',
          dateTo: '2024-01-31',
        })
      );
      expect(mockInteraction.reply).toHaveBeenCalled();
    });

    it('should handle pagination', async () => {
      // Настройка моков для пагинации
      const mockGoogleService = {
        enhancedSearch: jest.fn().mockResolvedValue({
          data: [
            { id: '1', name: 'Item 1' },
            { id: '2', name: 'Item 2' },
          ],
          total: 50,
          page: 1,
          totalPages: 5,
        }),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockGoogleService);
      mockInteraction.options.getString.mockReturnValue('test');
      mockInteraction.options.getInteger.mockReturnValue(1);

      // Выполнение
      await enhancedSearchCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.options.getInteger).toHaveBeenCalledWith('сторінка');
      expect(mockGoogleService.enhancedSearch).toHaveBeenCalled();
      expect(mockInteraction.reply).toHaveBeenCalledWith(
        expect.objectContaining({
          content: expect.stringContaining('Сторінка 1 з 5'),
        })
      );
    });
  });
}); 