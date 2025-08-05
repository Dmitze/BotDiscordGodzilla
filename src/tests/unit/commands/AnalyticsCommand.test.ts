/**
 * Unit тесты для AnalyticsCommand
 */

import { jest, describe, it, expect, beforeEach } from '@jest/globals';
import { AnalyticsCommand } from '../../../commands/AnalyticsCommand';
import { createMockConfig, createMockInteraction } from '../../utils/testHelpers';

describe('AnalyticsCommand', () => {
  let analyticsCommand: AnalyticsCommand;
  let mockConfig: any;
  let mockInteraction: any;

  beforeEach(() => {
    mockConfig = createMockConfig();
    analyticsCommand = new AnalyticsCommand(mockConfig);
    mockInteraction = createMockInteraction();
  });

  describe('constructor', () => {
    it('should create AnalyticsCommand instance', () => {
      expect(analyticsCommand).toBeInstanceOf(AnalyticsCommand);
    });

    it('should have correct name', () => {
      expect(analyticsCommand.getName()).toBe('аналітика');
    });

    it('should have correct description', () => {
      expect(analyticsCommand.getDescription()).toBe('Аналітика та звітність');
    });
  });

  describe('getData', () => {
    it('should return SlashCommandBuilder', () => {
      const data = analyticsCommand.getData();
      expect(data).toBeDefined();
      expect(data.name).toBe('аналітика');
    });
  });

  describe('execute', () => {
    it('should handle report subcommand', async () => {
      // Настройка моков
      const mockAnalyticsService = {
        generateReport: jest.fn().mockResolvedValue({
          type: 'daily',
          data: { users: 100, commands: 500 },
          timestamp: new Date(),
        }),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockAnalyticsService);
      mockInteraction.options.getSubcommand.mockReturnValue('звіт');
      mockInteraction.options.getString.mockReturnValue('daily');
      mockInteraction.options.getString.mockReturnValueOnce('daily').mockReturnValueOnce('excel');

      // Выполнение
      await analyticsCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.options.getSubcommand).toHaveBeenCalled();
      expect(mockInteraction.options.getString).toHaveBeenCalledWith('тип');
      expect(mockInteraction.options.getString).toHaveBeenCalledWith('формат');
      expect(mockAnalyticsService.generateReport).toHaveBeenCalledWith('daily', 'excel');
      expect(mockInteraction.reply).toHaveBeenCalled();
    });

    it('should handle statistics subcommand', async () => {
      // Настройка моков
      const mockAnalyticsService = {
        getStatistics: jest.fn().mockResolvedValue({
          totalUsers: 1000,
          totalCommands: 5000,
          popularCommands: ['пошук', 'ai_асистент'],
          dailyActiveUsers: 150,
        }),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockAnalyticsService);
      mockInteraction.options.getSubcommand.mockReturnValue('статистика');
      mockInteraction.options.getString.mockReturnValue('general');

      // Выполнение
      await analyticsCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.options.getSubcommand).toHaveBeenCalled();
      expect(mockInteraction.options.getString).toHaveBeenCalledWith('категорія');
      expect(mockAnalyticsService.getStatistics).toHaveBeenCalledWith('general');
      expect(mockInteraction.reply).toHaveBeenCalled();
    });

    it('should handle trends subcommand', async () => {
      // Настройка моков
      const mockAnalyticsService = {
        getTrends: jest.fn().mockResolvedValue({
          period: '7d',
          trends: [
            { date: '2024-01-01', users: 100, commands: 500 },
            { date: '2024-01-02', users: 120, commands: 600 },
          ],
        }),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockAnalyticsService);
      mockInteraction.options.getSubcommand.mockReturnValue('тренди');
      mockInteraction.options.getString.mockReturnValue('7d');

      // Выполнение
      await analyticsCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.options.getSubcommand).toHaveBeenCalled();
      expect(mockInteraction.options.getString).toHaveBeenCalledWith('період');
      expect(mockAnalyticsService.getTrends).toHaveBeenCalledWith('7d');
      expect(mockInteraction.reply).toHaveBeenCalled();
    });

    it('should handle insights subcommand', async () => {
      // Настройка моков
      const mockAnalyticsService = {
        getInsights: jest.fn().mockResolvedValue({
          insights: [
            'Популярність команди /пошук зросла на 25%',
            'Середній час відповіді AI зменшився на 15%',
          ],
          recommendations: [
            'Додати більше фільтрів для пошуку',
            'Оптимізувати AI відповіді',
          ],
        }),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockAnalyticsService);
      mockInteraction.options.getSubcommand.mockReturnValue('інсайти');

      // Выполнение
      await analyticsCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.options.getSubcommand).toHaveBeenCalled();
      expect(mockAnalyticsService.getInsights).toHaveBeenCalled();
      expect(mockInteraction.reply).toHaveBeenCalled();
    });

    it('should handle invalid subcommand', async () => {
      mockInteraction.options.getSubcommand.mockReturnValue('неіснуюча');

      // Выполнение
      await analyticsCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.reply).toHaveBeenCalledWith(
        expect.objectContaining({
          content: expect.stringContaining('Невідома підкоманда'),
          ephemeral: true,
        })
      );
    });

    it('should handle service error', async () => {
      // Настройка моков с ошибкой
      const mockAnalyticsService = {
        generateReport: jest.fn().mockRejectedValue(new Error('Analytics service error')),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockAnalyticsService);
      mockInteraction.options.getSubcommand.mockReturnValue('звіт');
      mockInteraction.options.getString.mockReturnValue('daily');

      // Выполнение
      await analyticsCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.reply).toHaveBeenCalledWith(
        expect.objectContaining({
          content: expect.stringContaining('Помилка'),
          ephemeral: true,
        })
      );
    });

    it('should handle empty report data', async () => {
      // Настройка моков с пустыми данными
      const mockAnalyticsService = {
        generateReport: jest.fn().mockResolvedValue({
          type: 'daily',
          data: {},
          timestamp: new Date(),
        }),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockAnalyticsService);
      mockInteraction.options.getSubcommand.mockReturnValue('звіт');
      mockInteraction.options.getString.mockReturnValue('daily');

      // Выполнение
      await analyticsCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.reply).toHaveBeenCalledWith(
        expect.objectContaining({
          content: expect.stringContaining('Дані для звіту відсутні'),
          ephemeral: true,
        })
      );
    });

    it('should handle export functionality', async () => {
      // Настройка моков для экспорта
      const mockAnalyticsService = {
        generateReport: jest.fn().mockResolvedValue({
          type: 'daily',
          data: { users: 100, commands: 500 },
          timestamp: new Date(),
          exportUrl: 'https://example.com/report.xlsx',
        }),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockAnalyticsService);
      mockInteraction.options.getSubcommand.mockReturnValue('звіт');
      mockInteraction.options.getString.mockReturnValue('daily');
      mockInteraction.options.getString.mockReturnValueOnce('daily').mockReturnValueOnce('excel');

      // Выполнение
      await analyticsCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.reply).toHaveBeenCalledWith(
        expect.objectContaining({
          content: expect.stringContaining('https://example.com/report.xlsx'),
        })
      );
    });
  });
}); 