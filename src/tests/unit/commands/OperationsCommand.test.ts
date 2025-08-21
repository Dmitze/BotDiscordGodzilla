/**
 * Unit тесты для OperationsCommand
 */

import { jest, describe, it, expect, beforeEach } from '@jest/globals';
import { OperationsCommand } from '../../../commands/OperationsCommand';
import { createMockConfig, createMockInteraction } from '../../utils/testHelpers';

describe('OperationsCommand', () => {
  let operationsCommand: OperationsCommand;
  let mockConfig: any;
  let mockInteraction: any;

  beforeEach(() => {
    mockConfig = createMockConfig();
    operationsCommand = new OperationsCommand(mockConfig);
    mockInteraction = createMockInteraction();
  });

  describe('constructor', () => {
    it('should create OperationsCommand instance', () => {
      expect(operationsCommand).toBeInstanceOf(OperationsCommand);
    });

    it('should have correct name', () => {
      expect(operationsCommand.getName()).toBe('operations');
    });

    it('should have correct description', () => {
      expect(operationsCommand.getDescription()).toBe('⚔️ Оперативне управління ЗСУ');
    });
  });

  describe('getData', () => {
    it('should return SlashCommandBuilder', () => {
      const data = operationsCommand.getData();
      expect(data).toBeDefined();
      expect(data.name).toBe('operations');
    });
  });

  describe('execute', () => {
    it('should handle situation subcommand', async () => {
      // Настройка моков
      const mockOperationsService = {
        getSituation: (jest.fn() as any).mockResolvedValue({
          status: 'active',
          incidents: 2,
          resources: 'available',
        }),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockOperationsService);
      mockInteraction.options.getSubcommand.mockReturnValue('situation');
      mockInteraction.options.getString.mockReturnValue('all');

      // Выполнение
      await operationsCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.options.getSubcommand).toHaveBeenCalled();
      expect(mockInteraction.options.getString).toHaveBeenCalledWith('sector');
      expect(mockOperationsService.getSituation).toHaveBeenCalledWith('all');
      expect(mockInteraction.reply).toHaveBeenCalled();
    });

    it('should handle tasks subcommand', async () => {
      // Настройка моков
      const mockOperationsService = {
        getTasks: (jest.fn() as any).mockResolvedValue([
          { id: '1', title: 'Task 1', status: 'in_progress' },
          { id: '2', title: 'Task 2', status: 'completed' },
        ]),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockOperationsService);
      mockInteraction.options.getSubcommand.mockReturnValue('tasks');
      // first required option 'action', then optional 'query'
      mockInteraction.options.getString
        .mockReturnValueOnce('current')
        .mockReturnValueOnce(undefined);

      // Выполнение
      await operationsCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.options.getSubcommand).toHaveBeenCalled();
      expect(mockInteraction.options.getString).toHaveBeenNthCalledWith(1, 'action', true);
      expect(mockInteraction.options.getString).toHaveBeenNthCalledWith(2, 'query');
      expect(mockOperationsService.getTasks).toHaveBeenCalledWith('current');
      expect(mockInteraction.reply).toHaveBeenCalled();
    });

    it('should handle coordination subcommand', async () => {
      // Настройка моков
      const mockOperationsService = {
        coordinate: (jest.fn() as any).mockResolvedValue({
          success: true,
          message: 'Coordination completed',
        }),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockOperationsService);
      mockInteraction.options.getSubcommand.mockReturnValue('coordination');
      // handler reads 'тип' (ua key) as required; service called with 'emergency'
      mockInteraction.options.getString
        .mockReturnValueOnce('emergency') // for 'тип'
        .mockReturnValueOnce(undefined); // for 'підрозділ'

      // Выполнение
      await operationsCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.options.getSubcommand).toHaveBeenCalled();
      // Options must be ASCII a-z per Discord API, implementation uses 'type' and 'unit'
      expect(mockInteraction.options.getString).toHaveBeenNthCalledWith(1, 'type', true);
      expect(mockInteraction.options.getString).toHaveBeenNthCalledWith(2, 'unit');
      expect(mockOperationsService.coordinate).toHaveBeenCalledWith('emergency');
      expect(mockInteraction.reply).toHaveBeenCalled();
    });

    it('should handle intelligence subcommand', async () => {
      // Настройка моков
      const mockOperationsService = {
        getIntelligence: (jest.fn() as any).mockResolvedValue({
          reports: 5,
          alerts: 2,
          analysis: 'Intelligence summary',
        }),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockOperationsService);
      mockInteraction.options.getSubcommand.mockReturnValue('intelligence');
      mockInteraction.options.getString
        .mockReturnValueOnce('daily') // for 'тип'
        .mockReturnValueOnce(undefined); // for 'район'

      // Выполнение
      await operationsCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.options.getSubcommand).toHaveBeenCalled();
      // Implementation uses ASCII option names 'type' and 'area'
      expect(mockInteraction.options.getString).toHaveBeenNthCalledWith(1, 'type', true);
      expect(mockInteraction.options.getString).toHaveBeenNthCalledWith(2, 'area');
      expect(mockOperationsService.getIntelligence).toHaveBeenCalledWith('daily');
      expect(mockInteraction.reply).toHaveBeenCalled();
    });

    it('should handle invalid subcommand', async () => {
      mockInteraction.options.getSubcommand.mockReturnValue('неіснуюча');

      // Выполнение
      await operationsCommand.execute(mockInteraction);

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
      const mockOperationsService = {
        getSituation: (jest.fn() as any).mockRejectedValue(new Error('Service error')),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockOperationsService);
      mockInteraction.options.getSubcommand.mockReturnValue('situation');
      mockInteraction.options.getString.mockReturnValue('all');

      // Выполнение
      await operationsCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.reply).toHaveBeenCalledWith(
        expect.objectContaining({
          content: expect.stringContaining('Помилка'),
          ephemeral: true,
        })
      );
    });

    it('should handle empty results', async () => {
      // Настройка моков с пустыми результатами
      const mockOperationsService = {
        getTasks: (jest.fn() as any).mockResolvedValue([]),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockOperationsService);
      mockInteraction.options.getSubcommand.mockReturnValue('tasks');
      mockInteraction.options.getString
        .mockReturnValueOnce('current') // action
        .mockReturnValueOnce(undefined); // query

      // Выполнение
      await operationsCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.reply).toHaveBeenCalledWith(
        expect.objectContaining({
          content: expect.stringContaining('Завдань не знайдено'),
          ephemeral: true,
        })
      );
    });
  });
}); 

