/**
 * Unit тесты для DocumentsCommand
 */

import { jest, describe, it, expect, beforeEach } from '@jest/globals';
import { DocumentsCommand } from '../../../commands/DocumentsCommand';
import { createMockConfig, createMockInteraction } from '../../utils/testHelpers';

describe('DocumentsCommand', () => {
  let documentsCommand: DocumentsCommand;
  let mockConfig: any;
  let mockInteraction: any;

  beforeEach(() => {
    mockConfig = createMockConfig();
    documentsCommand = new DocumentsCommand(mockConfig);
    mockInteraction = createMockInteraction();
  });

  describe('constructor', () => {
    it('should create DocumentsCommand instance', () => {
      expect(documentsCommand).toBeInstanceOf(DocumentsCommand);
    });

    it('should have correct name', () => {
      expect(documentsCommand.getName()).toBe('документи');
    });

    it('should have correct description', () => {
      expect(documentsCommand.getDescription()).toBe('Управління документами та експорт');
    });
  });

  describe('getData', () => {
    it('should return SlashCommandBuilder', () => {
      const data = documentsCommand.getData();
      expect(data).toBeDefined();
      expect(data.name).toBe('документи');
    });
  });

  describe('execute', () => {
    it('should handle search subcommand', async () => {
      // Настройка моков
      const mockGoogleService = {
        searchDocuments: jest.fn().mockResolvedValue([
          { id: '1', name: 'Document 1', type: 'pdf' },
          { id: '2', name: 'Document 2', type: 'docx' },
        ]),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockGoogleService);
      mockInteraction.options.getSubcommand.mockReturnValue('пошук');
      mockInteraction.options.getString.mockReturnValue('тест');

      // Выполнение
      await documentsCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.options.getSubcommand).toHaveBeenCalled();
      expect(mockInteraction.options.getString).toHaveBeenCalledWith('запит');
      expect(mockGoogleService.searchDocuments).toHaveBeenCalledWith('тест');
      expect(mockInteraction.reply).toHaveBeenCalled();
    });

    it('should handle export subcommand', async () => {
      // Настройка моков
      const mockGoogleService = {
        exportData: jest.fn().mockResolvedValue('exported_data'),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockGoogleService);
      mockInteraction.options.getSubcommand.mockReturnValue('експорт');
      mockInteraction.options.getString.mockReturnValue('excel');

      // Выполнение
      await documentsCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.options.getSubcommand).toHaveBeenCalled();
      expect(mockInteraction.options.getString).toHaveBeenCalledWith('формат');
      expect(mockGoogleService.exportData).toHaveBeenCalledWith('excel');
      expect(mockInteraction.reply).toHaveBeenCalled();
    });

    it('should handle invalid subcommand', async () => {
      mockInteraction.options.getSubcommand.mockReturnValue('неіснуюча');

      // Выполнение
      await documentsCommand.execute(mockInteraction);

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
      const mockGoogleService = {
        searchDocuments: jest.fn().mockRejectedValue(new Error('Service error')),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockGoogleService);
      mockInteraction.options.getSubcommand.mockReturnValue('пошук');
      mockInteraction.options.getString.mockReturnValue('тест');

      // Выполнение
      await documentsCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.reply).toHaveBeenCalledWith(
        expect.objectContaining({
          content: expect.stringContaining('Помилка'),
          ephemeral: true,
        })
      );
    });
  });
}); 