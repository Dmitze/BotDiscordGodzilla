/**
 * Unit тесты для FileManagerCommand
 */

import { jest, describe, it, expect, beforeEach } from '@jest/globals';
import { FileManagerCommand } from '../../../commands/FileManagerCommand';
import { createMockConfig, createMockInteraction } from '../../utils/testHelpers';

describe('FileManagerCommand', () => {
  let fileManagerCommand: FileManagerCommand;
  let mockConfig: any;
  let mockInteraction: any;

  beforeEach(() => {
    mockConfig = createMockConfig();
    fileManagerCommand = new FileManagerCommand(mockConfig);
    mockInteraction = createMockInteraction();
  });

  describe('constructor', () => {
    it('should create FileManagerCommand instance', () => {
      expect(fileManagerCommand).toBeInstanceOf(FileManagerCommand);
    });

    it('should have correct name', () => {
      expect(fileManagerCommand.getName()).toBe('файли');
    });

    it('should have correct description', () => {
      expect(fileManagerCommand.getDescription()).toBe('Управління файлами та документами');
    });
  });

  describe('getData', () => {
    it('should return SlashCommandBuilder', () => {
      const data = fileManagerCommand.getData();
      expect(data).toBeDefined();
      expect(data.name).toBe('файли');
    });
  });

  describe('execute', () => {
    it('should handle search subcommand', async () => {
      // Настройка моков
      const mockGoogleService = {
        searchFiles: jest.fn().mockResolvedValue([
          { id: '1', name: 'File 1.pdf', mimeType: 'application/pdf' },
          { id: '2', name: 'File 2.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
        ]),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockGoogleService);
      mockInteraction.options.getSubcommand.mockReturnValue('пошук');
      mockInteraction.options.getString.mockReturnValue('документ');

      // Выполнение
      await fileManagerCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.options.getSubcommand).toHaveBeenCalled();
      expect(mockInteraction.options.getString).toHaveBeenCalledWith('запит');
      expect(mockGoogleService.searchFiles).toHaveBeenCalledWith('документ');
      expect(mockInteraction.reply).toHaveBeenCalled();
    });

    it('should handle analyze subcommand', async () => {
      // Настройка моков
      const mockGoogleService = {
        analyzeFile: jest.fn().mockResolvedValue({
          summary: 'File analysis summary',
          metadata: { size: '1MB', type: 'pdf' },
        }),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockGoogleService);
      mockInteraction.options.getSubcommand.mockReturnValue('аналіз');
      mockInteraction.options.getString.mockReturnValue('file_id_123');
      mockInteraction.options.getString.mockReturnValueOnce('file_id_123').mockReturnValueOnce('summary');

      // Выполнение
      await fileManagerCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.options.getSubcommand).toHaveBeenCalled();
      expect(mockInteraction.options.getString).toHaveBeenCalledWith('id');
      expect(mockInteraction.options.getString).toHaveBeenCalledWith('тип');
      expect(mockGoogleService.analyzeFile).toHaveBeenCalledWith('file_id_123', 'summary');
      expect(mockInteraction.reply).toHaveBeenCalled();
    });

    it('should handle download subcommand', async () => {
      // Настройка моков
      const mockGoogleService = {
        downloadFile: jest.fn().mockResolvedValue('file_content'),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockGoogleService);
      mockInteraction.options.getSubcommand.mockReturnValue('завантаження');
      mockInteraction.options.getString.mockReturnValue('file_id_123');

      // Выполнение
      await fileManagerCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.options.getSubcommand).toHaveBeenCalled();
      expect(mockInteraction.options.getString).toHaveBeenCalledWith('id');
      expect(mockGoogleService.downloadFile).toHaveBeenCalledWith('file_id_123');
      expect(mockInteraction.reply).toHaveBeenCalled();
    });

    it('should handle invalid subcommand', async () => {
      mockInteraction.options.getSubcommand.mockReturnValue('неіснуюча');

      // Выполнение
      await fileManagerCommand.execute(mockInteraction);

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
        searchFiles: jest.fn().mockRejectedValue(new Error('Service error')),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockGoogleService);
      mockInteraction.options.getSubcommand.mockReturnValue('пошук');
      mockInteraction.options.getString.mockReturnValue('тест');

      // Выполнение
      await fileManagerCommand.execute(mockInteraction);

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
        searchFiles: jest.fn().mockResolvedValue([]),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockGoogleService);
      mockInteraction.options.getSubcommand.mockReturnValue('пошук');
      mockInteraction.options.getString.mockReturnValue('неіснуючий');

      // Выполнение
      await fileManagerCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.reply).toHaveBeenCalledWith(
        expect.objectContaining({
          content: expect.stringContaining('Файлів не знайдено'),
          ephemeral: true,
        })
      );
    });
  });
}); 