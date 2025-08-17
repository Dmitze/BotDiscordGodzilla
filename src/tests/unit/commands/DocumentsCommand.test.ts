/**
 * Unit тесты для DocumentsCommand
 */

import { describe, it, expect, beforeEach } from '@jest/globals';
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
      expect(documentsCommand.getDescription()).toBe('📄 Робота з військовими документами ЗСУ');
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
    it('should handle personnel search subcommand', async () => {
      // Моки ввода: підкоманда та опції
      mockInteraction.options.getSubcommand.mockReturnValue('особовий-склад');
      mockInteraction.options.getString.mockImplementation((name: string) => {
        if (name === 'дія') return 'search';
        if (name === 'запит') return 'тест';
        return null;
      });

      await documentsCommand.execute(mockInteraction);

      expect(mockInteraction.options.getSubcommand).toHaveBeenCalled();
      expect(mockInteraction.options.getString).toHaveBeenCalledWith('дія', true);
      expect(mockInteraction.options.getString).toHaveBeenCalledWith('запит');
      expect(mockInteraction.reply).toHaveBeenCalledWith(
        expect.objectContaining({ embeds: expect.any(Array) })
      );
    });

    it('should handle invalid subcommand', async () => {
      mockInteraction.options.getSubcommand.mockReturnValue('неіснуюча');

      await documentsCommand.execute(mockInteraction);

      expect(mockInteraction.reply).toHaveBeenCalledWith(
        expect.stringContaining('Невідома підкоманда')
      );
    });

    it('should handle error during execution', async () => {
      mockInteraction.options.getSubcommand.mockReturnValue('особовий-склад');
      // Заставим getString кинути помилку, щоб перейти в catch
      mockInteraction.options.getString.mockImplementation(() => {
        throw new Error('boom');
      });

      await documentsCommand.execute(mockInteraction);

      expect(mockInteraction.reply).toHaveBeenCalledWith(
        expect.stringContaining('Помилка обробки документів')
      );
    });
  });
}); 