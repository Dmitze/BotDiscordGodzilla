/**
 * Unit тесты для AIAssistantCommand
 */

import { jest, describe, it, expect, beforeEach } from '@jest/globals';
import { AIAssistantCommand } from '../../../commands/AIAssistantCommand';
import { createMockConfig, createMockInteraction } from '../../utils/testHelpers';

describe('AIAssistantCommand', () => {
  let aiAssistantCommand: AIAssistantCommand;
  let mockConfig: any;
  let mockInteraction: any;

  beforeEach(() => {
    mockConfig = createMockConfig();
    aiAssistantCommand = new AIAssistantCommand(mockConfig);
    mockInteraction = createMockInteraction();
  });

  describe('constructor', () => {
    it('should create AIAssistantCommand instance', () => {
      expect(aiAssistantCommand).toBeInstanceOf(AIAssistantCommand);
    });

    it('should have correct name', () => {
      expect(aiAssistantCommand.getName()).toBe('ai_асистент');
    });

    it('should have correct description', () => {
      expect(aiAssistantCommand.getDescription()).toBe('AI-асистент для відповідей на запитання');
    });
  });

  describe('getData', () => {
    it('should return SlashCommandBuilder', () => {
      const data = aiAssistantCommand.getData();
      expect(data).toBeDefined();
      expect(data.name).toBe('ai_асистент');
    });
  });

  describe('execute', () => {
    it('should handle AI request', async () => {
      // Настройка моков
      const mockAIService = {
        generateResponse: jest.fn().mockResolvedValue('AI response'),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockAIService);
      mockInteraction.options.getString.mockReturnValue('Привіт, як справи?');

      // Выполнение
      await aiAssistantCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.options.getString).toHaveBeenCalledWith('запит');
      expect(mockAIService.generateResponse).toHaveBeenCalledWith('Привіт, як справи?');
      expect(mockInteraction.reply).toHaveBeenCalled();
    });

    it('should handle empty query', async () => {
      mockInteraction.options.getString.mockReturnValue('');

      // Выполнение
      await aiAssistantCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.reply).toHaveBeenCalledWith(
        expect.objectContaining({
          content: expect.stringContaining('Будь ласка, вкажіть запит'),
          ephemeral: true,
        })
      );
    });

    it('should handle AI service error', async () => {
      // Настройка моков с ошибкой
      const mockAIService = {
        generateResponse: jest.fn().mockRejectedValue(new Error('AI service error')),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockAIService);
      mockInteraction.options.getString.mockReturnValue('тест');

      // Выполнение
      await aiAssistantCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.reply).toHaveBeenCalledWith(
        expect.objectContaining({
          content: expect.stringContaining('Помилка'),
          ephemeral: true,
        })
      );
    });

    it('should handle long response', async () => {
      // Настройка моков с длинным ответом
      const longResponse = 'A'.repeat(2000);
      const mockAIService = {
        generateResponse: jest.fn().mockResolvedValue(longResponse),
      };

      mockInteraction.client.serviceContainer.get.mockReturnValue(mockAIService);
      mockInteraction.options.getString.mockReturnValue('тест');

      // Выполнение
      await aiAssistantCommand.execute(mockInteraction);

      // Проверки
      expect(mockInteraction.reply).toHaveBeenCalled();
    });
  });
}); 