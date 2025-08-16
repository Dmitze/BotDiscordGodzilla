/**
 * Unit тесты для AIAssistantCommand
 */

import { describe, it, expect, beforeEach } from '@jest/globals';
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
      expect(aiAssistantCommand.getName()).toBe('ai');
    });

    it('should have correct description', () => {
      expect(aiAssistantCommand.getDescription()).toBe('🤖 AI-асистент для роботи з Google Sheets');
    });
  });

  describe('getData', () => {
    it('should return SlashCommandBuilder', () => {
      const data = aiAssistantCommand.getData();
      expect(data).toBeDefined();
      expect(data.name).toBe('ai');
    });
  });

  describe('execute', () => {
    it('should handle AI request', async () => {
      // Настройка моков: успешный ответ AI
      (aiAssistantCommand as any).processAIQuery = async () => ({
        response: 'AI response',
        confidence: 0.9,
        action: 'search',
      });
      mockInteraction.options.getString.mockReturnValue('Привіт, як справи?');

      // Выполнение
      await aiAssistantCommand.execute({ interaction: mockInteraction } as any);

      // Проверки
      expect(mockInteraction.options.getString).toHaveBeenCalledWith('запит');
      expect(mockInteraction.deferReply).toHaveBeenCalled();
      expect(mockInteraction.editReply).toHaveBeenCalled();
    });

    it('should handle empty query', async () => {
      mockInteraction.options.getString.mockReturnValue('');

      // Выполнение
      await aiAssistantCommand.execute({ interaction: mockInteraction } as any);

      // Проверки
      expect(mockInteraction.deferReply).toHaveBeenCalled();
      const calls = (mockInteraction.editReply.mock.calls?.length || 0) +
        (mockInteraction.reply.mock.calls?.length || 0);
      expect(calls).toBeGreaterThan(0);
    });

    it('should handle AI service error', async () => {
      // Настройка: processAIQuery кидает ошибку
      (aiAssistantCommand as any).processAIQuery = async () => {
        throw new Error('AI service error');
      };
      mockInteraction.options.getString.mockReturnValue('тест');

      // Выполнение
      await aiAssistantCommand.execute({ interaction: mockInteraction } as any);

      // Проверки
      expect(mockInteraction.editReply).toHaveBeenCalled();
    });

    it('should handle long response', async () => {
      // Настройка моков с длинным ответом
      const longResponse = 'A'.repeat(2000);
      (aiAssistantCommand as any).processAIQuery = async () => ({
        response: longResponse,
        confidence: 0.95,
      });
      mockInteraction.options.getString.mockReturnValue('тест');

      // Выполнение
      await aiAssistantCommand.execute({ interaction: mockInteraction } as any);

      // Проверки
      expect(mockInteraction.editReply).toHaveBeenCalled();
    });
  });
}); 