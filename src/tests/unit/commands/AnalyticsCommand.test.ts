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
      expect(analyticsCommand.getName()).toBe('analytics');
    });

    it('should have correct description', () => {
      expect(analyticsCommand.getDescription()).toBe('Аналітика та звіти про використання бота');
    });
  });

  describe('getData', () => {
    it('should return SlashCommandBuilder', () => {
      const data = analyticsCommand.getData();
      expect(data).toBeDefined();
      expect(data.name).toBe('analytics');
    });
  });

  describe('execute', () => {
    it('should handle report option', async () => {
      // Setup mocks
      mockInteraction.options.getString.mockImplementation((name: string, required?: boolean) => {
        if (name === 'report') return 'usage';
        if (name === 'format') return 'text';
        return null;
      });
      mockInteraction.options.getInteger.mockReturnValue(10);

      // Execute
      await analyticsCommand.execute({ interaction: mockInteraction } as any);

      // Assertions
      expect(mockInteraction.options.getString).toHaveBeenCalledWith('report', true);
      expect(mockInteraction.options.getString).toHaveBeenCalledWith('format');
      expect(mockInteraction.editReply).toHaveBeenCalled();
    });

    it('should handle service error', async () => {
      // Setup mocks with error
      mockInteraction.options.getString.mockImplementation((name: string, required?: boolean) => {
        if (name === 'report') return 'invalid';
        if (name === 'format') return 'text';
        return null;
      });
      mockInteraction.options.getInteger.mockReturnValue(10);

      // Execute
      await analyticsCommand.execute({ interaction: mockInteraction } as any);

      // Assertions
      expect(mockInteraction.editReply).toHaveBeenCalledWith(
        expect.objectContaining({
          content: expect.stringContaining('Невірний тип звіту'),
        })
      );
    });
  });
}); 