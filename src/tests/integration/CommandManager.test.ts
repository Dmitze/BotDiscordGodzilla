/**
 * Интеграционные тесты для CommandManager
 */

import { jest, describe, it, expect, beforeEach } from '@jest/globals';
import { CommandManager } from '../../core/CommandManager';
import { createMockConfig } from '../utils/testHelpers';

describe('CommandManager Integration', () => {
  let commandManager: CommandManager;
  let mockConfig: any;
  let mockClient: any;

  beforeEach(() => {
    mockConfig = createMockConfig();
    mockClient = {
      application: {
        commands: {
          set: jest.fn(),
          create: jest.fn(),
        },
      },
    };
    commandManager = new CommandManager(mockClient, mockConfig);
  });

  describe('initialization', () => {
    it('should initialize command manager', async () => {
      await expect(commandManager.initialize()).resolves.not.toThrow();
    });

    it('should load all commands', async () => {
      await commandManager.initialize();
      
      const commands = commandManager.getCommands();
      expect(commands.size).toBeGreaterThan(0);
    });

    it('should register commands with Discord', async () => {
      await commandManager.initialize();
      
      expect(mockClient.application.commands.set).toHaveBeenCalled();
    });
  });

  describe('command execution', () => {
    it('should execute valid command', async () => {
      await commandManager.initialize();
      
      const mockInteraction = {
        commandName: 'пошук',
        options: {
          getString: jest.fn().mockReturnValue('тест'),
        },
        reply: jest.fn(),
        client: {
          serviceContainer: {
            get: jest.fn().mockReturnValue({
              searchData: jest.fn().mockResolvedValue([['test', 'data']]),
            }),
          },
        },
      };

      await expect(commandManager.execute(mockInteraction)).resolves.not.toThrow();
    });

    it('should handle invalid command', async () => {
      await commandManager.initialize();
      
      const mockInteraction = {
        commandName: 'неіснуюча_команда',
        reply: jest.fn(),
      };

      await commandManager.execute(mockInteraction);
      
      expect(mockInteraction.reply).toHaveBeenCalledWith(
        expect.objectContaining({
          content: expect.stringContaining('Команда не знайдена'),
          ephemeral: true,
        })
      );
    });
  });

  describe('command validation', () => {
    it('should validate command structure', async () => {
      await commandManager.initialize();
      
      const commands = commandManager.getCommands();
      
      for (const [name, command] of commands) {
        expect(command.getName()).toBeDefined();
        expect(command.getDescription()).toBeDefined();
        expect(command.getData()).toBeDefined();
      }
    });
  });
}); 