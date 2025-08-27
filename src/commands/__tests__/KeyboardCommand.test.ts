import { KeyboardCommand } from '../KeyboardCommand';
import type { BotConfig } from '@/types';

// Mock config for testing
const mockConfig = {
  discord: {
    prefix: '!',
    intents: ['GUILDS', 'GUILD_MESSAGES']
  }
} as unknown as BotConfig;

describe('KeyboardCommand', () => {
  let command: KeyboardCommand;

  beforeEach(() => {
    command = new KeyboardCommand(mockConfig);
  });

  test('should create KeyboardCommand instance', () => {
    expect(command).toBeInstanceOf(KeyboardCommand);
  });

  test('should have correct command name and description', () => {
    expect(command.name).toBe('keyboard');
    expect(command.description).toBe('Keyboard navigation settings');
  });

  test('should have correct category', () => {
    expect(command.category).toBe('settings');
  });
});