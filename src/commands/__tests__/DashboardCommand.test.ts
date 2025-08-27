import { DashboardCommand } from '../DashboardCommand';
import type { BotConfig } from '@/types';

// Mock config for testing
const mockConfig = {
  discord: {
    prefix: '!',
    intents: ['GUILDS', 'GUILD_MESSAGES']
  }
} as unknown as BotConfig;

describe('DashboardCommand', () => {
  let command: DashboardCommand;

  beforeEach(() => {
    command = new DashboardCommand(mockConfig);
  });

  test('should create DashboardCommand instance', () => {
    expect(command).toBeInstanceOf(DashboardCommand);
  });

  test('should have correct command name and description', () => {
    expect(command.name).toBe('dashboard');
    expect(command.description).toBe('Dashboard views and file display configuration');
  });

  test('should have correct category', () => {
    expect(command.category).toBe('files');
  });
});