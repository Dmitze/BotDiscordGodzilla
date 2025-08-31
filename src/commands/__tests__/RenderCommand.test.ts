import { RenderCommand } from '../RenderCommand';
import type { BotConfig } from '@/types';

// Mock config
const mockConfig: BotConfig = {
  discord: {
    token: 'mock-token',
    clientId: 'mock-client-id',
    guildId: 'mock-guild-id',
    enableChat: false,
    intents: []
  },
  google: {
    credentials: {
      client_email: 'test@example.com',
      private_key: 'mock-private-key'
    },
    folderId: 'mock-folder-id'
  },
  openai: {
    apiKey: 'mock-api-key'
  }
} as any;

describe('RenderCommand', () => {
  let renderCommand: RenderCommand;

  beforeEach(() => {
    renderCommand = new RenderCommand(mockConfig);
  });

  test('should create command with correct name', () => {
    expect(renderCommand.data.name).toBe('render');
  });

  test('should have correct description', () => {
    expect(renderCommand.data.description).toBe('Рендерить markdown в зображення');
  });
});