// Mock all external dependencies
jest.mock('@/i18n', () => ({
  t: jest.fn().mockImplementation((key: string) => {
    return key.includes('description') ? 'Test description' : 
           key.includes('name') ? 'test' : 
           key.includes('service_unavailable') ? 'Service unavailable' :
           key.includes('generation_failed') ? 'Generation failed' :
           key.includes('history_reset') ? 'History reset' : key;
  }),
}));

jest.mock('@/commands/BaseCommand', () => {
  return {
    BaseCommand: class {
      name: string;
      description: string;
      config: any;
      
      constructor(name: string, description: string, config: any) {
        this.name = name;
        this.description = description;
        this.config = config;
      }
      
      async execute(options: any) {
        // Mock execute method
        return this.onExecute(options);
      }
    }
  };
});

// Import after mocks
import { OllamaCommand } from '@/commands/OllamaCommand';

const makeInteraction = (prompt: string, model: string | undefined = undefined, reset: boolean = false) => {
  const replies: any[] = [];
  const interaction: any = {
    user: { id: 'u1', tag: 'user#0001' },
    channelId: 'test-channel',
    deferred: false,
    options: {
      getString: (name: string, required?: boolean) => {
        if (name === 'prompt') return prompt;
        if (name === 'model') return model;
        return null;
      },
      getBoolean: (name: string) => {
        if (name === 'reset') return reset;
        return false;
      }
    },
    client: {
      serviceContainer: {
        get: (_key: string) => undefined,
      },
    },
    deferReply: jest.fn(async () => { interaction.deferred = true; }),
    reply: jest.fn(async (p: any) => { replies.push({ type: 'reply', payload: p }); }),
    editReply: jest.fn(async (p: any) => { replies.push({ type: 'edit', payload: p }); }),
    __replies: replies,
  };
  return interaction;
};

describe('OllamaCommand', () => {
  beforeEach(() => {
    jest.resetAllMocks();
  });

  test('should generate response from Ollama service', async () => {
    // Create a minimal mock command
    const cmd: any = {
      onExecute: (await import('@/commands/OllamaCommand')).OllamaCommand.prototype.onExecute
    };
    
    const interaction = makeInteraction('Hello, how are you?');
    
    const ollamaService = {
      generate: jest.fn().mockResolvedValue('I am doing well, thank you for asking!'),
    };
    
    interaction.client.serviceContainer.get = (k: string) => (k === 'ollama' ? ollamaService : undefined);

    await cmd.onExecute({ interaction });

    expect(interaction.deferReply).toHaveBeenCalled();
    expect(ollamaService.generate).toHaveBeenCalledWith('Hello, how are you?', {
      channelId: 'test-channel'
    });
    
    const edits = interaction.__replies.filter((r: any) => r.type === 'edit');
    expect(edits.length).toBe(1);
  });

  test('should handle service unavailability', async () => {
    // Create a minimal mock command
    const cmd: any = {
      onExecute: (await import('@/commands/OllamaCommand')).OllamaCommand.prototype.onExecute
    };
    
    const interaction = makeInteraction('Hello, how are you?');
    
    interaction.client.serviceContainer.get = (k: string) => null;

    await cmd.onExecute({ interaction });

    expect(interaction.deferReply).toHaveBeenCalled();
    
    const edits = interaction.__replies.filter((r: any) => r.type === 'edit');
    expect(edits.length).toBe(1);
  });

  test('should handle generation errors', async () => {
    // Create a minimal mock command
    const cmd: any = {
      onExecute: (await import('@/commands/OllamaCommand')).OllamaCommand.prototype.onExecute
    };
    
    const interaction = makeInteraction('Hello, how are you?');
    
    const ollamaService = {
      generate: jest.fn().mockRejectedValue(new Error('Generation failed')),
    };
    
    interaction.client.serviceContainer.get = (k: string) => (k === 'ollama' ? ollamaService : undefined);

    await cmd.onExecute({ interaction });

    expect(interaction.deferReply).toHaveBeenCalled();
    
    const edits = interaction.__replies.filter((r: any) => r.type === 'edit');
    expect(edits.length).toBe(1);
  });

  test('should reset channel history', async () => {
    // Create a minimal mock command
    const cmd: any = {
      onExecute: (await import('@/commands/OllamaCommand')).OllamaCommand.prototype.onExecute
    };
    
    const interaction = makeInteraction('Hello, how are you?', undefined, true);
    
    const ollamaService = {
      resetChannelHistory: jest.fn().mockResolvedValue(undefined),
    };
    
    interaction.client.serviceContainer.get = (k: string) => (k === 'ollama' ? ollamaService : undefined);

    await cmd.onExecute({ interaction });

    expect(interaction.deferReply).toHaveBeenCalled();
    expect(ollamaService.resetChannelHistory).toHaveBeenCalledWith('test-channel');
    
    const edits = interaction.__replies.filter((r: any) => r.type === 'edit');
    expect(edits.length).toBe(1);
  });
});