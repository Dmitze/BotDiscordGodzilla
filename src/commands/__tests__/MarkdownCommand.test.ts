import { MarkdownCommand } from '@/commands/MarkdownCommand';
import type { BotConfig } from '@/types';

const makeInteraction = (content: string, format: string = 'text') => {
  const replies: any[] = [];
  const interaction: any = {
    user: { id: 'u1', tag: 'user#0001' },
    deferred: false,
    options: {
      getString: (name: string, required?: boolean) => {
        if (name === 'content') return content;
        if (name === 'format') return format;
        return null;
      },
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

const baseConfig = (): BotConfig => ({
  discord: { token: '', clientId: '', guildId: '' } as any,
  google: { credentials: { client_email: 'x', private_key: 'y' } as any, driveFolderId: 'root-folder' } as any,
  ai: { provider: 'none', openai: {} as any, ollama: {} as any } as any,
  server: {} as any,
  cache: {} as any,
} as unknown as BotConfig);

describe('MarkdownCommand', () => {
  beforeEach(() => {
    jest.resetAllMocks();
  });

  test('should render markdown as text', async () => {
    const cfg = baseConfig();
    const cmd = new MarkdownCommand(cfg);
    const interaction = makeInteraction('# Test Markdown', 'text');
    
    const markdownService = {
      renderToText: jest.fn().mockResolvedValue('# Rendered Test Markdown'),
    };
    
    interaction.client.serviceContainer.get = (k: string) => (k === 'markdownRendering' ? markdownService : undefined);

    await cmd.execute({ interaction } as any);

    expect(interaction.deferReply).toHaveBeenCalled();
    expect(markdownService.renderToText).toHaveBeenCalledWith('# Test Markdown');
    
    const edits = interaction.__replies.filter((r: any) => r.type === 'edit');
    expect(edits.length).toBe(1);
    expect(edits[0].payload.content).toBe('# Rendered Test Markdown');
  });

  test('should render markdown as image', async () => {
    const cfg = baseConfig();
    const cmd = new MarkdownCommand(cfg);
    const interaction = makeInteraction('# Test Markdown', 'image');
    
    const attachment = { name: 'markdown-render.png' };
    const markdownService = {
      renderToImage: jest.fn().mockResolvedValue(attachment),
    };
    
    interaction.client.serviceContainer.get = (k: string) => (k === 'markdownRendering' ? markdownService : undefined);

    await cmd.execute({ interaction } as any);

    expect(interaction.deferReply).toHaveBeenCalled();
    expect(markdownService.renderToImage).toHaveBeenCalledWith('# Test Markdown', undefined);
    
    const edits = interaction.__replies.filter((r: any) => r.type === 'edit');
    expect(edits.length).toBe(1);
    expect(edits[0].payload.files).toEqual([attachment]);
  });

  test('should handle service unavailability', async () => {
    const cfg = baseConfig();
    const cmd = new MarkdownCommand(cfg);
    const interaction = makeInteraction('# Test Markdown', 'text');
    
    interaction.client.serviceContainer.get = (k: string) => null;

    await cmd.execute({ interaction } as any);

    expect(interaction.deferReply).toHaveBeenCalled();
    
    const edits = interaction.__replies.filter((r: any) => r.type === 'edit');
    expect(edits.length).toBe(1);
    expect(edits[0].payload.content).toBe('Markdown rendering service is unavailable');
  });

  test('should handle rendering errors', async () => {
    const cfg = baseConfig();
    const cmd = new MarkdownCommand(cfg);
    const interaction = makeInteraction('# Test Markdown', 'text');
    
    const markdownService = {
      renderToText: jest.fn().mockRejectedValue(new Error('Rendering failed')),
    };
    
    interaction.client.serviceContainer.get = (k: string) => (k === 'markdownRendering' ? markdownService : undefined);

    await cmd.execute({ interaction } as any);

    expect(interaction.deferReply).toHaveBeenCalled();
    
    const edits = interaction.__replies.filter((r: any) => r.type === 'edit');
    expect(edits.length).toBe(1);
    expect(edits[0].payload.content).toBe('Failed to render markdown content');
  });
});