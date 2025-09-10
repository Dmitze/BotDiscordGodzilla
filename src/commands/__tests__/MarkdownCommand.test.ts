// Mock all external dependencies
jest.mock('@/i18n', () => ({
  t: jest.fn().mockImplementation((key: string) => {
    return key.includes('description') ? 'Test description' : 
           key.includes('name') ? 'test' : 
           key.includes('service_unavailable') ? 'Service unavailable' :
           key.includes('rendering_failed') ? 'Rendering failed' : key;
  }),
}));

// Mock discord.js components to avoid import issues in tests
jest.mock('discord.js', () => {
  return {
    SlashCommandBuilder: jest.fn().mockImplementation(() => {
      return {
        setName: jest.fn().mockReturnThis(),
        setDescription: jest.fn().mockReturnThis(),
        setNameLocalizations: jest.fn().mockReturnThis(),
        setDescriptionLocalizations: jest.fn().mockReturnThis(),
        setDefaultMemberPermissions: jest.fn().mockReturnThis(),
        setDMPermission: jest.fn().mockReturnThis(),
        addStringOption: jest.fn().mockReturnThis(),
        toJSON: jest.fn().mockReturnValue({}),
      };
    }),
    SlashCommandStringOption: jest.fn().mockImplementation(() => {
      return {
        setName: jest.fn().mockReturnThis(),
        setDescription: jest.fn().mockReturnThis(),
        setRequired: jest.fn().mockReturnThis(),
        setMaxLength: jest.fn().mockReturnThis(),
        addChoices: jest.fn().mockReturnThis(),
      };
    }),
    AttachmentBuilder: jest.fn().mockImplementation(() => {
      return {
        setName: jest.fn().mockReturnThis(),
        setDescription: jest.fn().mockReturnThis(),
      };
    }),
  };
});

// Mock BaseCommand to avoid discord.js import issues
jest.mock('@/commands/BaseCommand', () => {
  return {
    BaseCommand: class {
      name: string;
      description: string;
      config: any;
      data: any;
      
      constructor(name: string, description: string, config: any) {
        this.name = name;
        this.description = description;
        this.config = config;
        this.data = {
          setName: jest.fn().mockReturnThis(),
          setDescription: jest.fn().mockReturnThis(),
          addStringOption: jest.fn().mockReturnThis(),
          toJSON: jest.fn().mockReturnValue({}),
        };
      }
      
      async execute(options: any) {
        // Mock execute method that calls onExecute
        return this.onExecute(options);
      }
      
      // This is what will be implemented by the actual command
      protected async onExecute(_options: any) {
        // This will be overridden by the actual implementation
      }
    }
  };
});

// Import after mocks

const makeMarkdownInteraction = (content: string = 'test content', format: string = 'text') => {
  const replies: any[] = [];
  return {
    user: { id: 'test-user' },
    channelId: 'test-channel',
    deferred: false,
    options: {
      getString: (name: string, _required?: boolean) => {
        if (name === 'content') return content;
        if (name === 'format') return format;
        return null;
      },
      getAttachment: (_name: string, _required?: boolean) => null,
    },
    client: {
      serviceContainer: {
        get: (_k: string) => null,
      },
    },
    deferReply: jest.fn(async () => { /* mock implementation */ }),
    reply: jest.fn(async (p: any) => { replies.push({ type: 'reply', payload: p }); }),
    editReply: jest.fn(async (p: any) => { replies.push({ type: 'edit', payload: p }); }),
    followUp: jest.fn(async (p: any) => { replies.push({ type: 'followUp', payload: p }); }),
    __replies: replies,
  };
};

describe('MarkdownCommand', () => {
  beforeEach(() => {
    jest.resetAllMocks();
  });

  test('should render markdown as text', async () => {
    // Create a minimal mock command
    const cmd = new (await import('@/commands/MarkdownCommand')).MarkdownCommand({} as any);
    
    const interaction = makeMarkdownInteraction('# Test Markdown', 'text');
    
    const markdownService = {
      renderToText: jest.fn().mockResolvedValue('# Rendered Test Markdown'),
    };
    
    interaction.client.serviceContainer.get = (_k: string) => (markdownService as any);

    await cmd.execute({ interaction } as any);

    expect(interaction.deferReply).toHaveBeenCalled();
    expect(markdownService.renderToText).toHaveBeenCalledWith('# Test Markdown');
    
    const edits = interaction.__replies.filter((r: any) => r.type === 'edit');
    expect(edits.length).toBe(1);
  });

  test('should render markdown as image', async () => {
    // Create a minimal mock command
    const cmd = new (await import('@/commands/MarkdownCommand')).MarkdownCommand({} as any);
    
    const interaction = makeMarkdownInteraction('# Test Markdown', 'image');
    
    const attachment = { name: 'markdown-render.png' };
    const markdownService = {
      renderToImage: jest.fn().mockResolvedValue(attachment),
    };
    
    interaction.client.serviceContainer.get = (_k: string) => (markdownService as any);

    await cmd.execute({ interaction } as any);

    expect(interaction.deferReply).toHaveBeenCalled();
    expect(markdownService.renderToImage).toHaveBeenCalledWith('# Test Markdown');
    
    const edits = interaction.__replies.filter((r: any) => r.type === 'edit');
    expect(edits.length).toBe(1);
  });

  test('should format markdown content', async () => {
    // Create a minimal mock command
    const CmdClass = (await import('@/commands/MarkdownCommand')).MarkdownCommand;
    const cmd = new CmdClass({} as any);
    
    const interaction = makeMarkdownInteraction('# Hello World\nThis is **bold** text!');
    
    const markdownService = {
      renderToText: jest.fn().mockResolvedValue('Formatted markdown content'),
    };
    
    interaction.client.serviceContainer.get = (_key: string) => (markdownService as any);

    await cmd.execute({ interaction } as any);

    expect(interaction.deferReply).toHaveBeenCalled();
    expect(markdownService.renderToText).toHaveBeenCalledWith('# Hello World\nThis is **bold** text!');
    
    const edits = interaction.__replies.filter((r: any) => r.type === 'edit');
    expect(edits.length).toBe(1);
  });

  test('should handle service unavailability', async () => {
    // Create a minimal mock command
    const cmd = new (await import('@/commands/MarkdownCommand')).MarkdownCommand({} as any);
    
    const interaction = makeMarkdownInteraction('# Test Markdown', 'text');
    
    interaction.client.serviceContainer.get = (_k: string) => null;

    await cmd.execute({ interaction } as any);

    expect(interaction.deferReply).toHaveBeenCalled();
    
    const edits = interaction.__replies.filter((r: any) => r.type === 'edit');
    expect(edits.length).toBe(1);
  });

  test('should handle rendering errors', async () => {
    // Create a minimal mock command
    const cmd = new (await import('@/commands/MarkdownCommand')).MarkdownCommand({} as any);
    
    const interaction = makeMarkdownInteraction('# Test Markdown', 'text');
    
    const markdownService = {
      renderToText: jest.fn().mockRejectedValue(new Error('Rendering failed')),
    };
    
    interaction.client.serviceContainer.get = (_k: string) => (markdownService as any);

    await cmd.execute({ interaction } as any);

    expect(interaction.deferReply).toHaveBeenCalled();
    
    const edits = interaction.__replies.filter((r: any) => r.type === 'edit');
    expect(edits.length).toBe(1);
  });
});