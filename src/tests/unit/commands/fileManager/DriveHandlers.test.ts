import { handleDriveAction } from '@/commands/modules/fileManager/handlers';

describe('fileManager handlers: handleDriveAction', () => {
  function makeInteraction() {
    const state = { deferred: false, replied: false } as any;
    return Object.assign(state, {
      deferReply: jest.fn(async () => { state.deferred = true; }),
      editReply: jest.fn(async (_: any) => { state.replied = true; }),
      reply: jest.fn(async (_: any) => { state.replied = true; }),
      followUp: jest.fn(async (_: any) => { state.replied = true; }),
      client: { serviceContainer: new Map() },
    });
  }

  function makeDeps(overrides?: Partial<import('@/commands/modules/fileManager/handlers').DriveDeps>): import('@/commands/modules/fileManager/handlers').DriveDeps {
    const container = new Map<string, any>();
    const base: import('@/commands/modules/fileManager/handlers').DriveDeps = {
      config: {},
      getGoogleService: (_i: any) => undefined,
      isMimeAllowed: () => true,
      isOwnerAllowed: () => true,
      isTooLarge: () => false,
      getAnalysisTypeName: () => 'summary',
      resolve: <T = unknown>(interaction: any, name: string) => {
        if (interaction?.client?.serviceContainer instanceof Map) {
          return interaction.client.serviceContainer.get(name) as T;
        }
        return undefined as unknown as T;
      },
    };
    return Object.assign(base, overrides);
  }

  test('question: uses RAG first with filters.fileId', async () => {
    const interaction = makeInteraction();
    const rag = { answer: jest.fn(async () => ({ text: 'Відповідь RAG' })) };
    (interaction.client.serviceContainer as Map<string, any>).set('rag', rag);

    const deps = makeDeps();
    await handleDriveAction(interaction as any, 'question', 'FILE123', deps);

    // deferred and editReply called with RAG content
    expect(interaction.deferReply).toHaveBeenCalled();
    expect(rag.answer).toHaveBeenCalled();
    expect(rag.answer.mock.calls[0][1]).toEqual({ filters: { fileId: ['FILE123'] } });
    expect(interaction.editReply).toHaveBeenCalledWith({ content: expect.stringContaining('Відповідь RAG') });
  });

  test('question: falls back to SearchIndex snippets then AI', async () => {
    const interaction = makeInteraction();
    const hits = [ { fileId: 'FILEX', name: 'Doc X', snippet: 'snippet-x' } ];
    const searchIndex = { search: jest.fn(async () => ({ hits, total: 1 })) };
    const ai = { generateResponse: jest.fn(async () => ({ content: 'AI RESP' })) };
    (interaction.client.serviceContainer as Map<string, any>).set('searchIndex', searchIndex);
    (interaction.client.serviceContainer as Map<string, any>).set('ai', ai);

    const deps = makeDeps();
    await handleDriveAction(interaction as any, 'question', 'FILEX', deps);

    expect(interaction.deferReply).toHaveBeenCalled();
    expect(searchIndex.search).toHaveBeenCalledWith({ text: '*', limit: 6, filters: { fileId: ['FILEX'] } });
    expect(ai.generateResponse).toHaveBeenCalled();
    const promptArg = ai.generateResponse.mock.calls[0][0] as string;
    expect(promptArg).toContain('snippet-x');
    expect(interaction.editReply).toHaveBeenCalledWith({ content: expect.stringContaining('AI RESP') });
  });

  test('question: falls back to Google export when index empty, then AI', async () => {
    const interaction = makeInteraction();
    const searchIndex = { search: jest.fn(async () => ({ hits: [], total: 0 })) };
    const ai = { generateResponse: jest.fn(async () => ({ content: 'AI RESP 2' })) };
    const googleSvc = {
      getDriveFileMetadata: jest.fn(async () => ({ mimeType: 'application/vnd.google-apps.document' })),
      exportDriveFile: jest.fn(async () => Buffer.from('exported text content')),
    };
    (interaction.client.serviceContainer as Map<string, any>).set('searchIndex', searchIndex);
    (interaction.client.serviceContainer as Map<string, any>).set('ai', ai);

    const deps = makeDeps({ getGoogleService: () => googleSvc });
    await handleDriveAction(interaction as any, 'question', 'FILEZ', deps);

    expect(searchIndex.search).toHaveBeenCalled();
    expect(googleSvc.getDriveFileMetadata).toHaveBeenCalledWith('FILEZ');
    expect(googleSvc.exportDriveFile).toHaveBeenCalledWith('FILEZ', 'text/plain');
    const promptArg = ai.generateResponse.mock.calls[0][0] as string;
    expect(promptArg).toContain('exported text content');
    expect(interaction.editReply).toHaveBeenCalledWith({ content: expect.stringContaining('AI RESP 2') });
  });
});
