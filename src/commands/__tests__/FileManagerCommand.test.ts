import { FileManagerCommand } from '@/commands/FileManagerCommand';
import type { BotConfig } from '@/types';

const makeInteraction = (sub: string, opts: Record<string, string | null> = {}) => {
  const replies: any[] = [];
  const interaction: any = {
    user: { id: 'u1', tag: 'user#0001' },
    deferred: false,
    options: {
      getSubcommand: () => sub,
      getString: (name: string) => (name in opts ? opts[name] : null),
    },
    client: {
      serviceContainer: {
        get: (key: string) => undefined,
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
  env: 'test',
  discord: { token: '', clientId: '', guildId: '' } as any,
  google: { credentials: { client_email: 'x', private_key: 'y' } as any, driveFolderId: 'root-folder' } as any,
  ai: { provider: 'none', openai: {} as any, ollama: {} as any } as any,
  server: {} as any,
  cache: {} as any,
});

// Minimal GoogleService mock
const makeGoogleMock = () => {
  return {
    listDriveFilesInFolder: jest.fn(),
    getDriveFileMetadata: jest.fn(),
    exportDriveFile: jest.fn(),
    downloadDriveFile: jest.fn(),
  } as any;
};

describe('FileManagerCommand', () => {
  beforeEach(() => {
    jest.resetAllMocks();
  });

  test('search: fails when no driveFolderId and no folder option', async () => {
    const cfg = baseConfig();
    (cfg.google as any).driveFolderId = undefined;
    const cmd = new FileManagerCommand(cfg);
    const interaction = makeInteraction('пошук', { 'запит': 'test', 'папка': null });

    await cmd.execute({ interaction } as any);

    // Expect immediate error reply since validation passes but folderId missing after defer
    const edits = interaction.__replies.filter((r: any) => r.type === 'edit');
    expect(edits.length).toBe(1);
    expect(edits[0].payload.embeds?.[0]?.data?.description || edits[0].payload.content)
      .toMatch(/Не вказано ID папки|ID папки/);
  });

  test('search: empty results', async () => {
    const cfg = baseConfig();
    const cmd = new FileManagerCommand(cfg);
    const interaction = makeInteraction('пошук', { 'запит': 'nothing', 'папка': null });
    const google = makeGoogleMock();
    google.listDriveFilesInFolder.mockResolvedValueOnce([]);
    interaction.client.serviceContainer.get = (k: string) => (k === 'google' ? google : undefined);

    await cmd.execute({ interaction } as any);

    const edits = interaction.__replies.filter((r: any) => r.type === 'edit');
    expect(edits.length).toBe(1);
    const text = edits[0].payload.embeds?.[0]?.data?.description as string;
    expect(text).toMatch(/Нічого не знайдено/);
  });

  test('search: formats more than 20 results with summary', async () => {
    const cfg = baseConfig();
    const cmd = new FileManagerCommand(cfg);
    const interaction = makeInteraction('пошук', { 'запит': 'file', 'папка': null });
    const google = makeGoogleMock();

    const files = Array.from({ length: 25 }, (_, i) => ({ id: `id_${i}`, name: `File ${i}`, mimeType: 'application/octet-stream' }));
    google.listDriveFilesInFolder.mockResolvedValueOnce(files);
    interaction.client.serviceContainer.get = (k: string) => (k === 'google' ? google : undefined);

    await cmd.execute({ interaction } as any);

    const edits = interaction.__replies.filter((r: any) => r.type === 'edit');
    expect(edits.length).toBe(1);
    const description = edits[0].payload.embeds?.[0]?.data?.description as string;
    expect(description).toMatch(/…та ще 5/);
    expect(description.split('\n').filter(l => /^\d+\. /.test(l)).length).toBe(20);
  });

  test('read: downloads Google Sheet as xlsx attachment', async () => {
    const cfg = baseConfig();
    const cmd = new FileManagerCommand(cfg);
    const interaction = makeInteraction('читати', { 'id': 'sheet123' });
    const google = makeGoogleMock();
    google.getDriveFileMetadata.mockResolvedValueOnce({ id: 'sheet123', name: 'MySheet', mimeType: 'application/vnd.google-apps.spreadsheet' });
    google.exportDriveFile.mockResolvedValueOnce(Buffer.from('xlsx-data'));
    interaction.client.serviceContainer.get = (k: string) => (k === 'google' ? google : undefined);

    await cmd.execute({ interaction } as any);

    const edits = interaction.__replies.filter((r: any) => r.type === 'edit');
    expect(edits.length).toBe(1);
    expect(edits[0].payload.files?.[0]).toBeDefined();
    // file name should end with .xlsx
    const att: any = edits[0].payload.files[0];
    expect(att?.name || att?.attachment?.name || '').toMatch(/\.xlsx$/);
  });
});
