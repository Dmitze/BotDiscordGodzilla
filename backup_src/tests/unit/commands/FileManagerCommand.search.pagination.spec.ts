/* eslint-disable @typescript-eslint/no-unsafe-assignment */
/* eslint-disable @typescript-eslint/no-unsafe-member-access */
/* eslint-disable @typescript-eslint/no-unsafe-call */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-module-boundary-types */
/* eslint-disable @typescript-eslint/consistent-type-imports */
// No runtime imports needed from discord.js for this test
import { FileManagerCommand } from '@/commands/FileManagerCommand';
import * as Base from '@/commands/BaseCommand';
import type { BotConfig } from '@/types';

jest.mock('@/i18n', () => ({
  t: (k: string, args?: any) => {
    if (k === 'files.command.description') return 'desc';
    if (k === 'files.sub.search.description') return 'search desc';
    if (k === 'files.opt.query.description') return 'query';
    if (k === 'files.opt.folder.description') return 'folder';
    if (k.startsWith('files.search.buttons.')) return k.split('.').pop();
    if (k === 'files.error.serviceUnavailable') return 'serviceUnavailable';
    if (k === 'files.result.more') return `more ${args?.rest ?? ''}`;
    if (k === 'files.result.searchList') return `list ${args?.count}`;
    if (k === 'files.search.filteredByPolicy') return `filtered ${args?.count}`;
    if (k === 'files.search.changesSummary') return `changes ${args?.added}/${args?.removed}/${args?.modified}`;
    if (k === 'files.search.largeMark') return 'large';
    return k;
  }
}));

beforeAll(() => {
  jest.spyOn((Base as any).BaseCommand.prototype as any, 'startCleanupInterval').mockImplementation(() => {});
});

// minimal DriveFile type to satisfy tests
interface DriveFile {
  id: string;
  name: string;
  mimeType?: string;
  size?: string | number;
  owners?: Array<{ emailAddress?: string; displayName?: string } | undefined>;
  modifiedTime?: string;
}

type GoogleMock = {
  listDriveFiles: jest.Mock<Promise<{ files: DriveFile[]; nextPageToken?: string; changes?: { addedIds: string[]; removedIds: string[]; modified: { id: string }[] } }>, any>;
};

type OptionsMock = {
  getString: (name: string) => string | null;
  getInteger: (name: string) => number | null;
};

type ChatInteractionMock = {
  user: { id: string };
  channelId: string;
  options: OptionsMock;
  editReply: jest.Mock<any, any>;
};

function createInteraction(opts: Partial<OptionsMock> = {}): ChatInteractionMock {
  const base: OptionsMock = {
    getString: (_n: string) => null,
    getInteger: (_n: string) => null,
  };
  return {
    user: { id: 'u1' },
    channelId: 'c1',
    options: { ...base, ...opts },
    editReply: jest.fn(),
  };
}

function createConfig(): BotConfig {
  return {
    env: 'test',
    discord: { token: 'x', clientId: 'x', guildId: 'x', enableSlash: false, enableChat: false, enableMessageContentIntent: false },
    google: { apiKey: 'x', driveFolderId: 'root' },
    drive: { folderId: 'root', enableTextIndex: false, ttlTextSec: 60, indexCron: '* * * * *', allowedMime: [], ownerAllowlist: [], fileMaxSizeMb: 50 },
  } as unknown as BotConfig;
}

function createFiles(): DriveFile[] {
  const now = new Date();
  const day = 24 * 3600 * 1000;
  return [
    { id: 'a', name: 'alpha', mimeType: 'text/plain', size: 1024 * 1024 * 1, owners: [{ emailAddress: 'alice@a.com' }], modifiedTime: new Date(now.getTime() - 5 * day).toISOString() },
    { id: 'b', name: 'beta', mimeType: 'application/pdf', size: 1024 * 1024 * 20, owners: [{ displayName: 'Bob' }], modifiedTime: new Date(now.getTime() - 2 * day).toISOString() },
    { id: 'c', name: 'gamma', mimeType: 'application/pdf', size: 1024 * 1024 * 100, owners: [{ emailAddress: 'carol@c.com' }], modifiedTime: new Date(now.getTime() - 1 * day).toISOString() },
  ];
}

function makeCmd(google: GoogleMock) {
  const cmd = new FileManagerCommand(createConfig());
  // monkey-patch private accessor
  (cmd as any).getGoogleService = () => google;
  return cmd as any;
}

describe('FileManagerCommand search filters and pagination', () => {
  test('applies MIME, owner, date, size filters and sets footer x/y', async () => {
    const files = createFiles();
    const google: GoogleMock = {
      listDriveFiles: jest.fn().mockResolvedValue({ files, changes: { addedIds: ['c'], removedIds: [], modified: [{ id: 'b' }] } }),
    } as any;
    const cmd = makeCmd(google);

    const interaction = createInteraction({
      getString: (n: string) => {
        if (n === 'запит') return 'pdf';
        if (n === 'mime') return 'application/pdf';
        if (n === 'власник') return 'bo'; // matches Bob
        if (n === 'від') return new Date(Date.now() - 3 * 24 * 3600 * 1000).toISOString().slice(0, 10);
        if (n === 'до') return new Date(Date.now()).toISOString().slice(0, 10);
        if (n === 'сортування') return 'modifiedTime';
        return null;
      },
      getInteger: (n: string) => (n === 'ліміт' ? 1 : n === 'розмір_мін' ? 5 : n === 'розмір_макс' ? 80 : null),
    });

    // simulate handleSearch -> buildSearchPage
    await (cmd as any).handleSearch(interaction, { query: 'pdf', folder: 'root' });

    expect(google.listDriveFiles).toHaveBeenCalled();
    // ensure reply was edited with embed + components
    expect(interaction.editReply).toHaveBeenCalled();
    const arg = interaction.editReply.mock.calls[0][0];
    expect(arg.embeds).toBeTruthy();

    const embed = arg.embeds[0] as any;
    // Components rows exist and first page footer includes 1/1 due to single item after filters
    expect(arg.components).toBeTruthy();
    const footerText = (embed as any).data?.footer?.text || (embed as any).footer?.text;
    expect(footerText).toContain('1/1');
  });

  test('paginates 3 items with pageSize=2 shows 1/2 and disables prev/first', async () => {
    const files = createFiles();
    const google: GoogleMock = {
      listDriveFiles: jest.fn().mockResolvedValue({ files, changes: { addedIds: [], removedIds: [], modified: [] } }),
    } as any;
    const cmd = makeCmd(google);

    const interaction = createInteraction({
      getString: (n: string) => (n === 'запит' ? 'all' : null),
      getInteger: (n: string) => (n === 'ліміт' ? 2 : null),
    });

    await (cmd as any).handleSearch(interaction, { query: 'all', folder: 'root' });

    const arg = interaction.editReply.mock.calls[0][0];
    const footerText = (arg.embeds[0] as any).data?.footer?.text || (arg.embeds[0] as any).footer?.text;
    expect(footerText).toContain('1/2');

    // row1 buttons state
    const row1 = arg.components[0];
    const first = row1.components[0];
    const prev = row1.components[1];
    expect(first.data.disabled).toBe(true);
    expect(prev.data.disabled).toBe(true);
  });
});
