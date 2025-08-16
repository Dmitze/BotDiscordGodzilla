/* eslint-disable @typescript-eslint/no-unsafe-assignment */
/* eslint-disable @typescript-eslint/no-unsafe-member-access */
/* eslint-disable @typescript-eslint/no-unsafe-call */
/* eslint-disable @typescript-eslint/no-explicit-any */

import { FileManagerCommand } from '@/commands/FileManagerCommand';
import * as Base from '@/commands/BaseCommand';
import type { BotConfig } from '@/types';

jest.mock('@/i18n', () => ({
  t: (k: string, args?: any) => {
    if (k.startsWith('files.search.buttons.')) return k.split('.').pop();
    if (k === 'files.result.more') return `more ${args?.rest ?? ''}`;
    if (k === 'files.result.searchList') return `list ${args?.count}`;
    if (k === 'files.search.filteredByPolicy') return `filtered ${args?.count}`;
    if (k === 'files.search.changesSummary') return `changes ${args?.added}/${args?.removed}/${args?.modified}`;
    if (k === 'files.search.largeMark') return 'large';
    if (k === 'files.error.serviceUnavailable') return 'serviceUnavailable';
    if (k === 'doc.sessionExpired') return 'sessionExpired';
    if (k === 'doc.error.updatePage') return 'updateError';
    return k;
  }
}));

interface DriveFile { id: string; name: string; mimeType?: string; size?: number | string; modifiedTime?: string; owners?: Array<{ emailAddress?: string; displayName?: string } | undefined> }

function createConfig(): BotConfig {
  return {
    env: 'test',
    discord: { token: 'x', clientId: 'x', guildId: 'x', enableSlash: false, enableChat: false, enableMessageContentIntent: false },
    google: { apiKey: 'x', driveFolderId: 'root' },
    drive: { folderId: 'root', enableTextIndex: false, ttlTextSec: 60, indexCron: '* * * * *', allowedMime: [], ownerAllowlist: [], fileMaxSizeMb: 50 },
  } as unknown as BotConfig;
}

function files3(): DriveFile[] {
  const now = Date.now();
  const iso = (ms: number) => new Date(ms).toISOString();
  return [
    { id: 'f1', name: 'one', mimeType: 'text/plain', size: 1_000_000, modifiedTime: iso(now - 3*24*3600*1000) },
    { id: 'f2', name: 'two', mimeType: 'application/pdf', size: 2_000_000, modifiedTime: iso(now - 2*24*3600*1000) },
    { id: 'f3', name: 'three', mimeType: 'application/pdf', size: 3_000_000, modifiedTime: iso(now - 1*24*3600*1000) },
  ];
}

type GoogleMock = { listDriveFiles: jest.Mock<Promise<{ files: DriveFile[]; changes?: { addedIds: string[]; removedIds: string[]; modified: { id: string }[] } }>, any> };

type ChatOptions = { getString: (n: string) => string | null; getInteger: (n: string) => number | null };

function chatOptions(over?: Partial<ChatOptions>): ChatOptions {
  return { getString: () => null, getInteger: () => null, ...over } as ChatOptions;
}

function chatInteraction(over?: Partial<any>) {
  return {
    user: { id: 'u1' },
    channelId: 'c1',
    options: chatOptions({ getInteger: (n) => (n === 'ліміт' ? 2 : null) }),
    editReply: jest.fn(),
    deferReply: jest.fn(),
    replied: false,
    deferred: false,
    ...over,
  };
}

function componentInteraction(over?: Partial<any>) {
  return {
    isButton: () => true,
    deferred: false,
    replied: false,
    update: jest.fn(),
    editReply: jest.fn(),
    reply: jest.fn(),
    customId: '',
    channelId: 'c1',
    options: chatOptions(), // used by buildSearchPage
    ...over,
  };
}

function makeCmd(google: GoogleMock) {
  const cmd = new FileManagerCommand(createConfig());
  (cmd as any).getGoogleService = () => google;
  return cmd as any;
}

beforeAll(() => {
  jest.spyOn((Base as any).BaseCommand.prototype as any, 'startCleanupInterval').mockImplementation(() => {});
});

describe('FileManagerCommand onComponent buttons', () => {
  test('pagination next button updates to page 2/2', async () => {
    const google: GoogleMock = { listDriveFiles: jest.fn().mockResolvedValue({ files: files3(), changes: { addedIds: [], removedIds: [], modified: [] } }) } as any;
    const cmd = makeCmd(google);

    const ci = chatInteraction();
    await (cmd as any).handleSearch(ci, { query: 'all', folder: 'root' });
    const firstReply = ci.editReply.mock.calls[0][0];
    const nextBtn = firstReply.components[0].components[2]; // next
    const customId = nextBtn.data.custom_id || nextBtn.customId;

    const comp = componentInteraction({ customId });
    // ensure options exist for buildSearchPage
    comp.options = chatOptions({ getInteger: (n) => (n === 'ліміт' ? 2 : null) });

    await (cmd as any).onComponent({ interaction: comp });

    expect(google.listDriveFiles).toHaveBeenCalled();
    expect(comp.update).toHaveBeenCalled();
    const payload = comp.update.mock.calls[0][0];
    const footerText = (payload.embeds[0] as any).data?.footer?.text || (payload.embeds[0] as any).footer?.text;
    expect(footerText).toContain('2/2');
  });

  test('toggle changesOnly flips flag and updates', async () => {
    const google: GoogleMock = { listDriveFiles: jest.fn().mockResolvedValue({ files: files3(), changes: { addedIds: [], removedIds: [], modified: [] } }) } as any;
    const cmd = makeCmd(google);

    const ci = chatInteraction();
    await (cmd as any).handleSearch(ci, { query: 'all', folder: 'root' });
    const firstReply = ci.editReply.mock.calls[0][0];
    const toggleBtn = firstReply.components[1].components[0];
    const customId = toggleBtn.data.custom_id || toggleBtn.customId;

    const comp = componentInteraction({ customId });
    comp.options = chatOptions();

    // capture sid to check session
    const sid = (customId as string).split('|').find((p) => p.startsWith('sid='))!.slice(4);
    const before = (FileManagerCommand as any).sessions.get(sid).changesOnly;

    await (cmd as any).onComponent({ interaction: comp });

    const after = (FileManagerCommand as any).sessions.get(sid).changesOnly;
    expect(after).toBe(!before);
    expect(comp.update).toHaveBeenCalled();
  });

  test('reset baseline updates baseline and refreshes page', async () => {
    const google: GoogleMock = { listDriveFiles: jest.fn().mockResolvedValue({ files: files3(), changes: { addedIds: [], removedIds: [], modified: [] } }) } as any;
    const cmd = makeCmd(google);

    const ci = chatInteraction();
    await (cmd as any).handleSearch(ci, { query: 'all', folder: 'root' });
    const firstReply = ci.editReply.mock.calls[0][0];
    const resetBtn = firstReply.components[1].components[1];
    const customId = resetBtn.data.custom_id || resetBtn.customId;

    const sid = (customId as string).split('|').find((p) => p.startsWith('sid='))!.slice(4);
    const before = (FileManagerCommand as any).sessions.get(sid).baseline;

    const comp = componentInteraction({ customId });
    await (cmd as any).onComponent({ interaction: comp });

    const after = (FileManagerCommand as any).sessions.get(sid).baseline;
    expect(after).toBeGreaterThanOrEqual(before);
    expect(comp.update).toHaveBeenCalled();
  });

  test('close deletes session and removes components', async () => {
    const google: GoogleMock = { listDriveFiles: jest.fn().mockResolvedValue({ files: files3(), changes: { addedIds: [], removedIds: [], modified: [] } }) } as any;
    const cmd = makeCmd(google);

    const ci = chatInteraction();
    await (cmd as any).handleSearch(ci, { query: 'all', folder: 'root' });
    const firstReply = ci.editReply.mock.calls[0][0];
    const closeBtn = firstReply.components[1].components[2];
    const customId = closeBtn.data.custom_id || closeBtn.customId;
    const sid = (customId as string).split('|').find((p) => p.startsWith('sid='))!.slice(4);

    const comp = componentInteraction({ customId });
    await (cmd as any).onComponent({ interaction: comp });

    const session = (FileManagerCommand as any).sessions.get(sid);
    expect(session).toBeUndefined();
    expect(comp.update).toHaveBeenCalled();
    const payload = comp.update.mock.calls[0][0];
    expect(payload.components).toEqual([]);
  });
});
