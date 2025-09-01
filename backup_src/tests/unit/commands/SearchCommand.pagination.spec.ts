/* eslint-disable @typescript-eslint/no-unsafe-assignment */
/* eslint-disable @typescript-eslint/no-unsafe-member-access */
/* eslint-disable @typescript-eslint/no-unsafe-call */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/consistent-type-imports */
import { SearchCommand } from '@/commands/SearchCommand';
import type { BotConfig } from '@/types';

jest.mock('@/i18n', () => ({
  t: (k: string, args?: any) => {
    if (k.startsWith('files.search.buttons.')) return k.split('.').pop();
    if (k === 'search.log.start') return 'start';
    if (k === 'search.log.success') return 'success';
    if (k === 'files.error.process') return 'process';
    if (k === 'doc.sessionExpired') return 'expired';
    return k;
  }
}));

function createConfig(): BotConfig {
  return {
    env: 'test',
    discord: { token: 'x', clientId: 'x', guildId: 'x', enableSlash: false, enableChat: false, enableMessageContentIntent: false },
    google: { apiKey: 'x', driveFolderId: 'root' },
    drive: { folderId: 'root', enableTextIndex: false, ttlTextSec: 60, indexCron: '* * * * *', allowedMime: [], ownerAllowlist: [], fileMaxSizeMb: 50 },
  } as unknown as BotConfig;
}

function makeResult() {
  return {
    rows: [ ['a','b'], ['c','d'], ['e','f'], ['g','h'] ],
    headers: ['h1','h2'],
    totalCount: 4,
    filteredCount: 4,
    searchTime: 10,
    cacheHit: false,
    query: 'q',
    filters: { documentType: 'all', priority: 'all', limit: 2 },
  };
}

function makeInteraction(customId?: string) {
  const obj: any = {
    isButton: () => Boolean(customId),
    customId: customId ?? '',
    deferred: false,
    replied: false,
    update: jest.fn(),
    editReply: jest.fn(),
    reply: jest.fn(),
    followUp: jest.fn(),
  };
  return obj;
}

describe('SearchCommand pagination onComponent', () => {
  test('updates message on valid srch button (page change)', async () => {
    const cmd = new SearchCommand(createConfig());

    // Prepare session manually
    const sid = 'srch_abc_zzz';
    const state = {
      currentPage: 1,
      totalPages: 3,
      results: makeResult(),
      timestamp: Math.floor(Date.now() / 1000),
      userId: 'u1',
    } as any;
    (SearchCommand as any).sessions.set(sid, state);

    const ts = Math.floor(Date.now() / 1000);
    const interaction = makeInteraction(`srch|sid=${sid}|p=2|t=${ts}`);

    await (cmd as any).onComponent({ interaction } as any);

    // should update with embeds+components
    expect(interaction.update).toHaveBeenCalled();
    const arg = interaction.update.mock.calls[0][0];
    expect(arg.embeds).toBeTruthy();
    expect(arg.components).toBeTruthy();
    // session currentPage moves to 2
    expect((SearchCommand as any).sessions.get(sid).currentPage).toBe(2);
  });

  test('replies expired if TTL exceeded', async () => {
    const cmd = new SearchCommand(createConfig());
    const sid = 'srch_old_xxx';
    const oldTs = Math.floor(Date.now() / 1000) - (11 * 60);
    const state = {
      currentPage: 1,
      totalPages: 2,
      results: makeResult(),
      timestamp: oldTs,
      userId: 'u1',
    } as any;
    (SearchCommand as any).sessions.set(sid, state);

    const interaction = makeInteraction(`srch|sid=${sid}|p=1|t=${oldTs}`);
    await (cmd as any).onComponent({ interaction } as any);

    expect(interaction.reply).toHaveBeenCalled();
    const arg = interaction.reply.mock.calls[0][0];
    expect(arg.content).toContain('expired');
  });

  test('closes components on close action', async () => {
    const cmd = new SearchCommand(createConfig());
    const sid = 'srch_close_xxx';
    const ts = Math.floor(Date.now() / 1000);
    const state = {
      currentPage: 1,
      totalPages: 1,
      results: makeResult(),
      timestamp: ts,
      userId: 'u1',
    } as any;
    (SearchCommand as any).sessions.set(sid, state);

    const interaction = makeInteraction(`srch|sid=${sid}|p=1|a=close|t=${ts}`);
    await (cmd as any).onComponent({ interaction } as any);

    expect(interaction.update).toHaveBeenCalled();
    const arg = interaction.update.mock.calls[0][0];
    expect(arg.components).toEqual([]);
    expect((SearchCommand as any).sessions.has(sid)).toBe(false);
  });
});
