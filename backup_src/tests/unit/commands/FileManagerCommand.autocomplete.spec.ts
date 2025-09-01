/* eslint-disable @typescript-eslint/no-unsafe-assignment */
/* eslint-disable @typescript-eslint/no-unsafe-member-access */
/* eslint-disable @typescript-eslint/no-unsafe-call */
/* eslint-disable @typescript-eslint/no-explicit-any */

import { FileManagerCommand } from '@/commands/FileManagerCommand';
import * as Base from '@/commands/BaseCommand';
import type { BotConfig } from '@/types';

jest.mock('@/i18n', () => ({ t: (k: string) => k }));

function createConfig(): BotConfig {
  return {
    env: 'test',
    discord: { token: 'x', clientId: 'x', guildId: 'x', enableSlash: false, enableChat: false, enableMessageContentIntent: false },
    google: { apiKey: 'x', driveFolderId: 'team-folder' },
    drive: { folderId: 'root', enableTextIndex: false, ttlTextSec: 60, indexCron: '* * * * *', allowedMime: ['application/pdf'], ownerAllowlist: [], fileMaxSizeMb: 50 },
  } as unknown as BotConfig;
}

beforeAll(() => {
  jest.spyOn((Base as any).BaseCommand.prototype as any, 'startCleanupInterval').mockImplementation(() => {});
});

function makeAutocompleteInteraction(name: string, value: string) {
  const responded: any[] = [];
  return {
    options: {
      getFocused: (withName?: boolean) => (withName ? { name, value } : value),
    },
    respond: (choices: any[]) => {
      responded.splice(0, responded.length, ...choices);
      return Promise.resolve();
    },
    __getChoices: () => responded,
  } as any;
}

describe('FileManagerCommand onAutocomplete', () => {
  test('suggests MIME values including from config and defaults', async () => {
    const cmd = new FileManagerCommand(createConfig());
    const ia = makeAutocompleteInteraction('mime', 'pdf');
    await (cmd as any).onAutocomplete({ interaction: ia });
    const choices = ia.__getChoices();
    const values = choices.map((c: any) => c.value);
    expect(values).toContain('application/pdf');
    // default list also includes google docs/spreadsheet; filtered by "pdf" we still keep 'application/pdf'
  });

  test('suggests folders: root and configured driveFolderId', async () => {
    const cmd = new FileManagerCommand(createConfig());
    const ia = makeAutocompleteInteraction('папка', '');
    await (cmd as any).onAutocomplete({ interaction: ia });
    const choices = ia.__getChoices();
    const values = choices.map((c: any) => c.value);
    expect(values).toEqual(expect.arrayContaining(['root', 'team-folder']));
  });

  test('suggests query presets', async () => {
    const cmd = new FileManagerCommand(createConfig());
    const ia = makeAutocompleteInteraction('запит', 'type:');
    await (cmd as any).onAutocomplete({ interaction: ia });
    const choices = ia.__getChoices();
    const values = choices.map((c: any) => c.value);
    expect(values.some((v: string) => v.startsWith('type:'))).toBe(true);
  });
});
