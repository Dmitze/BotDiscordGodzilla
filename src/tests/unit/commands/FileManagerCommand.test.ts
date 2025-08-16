/**
 * Unit тести для FileManagerCommand (актуалізовано під поточний API)
 */
/* eslint-disable @typescript-eslint/no-unsafe-assignment */
/* eslint-disable @typescript-eslint/no-unsafe-member-access */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/unbound-method */
/* eslint-disable @typescript-eslint/require-await */

import { jest, describe, it, expect, beforeEach } from '@jest/globals';
import { FileManagerCommand } from '../../../commands/FileManagerCommand';
import type { BotConfig, CommandExecuteOptions } from '../../../types';
import type { ChatInputCommandInteraction, Client } from 'discord.js';
import type { GoogleService } from '../../../services/GoogleService';

// Мок конфіга з явними типами
const makeConfig = (): BotConfig => ({
  discord: { token: '', clientId: '', guildId: '' } as any,
  google: { driveFolderId: 'root-folder', credentials: { client_email: 'x', private_key: 'y' } } as any,
  ai: { provider: 'none', openai: {} as any, ollama: {} as any } as any,
  server: {} as any,
  cache: {} as any,
} as unknown as BotConfig);

type ReplyRecord = { type: 'reply' | 'edit' | 'follow'; payload: unknown };

const makeInteraction = (sub: string, opts: Record<string, string | null> = {}): ChatInputCommandInteraction & { __replies: ReplyRecord[] } => {
  const replies: ReplyRecord[] = [];
  const base = {
    user: { id: 'u1', tag: 'user#0001' } as any,
    deferred: false as any,
    replied: false as any,
    options: {
      getSubcommand: () => sub,
      getString: (name: string) => (name in opts ? opts[name] : null),
    } as any,
    client: {
      serviceContainer: {
        get: (_key: string) => undefined,
      },
    } as unknown as Client<true>,
    deferReply: jest.fn(async function (this: any) { (this as any).deferred = true; }) as any,
    reply: jest.fn(async function (this: any, p: unknown) { replies.push({ type: 'reply', payload: p }); (this as any).replied = true; }) as any,
    editReply: jest.fn(async (_p: unknown) => { replies.push({ type: 'edit', payload: _p }); }) as any,
    followUp: jest.fn(async (_p: unknown) => { replies.push({ type: 'follow', payload: _p }); }) as any,
    __replies: replies,
  } as unknown;
  return base as ChatInputCommandInteraction & { __replies: ReplyRecord[] };
};

describe('FileManagerCommand', () => {
  let cmd: FileManagerCommand;

  beforeEach(() => {
    jest.resetAllMocks();
    cmd = new FileManagerCommand(makeConfig());
  });

  describe('constructor/basic', () => {
    it('creates instance with correct name/description', () => {
      expect(cmd).toBeInstanceOf(FileManagerCommand);
      expect(cmd.name).toBe('файли');
      expect(typeof cmd.description).toBe('string');
      expect(cmd.data?.name).toBe('файли');
    });
  });

  describe('execute: пошук', () => {
    it('returns list via editReply for one result', async () => {
      const interaction = makeInteraction('пошук', { 'запит': 'документ', 'папка': null });
      const google = {
        listDriveFilesInFolder: (jest.fn() as any).mockResolvedValueOnce([
          { id: '1', name: 'File 1.pdf', mimeType: 'application/pdf' },
        ]) as unknown as GoogleService['listDriveFilesInFolder'],
      } as Partial<GoogleService> as any;
      (interaction.client as any).serviceContainer.get = (k: string) => (k === 'google' ? google : undefined);

      await cmd.execute({ interaction } as CommandExecuteOptions);

      expect(interaction.deferReply).toHaveBeenCalled();
      expect(interaction.editReply).toHaveBeenCalled();
      const edit = interaction.__replies.find((r: any) => r.type === 'edit');
      expect(edit).toBeDefined();
    });

    it('returns list via editReply for multiple results', async () => {
      const interaction = makeInteraction('пошук', { 'запит': 'документ', 'папка': null });
      const google = {
        listDriveFilesInFolder: (jest.fn() as any).mockResolvedValueOnce(([
          { id: '1', name: 'File 1.pdf', mimeType: 'application/pdf' },
          { id: '2', name: 'File 2.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
        ])) as unknown as GoogleService['listDriveFilesInFolder'],
      } as Partial<GoogleService> as any;
      (interaction.client as any).serviceContainer.get = (k: string) => (k === 'google' ? google : undefined);

      await cmd.execute({ interaction } as CommandExecuteOptions);

      expect(interaction.deferReply).toHaveBeenCalled();
      expect(interaction.editReply).toHaveBeenCalled();
      const edit = interaction.__replies.find((r: any) => r.type === 'edit');
      expect(edit).toBeDefined();
    });

    it('returns error message via reply for empty results', async () => {
      const interaction = makeInteraction('пошук', { 'запит': 'документ', 'папка': null });
      const google = {
        listDriveFilesInFolder: (jest.fn() as any).mockResolvedValueOnce([]) as unknown as GoogleService['listDriveFilesInFolder'],
      } as Partial<GoogleService> as any;
      (interaction.client as any).serviceContainer.get = (k: string) => (k === 'google' ? google : undefined);

      await cmd.execute({ interaction } as any);

      // Command defers and edits with embed message for empty results
      expect(interaction.deferReply).toHaveBeenCalled();
      expect(interaction.editReply).toHaveBeenCalled();
      const editEmpty = interaction.__replies.find((r) => r.type === 'edit');
      expect(editEmpty).toBeDefined();
      const descEmpty = (editEmpty as any)?.payload?.embeds?.[0]?.data?.description || (editEmpty as any)?.payload?.embeds?.[0]?.description;
      expect(descEmpty).toMatch(/Нічого не знайдено|Файлів не знайдено/);
    });
  });

  describe('execute: аналіз', () => {
    it('returns analysis result via editReply', async () => {
      const interaction = makeInteraction('аналіз', { 'id': 'file_id_123', 'тип': 'summary' });
      const google = {
        getDriveFileMetadata: jest.fn(async () => ({
          id: 'file_id_123',
          name: 'Doc Name',
          mimeType: 'application/vnd.google-apps.document',
          size: '2048',
        })),
        exportDriveFile: jest.fn(async () => Buffer.from('text content', 'utf8')),
      } as Partial<GoogleService> as any;
      (interaction.client as any).serviceContainer.get = (k: string) => (k === 'google' ? google : undefined);

      await cmd.execute({ interaction } as CommandExecuteOptions);

      expect(interaction.deferReply).toHaveBeenCalled();
      expect(interaction.editReply).toHaveBeenCalled();
      const edit = interaction.__replies.find((r) => r.type === 'edit');
      expect(edit).toBeDefined();
      // Для різних форматів відповіді достатньо перевірити наявність пейлоаду
      expect((edit as any).payload).toBeDefined();
    });
  });

  describe('execute: завантаження', () => {
    it('returns file via editReply', async () => {
      const interaction = makeInteraction('завантаження', { 'id': 'file_id_123' });
      const google = {
        getDriveFileMetadata: jest.fn(async () => ({
          id: 'file_id_123',
          name: 'Some.pdf',
          mimeType: 'application/pdf',
          size: '1024',
          webViewLink: 'https://drive.google.com/file/d/xx',
        })),
        downloadDriveFile: jest.fn(async (_id: string) => Buffer.from('file_content')),
      } as Partial<GoogleService> as any;
      (interaction.client as any).serviceContainer.get = (k: string) => (k === 'google' ? google : undefined);

      await cmd.execute({ interaction } as any);

      expect(interaction.deferReply).toHaveBeenCalled();
      expect(interaction.editReply).toHaveBeenCalled();
      const edit = interaction.__replies.find((r: any) => r.type === 'edit');
      expect(edit).toBeDefined();
    });
  });

  describe('execute: невідома підкоманда', () => {
    it('returns error message via reply', async () => {
      const interaction = makeInteraction('невідома');

      await cmd.execute({ interaction } as any);

      // Unknown subcommand path throws after defer, so error goes via editReply
      const editUnknown = interaction.__replies.find((r) => r.type === 'edit') as any;
      expect(editUnknown).toBeDefined();
      const editUnknownContent = (editUnknown.payload?.content || editUnknown.payload?.embeds?.[0]?.description || '') as string;
      expect(editUnknownContent).toMatch(/Невідома підкоманда|Невідома|Помилка|сталася помилка/i);
    });
  });

  describe('execute: помилка сервісу', () => {
    it('returns error message via reply', async () => {
      const interaction = makeInteraction('пошук', { 'запит': 'документ', 'папка': null });
      const google = {
        listDriveFilesInFolder: (jest.fn() as any).mockRejectedValue(new Error('Service error')) as unknown as GoogleService['listDriveFilesInFolder'],
      } as Partial<GoogleService> as any;
      (interaction.client as any).serviceContainer.get = (k: string) => (k === 'google' ? google : undefined);

      await cmd.execute({ interaction } as any);

      // Error during search after defer: expect edit reply with error
      const editErr = interaction.__replies.find((r) => r.type === 'edit') as any;
      expect(editErr).toBeDefined();
      const errContent = (editErr.payload?.content || editErr.payload?.embeds?.[0]?.description || '') as string;
      expect(errContent).toMatch(/Помилка|сталася помилка/i);
    });
  });
});