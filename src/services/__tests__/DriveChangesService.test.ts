import { DriveChangesService, type DriveChangeEvent } from '../DriveChangesService';

// Minimal BotConfig shape for our service
const makeConfig = (over: Partial<any> = {}): any => {
  const driveOver = (over as any)?.['drive'] ?? {};
  return {
    drive: {
      folderId: 'FOLDER123',
      hideWebLink: true,
      ...driveOver,
    },
    google: { credentials: { client_email: 'x', private_key: 'y' } },
    ...over,
  };
};

// In-memory cache stub with same API used in service
class MemCache {
  store = new Map<string, any>();
  async get<T>(key: string): Promise<T | undefined> { return this.store.get(key); }
  async set(key: string, val: any, _ttlSec?: number): Promise<void> { this.store.set(key, val); }
}

// Mock provider implementing IDriveChangesProvider contract
class MockProvider {
  constructor(private opts: { start: string; pages: Array<{ token: string; changes: any[]; next?: string; newStartToken?: string }>; }) {}
  async getStartPageToken(): Promise<string> { return this.opts.start; }
  async listChanges(pageToken: string) {
    const page = this.opts.pages.find(p => p.token === pageToken);
    if (!page) return { changes: [], nextPageToken: undefined, newStartPageToken: undefined };
    return {
      changes: page.changes,
      nextPageToken: page.next,
      newStartPageToken: page.newStartToken,
    };
  }
}

const change = (id: string, type: 'created'|'modified'|'removed', folderId = 'FOLDER123') => {
  const baseFile = {
    id,
    name: `file-${id}`,
    parents: [folderId],
    owners: [{ emailAddress: 'u@example.com' }],
    webViewLink: 'https://drive.google.com/file/'+id,
    createdTime: '2020-01-01T00:00:00Z',
    modifiedTime: '2020-01-01T00:00:00Z',
  };
  if (type === 'removed') {
    return { removed: true, fileId: id, time: '2020-01-02T00:00:00Z', file: { ...baseFile, trashed: true } };
  }
  if (type === 'created') {
    // For created files, we should not have a modifiedTime to differentiate them
    const createdFile = { ...baseFile, modifiedTime: undefined };
    return { file: createdFile, time: baseFile.createdTime };
  }
  // modified
  return { file: { ...baseFile, modifiedTime: '2020-01-03T00:00:00Z' }, time: '2020-01-03T00:00:00Z' };
};

describe('DriveChangesService', () => {
  test('initialize stores startPageToken when empty', async () => {
    const cache = new MemCache();
    const provider = new MockProvider({ start: 'tok-1', pages: [] });
    const svc = new DriveChangesService(makeConfig(), provider as any, cache as any);

    await svc.initialize();

    expect(await cache.get('drive:changes:startPageToken')).toBe('tok-1');
  });

  test('pollOnce maps events and updates token to newStartPageToken', async () => {
    const cache = new MemCache();
    const provider = new MockProvider({
      start: 'tok-1',
      pages: [
        { token: 'tok-1', changes: [change('A','created'), change('B','modified'), change('C','removed')], next: 'tok-2' },
        { token: 'tok-2', changes: [change('X','modified','OTHER-FOLDER')], newStartToken: 'tok-NEW' },
      ],
    });
    const svc = new DriveChangesService(makeConfig(), provider as any, cache as any);

    // seed token via initialize
    await svc.initialize();

    const { events, newToken } = await svc.pollOnce();

    // Should filter out changes not in folder, and map 3 events
    expect(events.map(e => e.fileId)).toEqual(['A','B','C']);
    const types = new Map(events.map(e => [e.fileId, e.type] as [string, DriveChangeEvent['type']]));
    expect(types.get('A')).toBe('created');
    expect(types.get('B')).toBe('modified');
    expect(types.get('C')).toBe('removed');

    // webViewLink is hidden by default
    expect(events.every(e => e.webViewLink === undefined)).toBe(true);

    // Token must be updated to newStartToken
    expect(newToken).toBe('tok-NEW');
    expect(await cache.get('drive:changes:startPageToken')).toBe('tok-NEW');
  });
});
