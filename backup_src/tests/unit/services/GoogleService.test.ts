import { GoogleService } from '@/services/GoogleService';

describe('GoogleService.listDriveFiles normalization and owner filter', () => {
  const baseConfig: any = {
    drive: {
      pageSize: 10,
      ttlListSec: 60,
      allowedMime: ['*'],
      hideWebLink: false,
      rateQps: 5,
      rateBurst: 10,
    },
  };

  function makeSvc() {
    const svc = new GoogleService(baseConfig as any);
    (svc as any).drive = {
      files: {
        list: jest.fn().mockResolvedValue({
          data: {
            nextPageToken: undefined,
            files: [
              {
                id: '1',
                name: 'Doc A',
                mimeType: 'application/pdf',
                size: '1234',
                modifiedTime: '2024-01-01T00:00:00.000Z',
                owners: [{ emailAddress: 'owner@example.com', displayName: 'Owner' }],
                webViewLink: 'https://drive/1',
                iconLink: 'https://icon/1',
              },
              {
                id: '2',
                name: 'Sheet B',
                mimeType: 'application/vnd.google-apps.spreadsheet',
                size: '0',
                modifiedTime: '2024-02-01T00:00:00.000Z',
                owners: [{ emailAddress: 'other@example.com' }],
                webViewLink: 'https://drive/2',
              },
            ],
          },
        }),
      },
    };
    // обойти пул и rate-limit
    (svc as any).executeWithRetry = async (cb: any) => cb();
    (svc as any).throttle = async () => 0;

    // подмена кэша на Map
    const store = new Map<string, any>();
    (svc as any).cacheService = {
      async get<T>(key: string): Promise<T | undefined> { return store.get(key); },
      async set<T>(key: string, value: T, _ttlSec?: number) { store.set(key, value); },
    };

    return svc;
  }

  it('normalizes Drive files and applies ownerAllowlist', async () => {
    const svc = makeSvc();
    const res = await svc.listDriveFiles({ folderId: 'FOLDER', ownerAllowlist: ['owner@example.com'] } as any);
    // eslint-disable-next-line no-console
    console.log('DBG_RES:\\n' + JSON.stringify(res, null, 2));
    expect(res.files).toHaveLength(1);
    const f = res.files[0]!;
    expect(f.id).toBe('1');
    expect(f.name).toBe('Doc A');
    expect(f.mimeType).toBe('application/pdf');
    expect(f.size).toBe(1234);
    expect(f.modifiedTime).toBe('2024-01-01T00:00:00.000Z');
    expect(Array.isArray(f.owners)).toBe(true);
    expect(f.owners?.some(o => o.toLowerCase() === 'owner@example.com')).toBe(true);
    expect(typeof f.webViewLink === 'undefined' || f.webViewLink === 'https://drive/1').toBe(true);
  });

  it('caches list results (no second API call)', async () => {
    const svc = makeSvc();
    const listSpy = (svc as any).drive.files.list as jest.Mock;
    const key = { folderId: 'FOLDER' } as any;
    const first = await svc.listDriveFiles(key);
    expect(first.files.length).toBeGreaterThan(0);
    expect(listSpy).toHaveBeenCalledTimes(1);
    const second = await svc.listDriveFiles(key);
    expect(second.files.length).toBe(first.files.length);
    expect(listSpy).toHaveBeenCalledTimes(1);
  });
});
