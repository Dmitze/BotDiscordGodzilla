import { GoogleService } from '@/services/GoogleService';
import { Config } from '@/config/Config';
import { createHash } from 'crypto';

describe('GoogleService.extractTextForChat PII masking + checksum + cache', () => {
  const sample = 'Email: john.doe@example.com, Phone: +1 415-555-1212';
  const buf = Buffer.from(sample, 'utf8');

  function makeSvc() {
    jest.spyOn(Config, 'get').mockReturnValue({
      drive: { ttlTextSec: 60 },
      features: { enablePiiMasking: true, piiMaskEmail: true, piiMaskPhone: true },
    } as any);

    const svc = new GoogleService({} as any);
    // Disable rate limit / retry wrappers
    (svc as any).executeWithRetry = async (cb: any) => cb();
    (svc as any).throttle = async () => 0;

    // Minimal drive mocks
    (svc as any).getDriveFileMetadata = jest.fn().mockResolvedValue({
      id: 'f1', mimeType: 'text/plain', modifiedTime: '2024-01-01T00:00:00.000Z',
    });
    (svc as any).downloadFile = jest.fn().mockResolvedValue(buf);

    // In-memory cache
    const store = new Map<string, any>();
    (svc as any).cacheService = {
      async get<T>(key: string): Promise<T | undefined> { return store.get(key); },
      async set<T>(key: string, value: T, _ttl?: number) { store.set(key, value); },
    };
    return svc;
  }

  it('masks PII in returned text while checksum is from original buffer', async () => {
    const svc = makeSvc();
    const res = await svc.extractTextForChat('f1');
    // Text should be masked
    expect(res.text).not.toContain('john.doe@example.com');
    expect(res.text).not.toMatch(/415\D*555/);
    // Checksum should be from buffer
    const expected = createHash('sha256').update(buf).digest('hex');
    expect(res.checksum).toBe(expected);
    expect(res.source).toBe('raw');
  });

  it('uses cache on second call (no extra download)', async () => {
    const svc = makeSvc();
    const dl = (svc as any).downloadFile as jest.Mock;
    const first = await svc.extractTextForChat('f1');
    expect(first.text.length).toBeGreaterThan(0);
    expect(dl).toHaveBeenCalledTimes(1);
    const second = await svc.extractTextForChat('f1');
    expect(second.text).toBe(first.text);
    expect(second.checksum).toBe(first.checksum);
    expect(dl).toHaveBeenCalledTimes(1);
  });

  it('respects flags: master OFF disables masking', async () => {
    const svc = makeSvc();
    (Config.get as jest.Mock).mockReturnValue({
      drive: { ttlTextSec: 60 },
      features: { enablePiiMasking: false, piiMaskEmail: true, piiMaskPhone: true },
    } as any);
    const res = await svc.extractTextForChat('f1');
    expect(res.text).toContain('john.doe@example.com');
    expect(res.text).toMatch(/415\D*555/);
  });
});
