import { sanitizeTextForChat, chunkTextForDiscord } from '@/utils/fileProcessor';
import { Config } from '@/config/Config';

describe('PII masking flags in fileProcessor', () => {
  const sample = 'Contact me at john.doe@example.com or +1 (415) 555-2671';

  afterEach(() => {
    jest.restoreAllMocks();
  });

  it('masks email and phone when flags ON', () => {
    jest.spyOn(Config, 'get').mockReturnValue({
      features: { enablePiiMasking: true, piiMaskEmail: true, piiMaskPhone: true },
    } as any);
    const out = sanitizeTextForChat(sample);
    expect(out).not.toContain('john.doe@example.com');
    expect(out).not.toMatch(/415\D*555/);
  });

  it('masks only email when phone OFF', () => {
    jest.spyOn(Config, 'get').mockReturnValue({
      features: { enablePiiMasking: true, piiMaskEmail: true, piiMaskPhone: false },
    } as any);
    const out = sanitizeTextForChat(sample);
    expect(out).not.toContain('john.doe@example.com');
    expect(out).toMatch(/415\D*555/);
  });

  it('masks only phone when email OFF', () => {
    jest.spyOn(Config, 'get').mockReturnValue({
      features: { enablePiiMasking: true, piiMaskEmail: false, piiMaskPhone: true },
    } as any);
    const out = sanitizeTextForChat(sample);
    expect(out.includes('example.com') || /john(\.|\*)?doe/i.test(out)).toBe(true);
    expect(out).not.toMatch(/415\D*555/);
  });

  it('no masking when master switch OFF', () => {
    jest.spyOn(Config, 'get').mockReturnValue({
      features: { enablePiiMasking: false, piiMaskEmail: true, piiMaskPhone: true },
    } as any);
    const out = sanitizeTextForChat(sample);
    expect(out).toContain('john.doe@example.com');
    expect(out).toMatch(/415\D*555/);
  });

  it('applies same logic in chunkTextForDiscord', () => {
    jest.spyOn(Config, 'get').mockReturnValue({
      features: { enablePiiMasking: true, piiMaskEmail: true, piiMaskPhone: true },
    } as any);
    const chunks = chunkTextForDiscord(sample, { maxChunkLen: 500 });
    expect(chunks.length).toBe(1);
    const out = chunks[0]!;
    expect(out).not.toContain('john.doe@example.com');
    expect(out).not.toMatch(/415\D*555/);
  });
});
