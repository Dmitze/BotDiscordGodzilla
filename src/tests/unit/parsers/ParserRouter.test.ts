import type { DriveFile } from '@/types/drive';
import { ParserRouter } from '@/parsers/ParserRouter';
import type { IParser, ParseResult } from '@/parsers/IParser';
import { setLocale } from '@/i18n';

describe('ParserRouter', () => {
  beforeAll(() => {
    // Ensure default is Ukrainian for deterministic i18n
    setLocale('uk');
  });

  const meta: DriveFile = {
    id: 'id-1',
    name: 'doc.txt',
    mimeType: 'text/plain',
  };

  const ctx = {
    exportFile: async () => Buffer.from(''),
    downloadFile: async () => Buffer.from('hello'),
    extractTextFromImage: async () => 'image-text',
    extractTextFromBuffer: async () => 'buf-text',
  };

  function makeParser(name: string, opts: { can: boolean; fail?: boolean; text?: string }): IParser {
    return {
      canParse: () => opts.can,
      parse: async (): Promise<ParseResult> => {
        if (opts.fail) throw new Error(name + ' failed');
        return { text: opts.text ?? name, source: 'parser' } as ParseResult;
      },
    } as IParser;
  }

  it('uses first matching parser on success', async () => {
    const p1 = makeParser('p1', { can: true, text: 'ok1' });
    const p2 = makeParser('p2', { can: true, text: 'ok2' });
    const router = new ParserRouter([p1, p2]);

    const res = await router.parse(meta, ctx);
    expect(res.text).toBe('ok1');
  });

  it('falls back to next parser when first throws', async () => {
    const p1 = makeParser('p1', { can: true, fail: true });
    const p2 = makeParser('p2', { can: true, text: 'ok2' });
    const router = new ParserRouter([p1, p2]);

    const res = await router.parse(meta, ctx);
    expect(res.text).toBe('ok2');
  });

  it('throws i18n message on unsupported mime', async () => {
    const router = new ParserRouter([
      makeParser('p1', { can: false }),
      makeParser('p2', { can: false }),
    ]);
    const unsupported: DriveFile = { ...meta, mimeType: 'application/unknown' };

    await expect(router.parse(unsupported, ctx)).rejects.toThrow('Непідтримуваний MIME');
  });
});
