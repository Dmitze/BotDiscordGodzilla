import { PlainTextParser } from '@/parsers/PlainTextParser';
import type { DriveFile } from '@/types/drive';

describe('PlainTextParser', () => {
  const parser = new PlainTextParser();

  const meta: DriveFile = {
    id: 'file-1',
    name: 'test.txt',
    mimeType: 'text/plain',
  };

  it('canParse supports text/plain, md, json, csv', () => {
    expect(parser.canParse({ ...meta, mimeType: 'text/plain' })).toBe(true);
    expect(parser.canParse({ ...meta, mimeType: 'text/markdown' })).toBe(true);
    expect(parser.canParse({ ...meta, mimeType: 'application/json' })).toBe(true);
    expect(parser.canParse({ ...meta, mimeType: 'text/csv' })).toBe(true);
    expect(parser.canParse({ ...meta, mimeType: 'application/pdf' })).toBe(false);
  });

  it('parse returns utf8 text, source="raw" and the original buffer', async () => {
    const sample = 'Привіт, світ!';
    const buf = Buffer.from(sample, 'utf8');

    const res = await parser.parse(
      { fileId: meta.id, mime: meta.mimeType },
      {
        // PlainTextParser only uses downloadFile
        downloadFile: async () => buf,
      } as any
    );

    expect(res.text).toBe(sample);
    expect(res.source).toBe('raw');
    expect(res.buffer).toBeInstanceOf(Buffer);
    expect(res.buffer?.equals(buf)).toBe(true);
  });

  it('throws if fileId is missing', async () => {
    await expect(
      parser.parse(
        { mime: meta.mimeType },
        { downloadFile: async () => Buffer.from('x') } as any
      )
    ).rejects.toThrow('fileId required');
  });
});
