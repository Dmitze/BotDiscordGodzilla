import { GoogleService } from '@/services/GoogleService';
import type { BotConfig } from '@/types';
import type { drive_v3 } from 'googleapis';
import pdfParse from 'pdf-parse';
import * as mammoth from 'mammoth';

// Mocks for external parsers (return Promises without async arrow to satisfy lint)
jest.mock('pdf-parse', () => jest.fn((_buf: Buffer) => Promise.resolve({ text: 'PDF TEXT' })));
jest.mock('mammoth', () => ({
  extractRawText: jest.fn((_input: { buffer: Buffer }) => Promise.resolve({ value: 'DOCX TEXT' })),
}));

// Helper to build minimal config
function createConfig(): BotConfig {
  return {
    env: 'test',
    google: {
      ocrProvider: 'off',
      ocrCacheTTL: 300,
    } as any,
    drive: {
      pageSize: 10,
      ttlTextSec: 300,
      allowedMime: ['*'],
    } as any,
    performance: {
      cacheTTL: 300,
    } as any,
  } as unknown as BotConfig;
}

// Simple in-memory cache mock compatible with CacheService API used
function createCacheMock() {
  const map = new Map<string, unknown>();
  return {
    get: jest.fn(async (k: string) => map.get(k)),
    set: jest.fn(async (k: string, v: unknown) => {
      map.set(k, v);
    }),
  };
}

// Utility to stub internal methods
function stubInternals(svc: GoogleService) {
  // executeWithRetry just runs the fn
  // @ts-expect-error access private
  svc.executeWithRetry = (fn: any) => fn();
  // @ts-expect-error access private
  svc.throttle = async () => 0;
}

describe('GoogleService.extractTextForChat', () => {
  it('exports Google Docs to text/plain', async () => {
    const svc = new GoogleService(createConfig());
    // @ts-expect-error replace cache
    svc.cacheService = createCacheMock();
    stubInternals(svc);

    const fileId = 'doc123';
    const modifiedTime = '2025-01-01T00:00:00Z';

    // Stubs
    const metaDoc: drive_v3.Schema$File = {
      id: fileId,
      mimeType: 'application/vnd.google-apps.document',
      modifiedTime,
    };
    jest.spyOn(svc, 'getDriveFileMetadata').mockResolvedValue(metaDoc);
    const exportSpy = jest
      .spyOn(svc, 'exportFile')
      .mockResolvedValue(Buffer.from('Hello from DOCS'));

    const res = await svc.extractTextForChat(fileId);
    expect(exportSpy).toHaveBeenCalledWith(fileId, 'text/plain');
    expect(res.text).toContain('Hello from DOCS');
    expect(res.source).toBe('export');
    expect(res.modifiedTime).toBe(modifiedTime);
    expect(res.checksum).toMatch(/^[a-f0-9]{64}$/);
  });

  it('parses PDF via pdf-parse, then falls back to OCR on error', async () => {
    const svc = new GoogleService(createConfig());
    // @ts-expect-error replace cache
    svc.cacheService = createCacheMock();
    stubInternals(svc);

    const fileId = 'pdf123';

    const metaSpy = jest.spyOn(svc, 'getDriveFileMetadata');
    metaSpy
      .mockResolvedValueOnce({
        id: fileId,
        mimeType: 'application/pdf',
        modifiedTime: '2025-01-02T00:00:00Z',
      } as any)
      .mockResolvedValueOnce({
        id: fileId,
        mimeType: 'application/pdf',
        modifiedTime: '2025-01-02T00:01:00Z', // змінюємо, щоб обійти кеш
      } as any);

    const dlSpy = jest
      .spyOn(svc, 'downloadFile')
      .mockResolvedValue(Buffer.from('PDF BUFF'));

    // Case 1: success via pdf-parse
    (pdfParse as unknown as jest.Mock).mockResolvedValueOnce({ text: 'PDF TEXT OK' });

    let res = await svc.extractTextForChat(fileId);
    expect(dlSpy).toHaveBeenCalled();
    expect(res.text).toContain('PDF TEXT OK');
    expect(res.source).toBe('parser');

    // Case 2: pdf-parse throws -> OCR used (modifiedTime змінено, кеш не спрацює)
    (pdfParse as unknown as jest.Mock).mockRejectedValueOnce(new Error('boom'));
    const ocrSpy = jest
      .spyOn(svc, 'extractTextFromBuffer')
      .mockResolvedValue('OCR TEXT');

    res = await svc.extractTextForChat(fileId);
    expect(ocrSpy).toHaveBeenCalled();
    expect(res.text).toContain('OCR TEXT');
    expect(res.source).toBe('ocr');
  });

  it('parses DOCX via mammoth with fallback to raw text', async () => {
    const svc = new GoogleService(createConfig());
    // @ts-expect-error replace cache
    svc.cacheService = createCacheMock();
    stubInternals(svc);

    const fileId = 'docx123';

    const metaSpy = jest.spyOn(svc, 'getDriveFileMetadata');
    metaSpy
      .mockResolvedValueOnce({
        id: fileId,
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        modifiedTime: '2025-01-03T00:00:00Z',
      } as any)
      .mockResolvedValueOnce({
        id: fileId,
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        modifiedTime: '2025-01-03T00:01:00Z', // змінюємо, щоб обійти кеш
      } as any);

    const dlSpy = jest
      .spyOn(svc, 'downloadFile')
      .mockResolvedValue(Buffer.from('DOCX BUFF'));

    // Case 1: mammoth success
    (mammoth.extractRawText as unknown as jest.Mock).mockResolvedValueOnce({ value: 'DOCX TEXT' });
    let res = await svc.extractTextForChat(fileId);
    expect(dlSpy).toHaveBeenCalled();
    expect(res.text).toContain('DOCX TEXT');
    expect(res.source).toBe('parser');

    // Case 2: mammoth fails -> raw
    (mammoth.extractRawText as unknown as jest.Mock).mockRejectedValueOnce(new Error('fail'));
    res = await svc.extractTextForChat(fileId);
    expect(res.text).toContain('DOCX BUFF');
    expect(res.source).toBe('raw');
  });

  it('uses cache across same modifiedTime', async () => {
    const svc = new GoogleService(createConfig());
    const cache = createCacheMock();
    // @ts-expect-error replace cache
    svc.cacheService = cache;
    stubInternals(svc);

    const fileId = 'text123';
    const modifiedTime = '2025-01-04T00:00:00Z';

    jest.spyOn(svc, 'getDriveFileMetadata').mockResolvedValue({
      id: fileId,
      mimeType: 'text/plain',
      modifiedTime,
    } as any);

    const dlSpy = jest
      .spyOn(svc, 'downloadFile')
      .mockResolvedValue(Buffer.from('HELLO TEXT'));

    const first = await svc.extractTextForChat(fileId);
    expect(first.text).toContain('HELLO TEXT');

    // Second call: should hit cache, not call download again
    dlSpy.mockClear();
    const second = await svc.extractTextForChat(fileId);
    expect(dlSpy).not.toHaveBeenCalled();
    expect(second.text).toBe(first.text);
  });
});
