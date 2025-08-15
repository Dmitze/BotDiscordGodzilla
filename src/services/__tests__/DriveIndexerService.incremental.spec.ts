import { createBotMock, createCacheMock, createGoogleMock, initIndexer } from './__utils__/driveIndexerTestHelpers';
import { fileDoc, filePdf, clone } from '../../tests/fixtures/drive/files';

const KEY = (id: string) => `drive:index:file:${id}`;

describe('DriveIndexerService - incremental index', () => {
  test('indexes only new/changed files by modifiedTime', async () => {
    const google = createGoogleMock();
    const cache = createCacheMock();

    // seed cache: fileDoc already indexed with same modifiedTime => skip
    await cache.set(KEY('doc1'), {
      id: 'doc1', name: 'Doc One', mimeType: fileDoc.mimeType, text: 'OLD', textLength: 3, updatedAt: Date.now(), modifiedTime: fileDoc.modifiedTime,
    });
    await cache.set('drive:index:keys', ['doc1']);

    const changedPdf = clone(filePdf);
    changedPdf.modifiedTime = '2025-08-20T10:00:00Z';

    google.listDriveFiles
      .mockResolvedValueOnce({ files: [fileDoc, changedPdf], nextPageToken: undefined });

    google.extractTextFromFile
      .mockImplementation(async ({ id }: any) => `NEW_${id}`);

    const bot = createBotMock(google as any, cache as any);
    const indexer = await initIndexer(bot);

    await indexer.reindexIncremental('root');

    // doc1 unchanged, pdf1 changed
    expect(google.extractTextFromFile).toHaveBeenCalledTimes(1);
    expect(google.extractTextFromFile).toHaveBeenCalledWith(expect.objectContaining({ id: 'pdf1' }));

    const doc = await cache.get<any>(KEY('doc1'));
    const pdf = await cache.get<any>(KEY('pdf1'));

    expect(doc?.text).toBe('OLD');
    expect(pdf?.text).toBe('NEW_pdf1');

    const keys = await cache.get<string[]>('drive:index:keys');
    expect(keys).toEqual(expect.arrayContaining(['doc1', 'pdf1']));
  });
});
