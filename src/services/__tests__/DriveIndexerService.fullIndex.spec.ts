import { createBotMock, createCacheMock, createGoogleMock, createMetricsMock, initIndexer } from './__utils__/driveIndexerTestHelpers';
import { fileDoc, filePdf, fileWord, clone } from '../../tests/fixtures/drive/files';

describe('DriveIndexerService - full index', () => {
  test('indexes all indexable files and stores entries in cache', async () => {
    const google = createGoogleMock();
    const cache = createCacheMock();
    const metrics = createMetricsMock();

    // two pages
    google.listDriveFiles
      .mockResolvedValueOnce({ files: [fileDoc, filePdf], nextPageToken: 'P2' })
      .mockResolvedValueOnce({ files: [fileWord] });

    google.extractTextFromFile
      .mockImplementation(async ({ id }: any) => `TEXT_${id}`);

    const bot = createBotMock(google as any, cache as any, { services: { metrics } });
    const indexer = await initIndexer(bot);

    await indexer.reindexAll('root');

    // cache keys
    const keys = await cache.get<string[]>('drive:index:keys');
    expect(keys).toEqual(expect.arrayContaining(['doc1', 'pdf1', 'docx1']));

    // per-file entries
    const doc = await cache.get<any>('drive:index:file:doc1');
    const pdf = await cache.get<any>('drive:index:file:pdf1');
    const docx = await cache.get<any>('drive:index:file:docx1');

    expect(doc?.text).toBe('TEXT_doc1');
    expect(pdf?.text).toBe('TEXT_pdf1');
    expect(docx?.text).toBe('TEXT_docx1');

    // ensure extract called only for indexable mime
    expect(google.extractTextFromFile).toHaveBeenCalledTimes(3);

    // metrics: run counter, duration histogram, total files
    expect(metrics.incCounter).toHaveBeenCalledWith('drive_index_runs_total', { mode: 'full' });
    expect(metrics.observeHistogram).toHaveBeenCalledWith('drive_index_duration_seconds', expect.any(Number), { mode: 'full' });
    expect(metrics.incCounter).toHaveBeenCalledWith('drive_index_files_indexed_total', { mode: 'full', total: 3 });
    // per-file metric
    expect(metrics.incCounter).toHaveBeenCalledWith('drive_index_file_indexed', { mime: fileDoc.mimeType });
    expect(metrics.incCounter).toHaveBeenCalledWith('drive_index_file_indexed', { mime: filePdf.mimeType });
    expect(metrics.incCounter).toHaveBeenCalledWith('drive_index_file_indexed', { mime: fileWord.mimeType });
  });

  test('skips non-indexable mime types', async () => {
    const google = createGoogleMock();
    const cache = createCacheMock();

    const f1 = clone(fileDoc);
    const f2 = { id: 'img1', name: 'Image', mimeType: 'image/png', modifiedTime: '2025-08-13T13:00:00Z' };

    google.listDriveFiles
      .mockResolvedValueOnce({ files: [f1, f2] });

    google.extractTextFromFile.mockImplementation(async ({ id }: any) => `T_${id}`);

    const bot = createBotMock(google as any, cache as any);
    const indexer = await initIndexer(bot);

    await indexer.reindexAll('root');

    const keys = await cache.get<string[]>('drive:index:keys');
    expect(keys).toEqual(['doc1']);
    expect(await cache.get('drive:index:file:doc1')).toBeTruthy();
    expect(await cache.get('drive:index:file:img1')).toBeNull();

    // metrics: skipped non-indexable
    expect(metrics.incCounter).toHaveBeenCalledWith('drive_index_skipped_total', { reason: 'non_indexable_mime', mime: 'image/png' });
  });
});
