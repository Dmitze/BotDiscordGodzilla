import { createBotMock, createCacheMock, createGoogleMock, createMetricsMock, initIndexer } from './__utils__/driveIndexerTestHelpers';
import { fileDoc, filePdf, nonIndexable } from '../../tests/fixtures/drive/files';

describe('DriveIndexerService - metrics', () => {
  test('reindexAll updates metrics for full run', async () => {
    const google = createGoogleMock();
    const cache = createCacheMock();
    const metrics = createMetricsMock();

    google.listDriveFiles
      .mockResolvedValueOnce({ files: [fileDoc], nextPageToken: 'T2' })
      .mockResolvedValueOnce({ files: [filePdf] });
    google.extractTextFromFile.mockImplementation(async ({ id }: any) => `TEXT_${id}`);

    const bot = createBotMock(google as any, cache as any, { services: { metrics } });
    const indexer = await initIndexer(bot);

    await indexer.reindexAll('root');

    // runs counter
    expect(metrics.incCounter).toHaveBeenCalledWith('drive_index_runs_total', { mode: 'full' });
    // files indexed total with label total
    expect(metrics.incCounter).toHaveBeenCalledWith('drive_index_files_indexed_total', expect.objectContaining({ mode: 'full', total: 2 }));
    // duration histogram observed
    expect(metrics.observeHistogram).toHaveBeenCalledWith('drive_index_duration_seconds', expect.any(Number), { mode: 'full' });
    // per-file metric
    expect(metrics.incCounter).toHaveBeenCalledWith('drive_index_file_indexed', { mime: fileDoc.mimeType });
    expect(metrics.incCounter).toHaveBeenCalledWith('drive_index_file_indexed', { mime: filePdf.mimeType });
  });

  test('reindexIncremental updates metrics and respects needReindex', async () => {
    const google = createGoogleMock();
    const cache = createCacheMock();
    const metrics = createMetricsMock();

    // same files on single page
    google.listDriveFiles.mockResolvedValue({ files: [fileDoc, filePdf] });
    google.extractTextFromFile.mockImplementation(async ({ id }: any) => `TEXT_${id}`);

    // simulate cache has one up-to-date entry (skip reindex)
    await cache.set('drive:index:file:doc1', {
      id: 'doc1', name: fileDoc.name, mimeType: fileDoc.mimeType, text: 'old', textLength: 3, updatedAt: Date.now(), modifiedTime: fileDoc.modifiedTime,
    });
    await cache.set('drive:index:keys', ['doc1']);

    const bot = createBotMock(google as any, cache as any, { services: { metrics } });
    const indexer = await initIndexer(bot);

    await indexer.reindexIncremental('root');

    expect(metrics.incCounter).toHaveBeenCalledWith('drive_index_runs_total', { mode: 'incremental' });
    // only one file should be reindexed (pdf)
    expect(metrics.incCounter).toHaveBeenCalledWith('drive_index_files_indexed_total', expect.objectContaining({ mode: 'incremental', total: 1 }));
    expect(metrics.observeHistogram).toHaveBeenCalledWith('drive_index_duration_seconds', expect.any(Number), { mode: 'incremental' });
    expect(metrics.incCounter).toHaveBeenCalledWith('drive_index_file_indexed', { mime: filePdf.mimeType });
  });

  test('indexOneFileByMeta skips non-indexable mime with metric', async () => {
    const google = createGoogleMock();
    const cache = createCacheMock();
    const metrics = createMetricsMock();

    const bot = createBotMock(google as any, cache as any, { services: { metrics } });
    const indexer = await initIndexer(bot);

    // action
    await indexer.indexOneFileByMeta(nonIndexable);

    expect(metrics.incCounter).toHaveBeenCalledWith('drive_index_skipped_total', { reason: 'non_indexable_mime', mime: nonIndexable.mimeType });
    expect(google.extractTextFromFile).not.toHaveBeenCalled();
  });
});
