import { createBotMock, createCacheMock, createGoogleMock, initIndexer } from './__utils__/driveIndexerTestHelpers';
import { fileDoc, filePdf } from '../../tests/fixtures/drive/files';

describe('DriveIndexerService - search', () => {
  test('returns snippets and respects limit', async () => {
    const google = createGoogleMock();
    const cache = createCacheMock();

    const bot = createBotMock(google as any, cache as any);
    const indexer = await initIndexer(bot);

    // prefill cache entries (simulate indexed docs)
    await cache.set('drive:index:keys', ['doc1', 'pdf1']);
    await cache.set('drive:index:file:doc1', {
      id: 'doc1', name: fileDoc.name, mimeType: fileDoc.mimeType, text: 'hello world, searching in document', textLength: 33, updatedAt: Date.now(), modifiedTime: fileDoc.modifiedTime,
    });
    await cache.set('drive:index:file:pdf1', {
      id: 'pdf1', name: filePdf.name, mimeType: filePdf.mimeType, text: 'another world line', textLength: 19, updatedAt: Date.now(), modifiedTime: filePdf.modifiedTime,
    });

    const res = await indexer.search('world', 1);
    expect(res).toHaveLength(1);
    expect(res[0]!.file.snippet).toMatch(/world/);
    expect(res[0]!.file.id).toBeDefined();
  });

  test('returns empty array if no index keys', async () => {
    const google = createGoogleMock();
    const cache = createCacheMock();
    const bot = createBotMock(google as any, cache as any);
    const indexer = await initIndexer(bot);

    const res = await indexer.search('q');
    expect(res).toEqual([]);
  });
});
