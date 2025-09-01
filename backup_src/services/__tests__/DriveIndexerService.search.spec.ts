import { createBotMock, createCacheMock, createGoogleMock, initIndexer } from './__utils__/driveIndexerTestHelpers';
import { fileDoc, filePdf, fileWord } from '../../tests/fixtures/drive/files';

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

  test('handles snippet boundaries at start and end', async () => {
    const google = createGoogleMock();
    const cache = createCacheMock();
    const bot = createBotMock(google as any, cache as any);
    const indexer = await initIndexer(bot);

    await cache.set('drive:index:keys', ['start', 'end']);
    await cache.set('drive:index:file:start', {
      id: 'start', name: 'Start', mimeType: fileDoc.mimeType, text: 'Query is at the very beginning of this text', textLength: 41, updatedAt: Date.now(),
    });
    await cache.set('drive:index:file:end', {
      id: 'end', name: 'End', mimeType: filePdf.mimeType, text: 'This text ends with the special Query', textLength: 36, updatedAt: Date.now(),
    });

    const resStart = await indexer.search('Query', 1);
    expect(resStart).toHaveLength(1);
    expect(resStart[0]!.file.snippet.startsWith('…')).toBe(false); // начало — без префикса
    expect(resStart[0]!.file.snippet).toMatch(/Query/);

    const resEnd = await indexer.search('Query', 2);
    const endItem = resEnd.find(r => r.file.id === 'end')!;
    expect(endItem.file.snippet.endsWith('…')).toBe(false); // конец — без суффикса (вхождение у конца)
    expect(endItem.file.snippet).toMatch(/Query$/);
  });

  test('includes different MIME types and keeps snippet trimming', async () => {
    const google = createGoogleMock();
    const cache = createCacheMock();
    const bot = createBotMock(google as any, cache as any);
    const indexer = await initIndexer(bot);

    await cache.set('drive:index:keys', ['d1', 'p1', 'w1']);
    await cache.set('drive:index:file:d1', {
      id: 'd1', name: 'Doc', mimeType: fileDoc.mimeType, text: '... alpha beta gamma world delta ...', textLength: 38, updatedAt: Date.now(),
    });
    await cache.set('drive:index:file:p1', {
      id: 'p1', name: 'PDF', mimeType: filePdf.mimeType, text: 'prefix '.repeat(50) + 'world' + ' suffix'.repeat(50), textLength: 6 * 100 + 5, updatedAt: Date.now(),
    });
    await cache.set('drive:index:file:w1', {
      id: 'w1', name: 'Word', mimeType: fileWord.mimeType, text: 'world in docx', textLength: 13, updatedAt: Date.now(),
    });

    const res = await indexer.search('world', 3);
    const ids = res.map(r => r.file.id);
    expect(ids).toEqual(expect.arrayContaining(['d1', 'p1', 'w1']));
    // длинный текст с триммингом должен иметь многоточия
    const p1 = res.find(r => r.file.id === 'p1')!;
    expect(p1.file.snippet.startsWith('…') || p1.file.snippet.endsWith('…')).toBe(true);
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
