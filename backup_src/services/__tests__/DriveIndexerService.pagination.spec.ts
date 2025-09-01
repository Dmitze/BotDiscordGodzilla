import { createBotMock, createCacheMock, createGoogleMock, initIndexer } from './__utils__/driveIndexerTestHelpers';
import { fileDoc, filePdf, fileWord } from '../../tests/fixtures/drive/files';

describe('DriveIndexerService - pagination handling', () => {
  test('iterates pages until nextPageToken is undefined', async () => {
    const google = createGoogleMock();
    const cache = createCacheMock();

    google.listDriveFiles
      .mockResolvedValueOnce({ files: [fileDoc], nextPageToken: 'T2' })
      .mockResolvedValueOnce({ files: [filePdf], nextPageToken: 'T3' })
      .mockResolvedValueOnce({ files: [fileWord] });

    google.extractTextFromFile.mockImplementation(async ({ id }: any) => `TEXT_${id}`);

    const bot = createBotMock(google as any, cache as any);
    const indexer = await initIndexer(bot);

    await indexer.reindexAll('root');

    // should process 3 files
    expect(google.extractTextFromFile).toHaveBeenCalledTimes(3);

    const keys = await cache.get<string[]>('drive:index:keys');
    expect(keys?.length).toBe(3);
    expect(keys).toEqual(expect.arrayContaining(['doc1', 'pdf1', 'docx1']));
  });
});
