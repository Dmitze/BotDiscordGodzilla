import { existsSync, unlinkSync } from 'fs';
import { join } from 'path';
import { SqliteSearchIndex } from '@/search/sqlite/SqliteSearchIndex';

// Simple helper to make a temporary DB path under project data/
function makeDbPath(): string {
  const name = `test-sqlite-fileid-${Date.now()}-${Math.floor(Math.random()*1e6)}.db`;
  return join(process.cwd(), 'data', name);
}

describe('SqliteSearchIndex: filters.fileId', () => {
  const dbPath = makeDbPath();
  const idx = new SqliteSearchIndex({ dbPath });

  afterAll(() => {
    // best-effort cleanup
    try { if (existsSync(dbPath)) unlinkSync(dbPath); } catch {}
  });

  it('returns only docs matching provided fileId[]', async () => {
    await idx.upsert({
      fileId: 'A',
      name: 'Doc A',
      text: 'alpha beta gamma',
      mimeType: 'text/plain',
    });
    await idx.upsert({
      fileId: 'B',
      name: 'Doc B',
      text: 'alpha delta epsilon',
      mimeType: 'text/plain',
    });

    // Without filter should be able to find both by a common term
    const all = await idx.search({ text: 'alpha', limit: 10 });
    expect(all.total).toBeGreaterThanOrEqual(2);

    // With filter: only fileId A
    const onlyA = await idx.search({ text: 'alpha', limit: 10, filters: { fileId: ['A'] } });
    expect(onlyA.hits.length).toBeGreaterThanOrEqual(1);
    for (const h of onlyA.hits) expect(h.fileId).toBe('A');

    // With filter: only fileId B
    const onlyB = await idx.search({ text: 'alpha', limit: 10, filters: { fileId: ['B'] } });
    expect(onlyB.hits.length).toBeGreaterThanOrEqual(1);
    for (const h of onlyB.hits) expect(h.fileId).toBe('B');

    // With filter unrelated: none
    const none = await idx.search({ text: 'alpha', limit: 10, filters: { fileId: ['NON_EXISTENT'] } });
    expect(none.hits.length).toBe(0);
    expect(none.total).toBe(0);
  });
});
