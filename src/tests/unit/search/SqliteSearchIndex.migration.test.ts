import { existsSync, unlinkSync } from 'fs';
import { join } from 'path';
import { SqliteSearchIndex } from '@/search/sqlite/SqliteSearchIndex';

function makeDbPath(): string {
  const name = `test-sqlite-mig-${Date.now()}-${Math.floor(Math.random()*1e6)}.db`;
  return join(process.cwd(), 'data', name);
}

describe('SqliteSearchIndex: tokenizer migration uses segment_cache', () => {
  const oldEnv = { ...process.env };
  const dbPath = makeDbPath();

  afterAll(() => {
    process.env = oldEnv;
    try { if (existsSync(dbPath)) unlinkSync(dbPath); } catch {}
  });

  it('rebuilds FTS and still finds documents after tokenizer change', async () => {
    // First run: default tokenizer (porter)
    delete process.env['SEARCH_FTS_TOKENIZER'];
    const idx1 = new SqliteSearchIndex({ dbPath });
    await idx1.upsert({ fileId: 'X', name: 'Doc X', text: 'hello world', mimeType: 'text/plain' });
    let res = await idx1.search({ text: 'hello', limit: 10 });
    expect(res.hits.length).toBeGreaterThan(0);

    // Second run: switch tokenizer to unicode61 -> triggers rebuild
    process.env['SEARCH_FTS_TOKENIZER'] = 'unicode61';
    const idx2 = new SqliteSearchIndex({ dbPath });
    res = await idx2.search({ text: 'hello', limit: 10 });
    expect(res.hits.length).toBeGreaterThan(0);
  });
});
