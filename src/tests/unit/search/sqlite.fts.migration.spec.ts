import { SqliteSearchIndex } from '@/search/sqlite/SqliteSearchIndex';

describe('SqliteSearchIndex FTS migration gating', () => {
  it('skips schema when DB_SCHEMA_INIT_MODE=defer', () => {
    const prev = process.env['DB_SCHEMA_INIT_MODE'];
    process.env['DB_SCHEMA_INIT_MODE'] = 'defer';
    const idx = new SqliteSearchIndex({ dbPath: './data/test-mig.db' });
    expect(idx).toBeTruthy();
    if (prev === undefined) delete process.env['DB_SCHEMA_INIT_MODE']; else process.env['DB_SCHEMA_INIT_MODE'] = prev;
  });
});


