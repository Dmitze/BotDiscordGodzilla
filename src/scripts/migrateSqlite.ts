import { resolve } from 'path';
import { readFileSync } from 'fs';
import { SqliteSearchIndex } from '@/search/sqlite/SqliteSearchIndex';
import { SqliteWorkspace } from '@/workspace/sqlite/SqliteWorkspace';

async function main() {
  process.env['DB_SCHEMA_INIT_MODE'] = 'run';
  // initialize and force schema/migrations
  // Search index
  const idx = new SqliteSearchIndex({});
  // Workspace
  const ws = new SqliteWorkspace({});
  console.log('SQLite schema ensured for SearchIndex and Workspace');
}

main().catch(err => {
  console.error('Migration failed', err);
  process.exit(1);
});


