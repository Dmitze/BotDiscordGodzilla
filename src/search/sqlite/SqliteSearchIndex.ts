import Database from 'better-sqlite3';
import type { Statement } from 'better-sqlite3';
import { mkdirSync, readFileSync, existsSync } from 'fs';
import { dirname, resolve } from 'path';
import { sha256 } from '../../utils/hash';
import type { SearchHit, SearchIndex, SearchQuery } from '../SearchIndex';

function nowMs(): number {
  return Date.now();
}

function asJson(value: unknown): string | null {
  if (value === undefined || value === null) return null;
  try { return JSON.stringify(value); } catch {
    return null;
  }
}

function parseJson<T = unknown>(value: unknown): T | undefined {
  if (typeof value !== 'string' || !value) return undefined;
  try { return JSON.parse(value) as T; } catch { return undefined; }
}

export interface SqliteIndexOptions {
  dbPath?: string; // default ./data/search-index.db
  schemaPath?: string; // default to co-located schema.sql
}

export class SqliteSearchIndex implements SearchIndex {
  private db: Database;

  private insertDocStmt: Statement;
  private updateDocStmt: Statement;
  private selectDocStmt: Statement;
  private insertFtsStmt: Statement;
  private deleteFtsByIdStmt: Statement;
  private insertVersionStmt: Statement;
  private selectVersionsStmt: Statement;
  private insertSegmentStmt: Statement;
  private selectAllDocsStmt: Statement;
  private selectSegmentByHashStmt: Statement;

  constructor(opts: SqliteIndexOptions = {}) {
    const dbPath = resolve(
      process.cwd(),
      opts.dbPath ||
        process.env['SEARCH_INDEX_PATH'] ||
        process.env['BOT_INDEX_DB_PATH'] ||
        './data/search-index.db'
    );
    const dir = dirname(dbPath);
    if (!existsSync(dir)) mkdirSync(dir, { recursive: true });

    this.db = new Database(dbPath);
    this.db.pragma('journal_mode = WAL');
    this.db.pragma('synchronous = NORMAL');

    const initMode = (process.env['DB_SCHEMA_INIT_MODE'] || 'run').toLowerCase();
    if (initMode !== 'defer') {
      const schemaPath = resolve(__dirname, './schema.sql');
      const schemaSql = readFileSync(schemaPath, 'utf8');
      this.db.exec(schemaSql);
      // Ensure meta table for migrations
      this.db.exec(`CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT)`);
      // Ensure required columns exist for metadata persistence
      this.ensureDocumentColumns();
    }

    // Prepared statements (guard for defer mode)
    if (initMode === 'defer') {
      // Avoid preparing statements against non-existent tables in defer mode.
      // Assign a harmless no-op statement so constructor remains safe.
      const noop = this.db.prepare(`SELECT 1`);
      this.insertDocStmt = noop as unknown as Statement;
      this.updateDocStmt = noop as unknown as Statement;
      this.selectDocStmt = noop as unknown as Statement;
      this.insertFtsStmt = noop as unknown as Statement;
      this.deleteFtsByIdStmt = noop as unknown as Statement;
      this.insertVersionStmt = noop as unknown as Statement;
      this.selectVersionsStmt = noop as unknown as Statement;
      this.insertSegmentStmt = noop as unknown as Statement;
      this.selectAllDocsStmt = noop as unknown as Statement;
      this.selectSegmentByHashStmt = noop as unknown as Statement;
    } else {
      // Normal mode: prepare real statements
      this.insertDocStmt = this.db.prepare(
      `INSERT INTO documents (fileId, name, mimeType, ownerEmail, sizeBytes, modifiedTime, createdTime, contentHash, textLen, lastIndexedAt, tags, meta, language, labels, path)
       VALUES (@fileId, @name, @mimeType, @ownerEmail, @sizeBytes, @modifiedTime, @createdTime, @contentHash, @textLen, @lastIndexedAt, @tags, @meta, @language, @labels, @path)`
      );
      this.updateDocStmt = this.db.prepare(
      `UPDATE documents SET name=@name, mimeType=@mimeType, ownerEmail=@ownerEmail, sizeBytes=@sizeBytes,
         modifiedTime=@modifiedTime, createdTime=@createdTime, contentHash=@contentHash, textLen=@textLen,
         lastIndexedAt=@lastIndexedAt, tags=@tags, meta=@meta, language=@language, labels=@labels, path=@path WHERE fileId=@fileId`
      );
      this.selectDocStmt = this.db.prepare(
      `SELECT * FROM documents WHERE fileId = ?`
      );

      this.insertFtsStmt = this.db.prepare(
      `INSERT INTO documents_fts (fileId, name, content) VALUES (?, ?, ?)`
      );
      this.deleteFtsByIdStmt = this.db.prepare(
      `DELETE FROM documents_fts WHERE fileId = ?`
      );

      this.insertVersionStmt = this.db.prepare(
      `INSERT INTO document_versions (fileId, version, contentHash, textLen, modifiedTime, createdAt, meta)
       VALUES (@fileId, @version, @contentHash, @textLen, @modifiedTime, @createdAt, @meta)`
      );
      this.selectVersionsStmt = this.db.prepare(
      `SELECT * FROM document_versions WHERE fileId = ? ORDER BY version DESC LIMIT 2`
      );

      // segment cache helpers
      this.insertSegmentStmt = this.db.prepare(
      `INSERT OR REPLACE INTO segment_cache (contentHash, text, normText, updatedAt) VALUES (?, ?, ?, ?)`
      );
      this.selectAllDocsStmt = this.db.prepare(
      `SELECT fileId, name, contentHash FROM documents`
      );
      this.selectSegmentByHashStmt = this.db.prepare(
      `SELECT normText FROM segment_cache WHERE contentHash = ?`
      );

      // Perform FTS tokenizer migration only after statements are ready
      this.maybeMigrateFtsTokenizer();
    }

    if (initMode === 'defer') {
      // Skip heavy migrations now; they can be run by a separate script before app start
    }
  }

  private ensureDocumentColumns(): void {
    const cols = this.db.prepare(`PRAGMA table_info(documents)`).all() as Array<{ name: string }>;
    const names = new Set(cols.map(c => c.name));
    const toAdd: string[] = [];
    if (!names.has('language')) toAdd.push(`ALTER TABLE documents ADD COLUMN language TEXT`);
    if (!names.has('labels')) toAdd.push(`ALTER TABLE documents ADD COLUMN labels TEXT`);
    if (!names.has('path')) toAdd.push(`ALTER TABLE documents ADD COLUMN path TEXT`);
    this.db.transaction(() => {
      for (const sql of toAdd) this.db.exec(sql);
      // indexes are idempotent
      this.db.exec(`CREATE INDEX IF NOT EXISTS idx_documents_language ON documents(language)`);
      this.db.exec(`CREATE INDEX IF NOT EXISTS idx_documents_path ON documents(path)`);
    })();
  }

  private getDesiredTokenizer(): 'porter' | 'unicode61' {
    const envTok = (process.env['SEARCH_FTS_TOKENIZER'] || '').toLowerCase();
    return envTok === 'unicode61' ? 'unicode61' : 'porter';
  }

  private maybeMigrateFtsTokenizer(): void {
    const desired = this.getDesiredTokenizer();
    const row = this.db.prepare(`SELECT value FROM meta WHERE key = 'fts_tokenizer'`).get() as { value?: string } | undefined;
    const current = (row?.value || '').toLowerCase();
    if (!current) {
      // first-run: record current tokenizer based on schema (porter by default) or desired if different
      const initial = desired;
      this.db.prepare(`INSERT OR REPLACE INTO meta (key, value) VALUES ('fts_tokenizer', ?)`).run(initial);
      if (initial !== 'porter') {
        // Recreate FTS with desired tokenizer on first run if not porter
        this.rebuildFts(desired);
      }
      return;
    }
    if (current !== desired) {
      this.rebuildFts(desired);
      this.db.prepare(`INSERT OR REPLACE INTO meta (key, value) VALUES ('fts_tokenizer', ?)`).run(desired);
    }
  }

  private rebuildFts(tokenizer: 'porter' | 'unicode61') {
    const createSql = `CREATE VIRTUAL TABLE IF NOT EXISTS documents_fts USING fts5(
      fileId UNINDEXED,
      name,
      content,
      tokenize='${tokenizer}'
    )`;
    const recreate = this.db.transaction(() => {
      this.db.exec(`DROP TABLE IF EXISTS documents_fts`);
      this.db.exec(createSql);
      const docs = this.selectAllDocsStmt.all() as { fileId: string; name: string; contentHash: string }[];
      for (const d of docs) {
        const seg = this.selectSegmentByHashStmt.get(d.contentHash) as { normText?: string } | undefined;
        const text = seg?.normText ?? '';
        if (text && text.length) {
          this.insertFtsStmt.run(d.fileId, d.name, text);
        }
      }
    });
    recreate();
  }

  async upsert(doc: {
    fileId: string;
    name: string;
    mimeType?: string;
    ownerEmail?: string;
    sizeBytes?: number;
    modifiedTime?: number;
    createdTime?: number;
    text: string;
    tags?: string[];
    meta?: unknown;
    language?: string;
    labels?: string[];
    path?: string;
  }): Promise<void> {
    const textNorm = doc.text; // нормализация вище по пайплайну (MVP)
    const contentHash = sha256(textNorm);
    const lastIndexedAt = nowMs();
    const textLen = textNorm.length;

    const existing = this.selectDocStmt.get(doc.fileId);

    const run = this.db.transaction(() => {
      // persist segment cache for potential FTS rebuilds
      this.insertSegmentStmt.run(contentHash, doc.text, textNorm, lastIndexedAt);
      if (!existing) {
        this.insertDocStmt.run({
          fileId: doc.fileId,
          name: doc.name,
          mimeType: doc.mimeType ?? null,
          ownerEmail: doc.ownerEmail ?? null,
          sizeBytes: doc.sizeBytes ?? null,
          modifiedTime: doc.modifiedTime ?? null,
          createdTime: doc.createdTime ?? null,
          contentHash,
          textLen,
          lastIndexedAt,
          tags: doc.tags ? JSON.stringify(doc.tags) : null,
          meta: asJson(doc.meta),
          language: doc.language ?? null,
          labels: doc.labels ? JSON.stringify(doc.labels) : null,
          path: doc.path ?? null,
        });
        // FTS insert
        this.deleteFtsByIdStmt.run(doc.fileId);
        this.insertFtsStmt.run(doc.fileId, doc.name, textNorm);
        // version 1
        this.insertVersionStmt.run({
          fileId: doc.fileId,
          version: 1,
          contentHash,
          textLen,
          modifiedTime: doc.modifiedTime ?? null,
          createdAt: lastIndexedAt,
          meta: asJson(doc.meta),
        });
      } else {
        const hashChanged = existing.contentHash !== contentHash;
        // update main row
        this.updateDocStmt.run({
          fileId: doc.fileId,
          name: doc.name,
          mimeType: doc.mimeType ?? null,
          ownerEmail: doc.ownerEmail ?? null,
          sizeBytes: doc.sizeBytes ?? null,
          modifiedTime: doc.modifiedTime ?? null,
          createdTime: doc.createdTime ?? null,
          contentHash,
          textLen,
          lastIndexedAt,
          tags: doc.tags ? JSON.stringify(doc.tags) : null,
          meta: asJson(doc.meta),
          language: doc.language ?? null,
          labels: doc.labels ? JSON.stringify(doc.labels) : null,
          path: doc.path ?? null,
        });
        if (hashChanged) {
          // refresh FTS content only if text changed
          this.deleteFtsByIdStmt.run(doc.fileId);
          this.insertFtsStmt.run(doc.fileId, doc.name, textNorm);
          // update segment cache timestamp and norm
          this.insertSegmentStmt.run(contentHash, doc.text, textNorm, lastIndexedAt);
          // new version = prev.version + 1
          const latest = this.db.prepare(`SELECT MAX(version) as v FROM document_versions WHERE fileId = ?`).get(doc.fileId) as { v?: number };
          const nextV = (latest?.v ?? 0) + 1;
          this.insertVersionStmt.run({
            fileId: doc.fileId,
            version: nextV,
            contentHash,
            textLen,
            modifiedTime: doc.modifiedTime ?? null,
            createdAt: lastIndexedAt,
            meta: asJson(doc.meta),
          });
        }
      }
    });

    run();
  }

  async search(q: SearchQuery): Promise<{ hits: SearchHit[]; total: number }> {
    const limit = Math.max(1, Math.min(200, q.limit ?? 20));
    const offset = Math.max(0, q.offset ?? 0);

    // FTS MATCH
    const text = (q.text || '').trim();
    const matchExpr = text.length ? text : '*';

    // Build filters
    const where: string[] = [];
    const params: any[] = [];

    if (q.filters?.fileId?.length) {
      where.push(`d.fileId IN (${q.filters.fileId.map(() => '?').join(',')})`);
      params.push(...q.filters.fileId);
    }
    if (q.filters?.mime?.length) {
      where.push(`d.mimeType IN (${q.filters.mime.map(() => '?').join(',')})`);
      params.push(...q.filters.mime);
    }
    if (q.filters?.owner?.length) {
      where.push(`d.ownerEmail IN (${q.filters.owner.map(() => '?').join(',')})`);
      params.push(...q.filters.owner);
    }
    if (q.filters?.modifiedFrom) { where.push(`d.modifiedTime >= ?`); params.push(q.filters.modifiedFrom); }
    if (q.filters?.modifiedTo) { where.push(`d.modifiedTime <= ?`); params.push(q.filters.modifiedTo); }
    if (q.filters?.sizeFrom) { where.push(`d.sizeBytes >= ?`); params.push(q.filters.sizeFrom); }
    if (q.filters?.sizeTo) { where.push(`d.sizeBytes <= ?`); params.push(q.filters.sizeTo); }
    if (q.filters?.tags?.length) {
      // simple contains check in JSON string (for MVP)
      for (const t of q.filters.tags) {
        where.push(`d.tags LIKE ?`);
        params.push(`%"${t}"%`);
      }
    }
    if (q.filters?.language?.length) {
      where.push(`d.language IN (${q.filters.language.map(() => '?').join(',')})`);
      params.push(...q.filters.language);
    }
    if (q.filters?.pathPrefix && q.filters.pathPrefix.length) {
      // normalize to leading slash; store paths like "/a/b"
      const pref = q.filters.pathPrefix.startsWith('/') ? q.filters.pathPrefix : `/${q.filters.pathPrefix}`;
      where.push(`d.path LIKE ?`);
      params.push(`${pref}%`);
    }
    if (q.filters?.labels?.length) {
      for (const l of q.filters.labels) {
        where.push(`d.labels LIKE ?`);
        params.push(`%"${l}"%`);
      }
    }

    const whereSql = where.length ? `AND ${where.join(' AND ')}` : '';

    const baseSql = `SELECT d.fileId, d.name, d.mimeType, d.ownerEmail, d.modifiedTime,
      d.contentHash, d.textLen,
      snippet(documents_fts, 2, '[', ']', '...', 10) as snippet,
      bm25(documents_fts) as score
      FROM documents_fts f
      JOIN documents d ON d.fileId = f.fileId
      WHERE documents_fts MATCH ? ${whereSql}
      ORDER BY score
      LIMIT ? OFFSET ?`;

    const rows = this.db.prepare(baseSql).all(matchExpr, ...params, limit, offset);

    const countSql = `SELECT COUNT(1) as cnt FROM documents_fts f JOIN documents d ON d.fileId=f.fileId
      WHERE documents_fts MATCH ? ${whereSql}`;
    const { cnt } = this.db.prepare(countSql).get(matchExpr, ...params) as { cnt: number };

    const hits: SearchHit[] = rows.map(r => ({
      fileId: r.fileId,
      name: r.name,
      mimeType: r.mimeType ?? undefined,
      ownerEmail: r.ownerEmail ?? undefined,
      modifiedTime: r.modifiedTime ?? undefined,
      snippet: r.snippet ?? undefined,
      contentHash: r.contentHash,
      textLen: r.textLen,
      score: typeof r.score === 'number' ? r.score : undefined,
    }));

    // changesOnly: отфильтровать те, у кого есть более одной версии
    if (q.changesOnly) {
      if (!hits.length) return { hits, total: cnt };
      const ids = hits.map(h => h.fileId);
      // batch одним запросом через IN; защищаемся от слишком длинного IN (ограничим до 999 — лимит SQLite по умолчанию)
      const chunk = (arr: string[], n: number) => {
        const res: string[][] = []; for (let i = 0; i < arr.length; i += n) res.push(arr.slice(i, i+n)); return res;
      };
      const chunks = chunk(ids, 900);
      const changed = new Set<string>();
      const base = `SELECT fileId FROM document_versions WHERE fileId IN ($PLACEHOLDERS) GROUP BY fileId HAVING MAX(version) > 1`;
      for (const part of chunks) {
        const sql = base.replace('$PLACEHOLDERS', part.map(() => '?').join(','));
        const rows = this.db.prepare(sql).all(...part) as Array<{ fileId: string }>;
        for (const r of rows) changed.add(r.fileId);
      }
      const filtered: SearchHit[] = hits.filter(h => changed.has(h.fileId));
      return { hits: filtered, total: filtered.length };
    }

    return { hits, total: cnt };
  }

  async getDiff(fileId: string): Promise<{ latestHash: string; previousHash?: string; diffMeta?: Record<string, { old?: unknown; new?: unknown }> }> {
    const versions = this.selectVersionsStmt.all(fileId);
    const latest = versions[0];
    const prev = versions[1];
    const res: { latestHash: string; previousHash?: string; diffMeta?: Record<string, { old?: unknown; new?: unknown }> } = {
      latestHash: latest?.contentHash,
      previousHash: prev?.contentHash,
    };
    // naive meta diff
    const latestMeta = (parseJson<Record<string, unknown>>(latest?.meta)) ?? {};
    const prevMeta = (parseJson<Record<string, unknown>>(prev?.meta)) ?? {};
    const changed: Record<string, { old?: unknown; new?: unknown }> = {};
    const keys = new Set([...Object.keys(latestMeta), ...Object.keys(prevMeta)]);
    for (const k of keys) {
      const a = latestMeta[k];
      const b = prevMeta[k];
      if (JSON.stringify(a) !== JSON.stringify(b)) {
        changed[k] = { old: b, new: a };
      }
    }
    if (Object.keys(changed).length) res.diffMeta = changed;
    return res;
  }
}
