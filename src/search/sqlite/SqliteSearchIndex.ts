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

    const schemaPath = resolve(__dirname, './schema.sql');
    const schemaSql = readFileSync(schemaPath, 'utf8');
    this.db.exec(schemaSql);

    // Prepared statements
    this.insertDocStmt = this.db.prepare(
      `INSERT INTO documents (fileId, name, mimeType, ownerEmail, sizeBytes, modifiedTime, createdTime, contentHash, textLen, lastIndexedAt, tags, meta)
       VALUES (@fileId, @name, @mimeType, @ownerEmail, @sizeBytes, @modifiedTime, @createdTime, @contentHash, @textLen, @lastIndexedAt, @tags, @meta)`
    );
    this.updateDocStmt = this.db.prepare(
      `UPDATE documents SET name=@name, mimeType=@mimeType, ownerEmail=@ownerEmail, sizeBytes=@sizeBytes,
         modifiedTime=@modifiedTime, createdTime=@createdTime, contentHash=@contentHash, textLen=@textLen,
         lastIndexedAt=@lastIndexedAt, tags=@tags, meta=@meta WHERE fileId=@fileId`
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
  }): Promise<void> {
    const textNorm = doc.text; // предполагаем, что текст уже нормализован выше
    const contentHash = sha256(textNorm);
    const lastIndexedAt = nowMs();
    const textLen = textNorm.length;

    const existing = this.selectDocStmt.get(doc.fileId) as any | undefined;

    const run = this.db.transaction(() => {
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
        });
        if (hashChanged) {
          // refresh FTS content only if text changed
          this.deleteFtsByIdStmt.run(doc.fileId);
          this.insertFtsStmt.run(doc.fileId, doc.name, textNorm);
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

    const rows = this.db.prepare(baseSql).all(matchExpr, ...params, limit, offset) as any[];

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
      const stmt = this.db.prepare(`SELECT COUNT(1) as c FROM document_versions WHERE fileId = ?`);
      const filtered: SearchHit[] = [];
      for (const h of hits) {
        const { c } = stmt.get(h.fileId) as { c: number };
        if (c > 1) filtered.push(h);
      }
      return { hits: filtered, total: cnt };
    }

    return { hits, total: cnt };
  }

  async getDiff(fileId: string): Promise<{ latestHash: string; previousHash?: string; diffMeta?: Record<string, { old?: unknown; new?: unknown }> }> {
    const versions = this.selectVersionsStmt.all(fileId) as any[];
    const latest = versions[0];
    const prev = versions[1];
    const res: { latestHash: string; previousHash?: string; diffMeta?: Record<string, { old?: unknown; new?: unknown }> } = {
      latestHash: latest?.contentHash,
      previousHash: prev?.contentHash,
    };
    // naive meta diff
    const latestMeta = parseJson(latest?.meta) as Record<string, unknown> | undefined;
    const prevMeta = parseJson(prev?.meta) as Record<string, unknown> | undefined;
    if (latestMeta && prevMeta) {
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
    }
    return res;
  }
}
