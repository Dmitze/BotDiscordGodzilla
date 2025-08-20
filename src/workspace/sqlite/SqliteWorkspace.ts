import Database from 'better-sqlite3';
import { existsSync, mkdirSync, readFileSync } from 'fs';
import { dirname, resolve } from 'path';

export interface RecentItem {
  userId: string;
  fileId: string;
  name?: string | null;
  snippet?: string | null;
  openedAt: number;
}

export interface SavedSearchRow {
  userId: string;
  name: string;
  queryJson: string;
  createdAt: number;
  updatedAt: number;
}

export interface FavoriteRow {
  userId: string;
  fileId: string;
  name?: string | null;
  tags?: string | null;
  addedAt: number;
}

export interface SubscriptionRow {
  userId: string;
  topic: string;
  criteriaJson?: string | null;
  createdAt: number;
}

export interface SqliteWorkspaceOptions {
  dbPath?: string; // default ./data/search-index.db (reuse)
}

export class SqliteWorkspace {
  private db: Database;

  constructor(opts: SqliteWorkspaceOptions = {}) {
    const dbPath = resolve(process.cwd(), opts.dbPath || process.env['BOT_INDEX_DB_PATH'] || './data/search-index.db');
    const dir = dirname(dbPath);
    if (!existsSync(dir)) mkdirSync(dir, { recursive: true });
    this.db = new Database(dbPath);
    this.db.pragma('journal_mode = WAL');
    this.db.pragma('synchronous = NORMAL');

    // apply schema
    const schemaPath = resolve(__dirname, './schema.sql');
    const sql = readFileSync(schemaPath, 'utf8');
    this.db.exec(sql);
  }

  // Favorites
  upsertStar(row: FavoriteRow): void {
    const stmt = this.db.prepare(`INSERT INTO user_stars (user_id, file_id, name, tags, added_at)
      VALUES (@userId, @fileId, @name, @tags, @addedAt)
      ON CONFLICT(user_id, file_id) DO UPDATE SET name=excluded.name, tags=excluded.tags, added_at=excluded.added_at`);
    stmt.run({
      userId: row.userId,
      fileId: row.fileId,
      name: row.name ?? null,
      tags: row.tags ?? null,
      addedAt: row.addedAt,
    });
  }

  removeStar(userId: string, fileId: string): number {
    return this.db.prepare(`DELETE FROM user_stars WHERE user_id = ? AND file_id = ?`).run(userId, fileId).changes;
  }

  listStars(userId: string): FavoriteRow[] {
    return this.db.prepare(`SELECT user_id as userId, file_id as fileId, name, tags, added_at as addedAt FROM user_stars WHERE user_id = ? ORDER BY added_at DESC`).all(userId) as FavoriteRow[];
  }

  // Saved searches
  upsertSavedSearch(row: SavedSearchRow): void {
    const stmt = this.db.prepare(`INSERT INTO saved_searches (user_id, name, query_json, created_at, updated_at)
      VALUES (@userId, @name, @queryJson, @createdAt, @updatedAt)
      ON CONFLICT(user_id, name) DO UPDATE SET query_json=excluded.query_json, updated_at=excluded.updated_at`);
    stmt.run(row);
  }

  getSavedSearch(userId: string, name: string): SavedSearchRow | undefined {
    return this.db.prepare(`SELECT user_id as userId, name, query_json as queryJson, created_at as createdAt, updated_at as updatedAt FROM saved_searches WHERE user_id = ? AND LOWER(name) = LOWER(?)`).get(userId, name) as SavedSearchRow | undefined;
  }

  listSavedSearches(userId: string): SavedSearchRow[] {
    return this.db.prepare(`SELECT user_id as userId, name, query_json as queryJson, created_at as createdAt, updated_at as updatedAt FROM saved_searches WHERE user_id = ? ORDER BY updated_at DESC`).all(userId) as SavedSearchRow[];
  }

  removeSavedSearch(userId: string, name: string): number {
    return this.db.prepare(`DELETE FROM saved_searches WHERE user_id = ? AND LOWER(name) = LOWER(?)`).run(userId, name).changes;
  }

  // Recent items
  upsertRecent(row: RecentItem): void {
    const stmt = this.db.prepare(`INSERT INTO recent_items (user_id, file_id, name, opened_at, snippet)
      VALUES (@userId, @fileId, @name, @openedAt, @snippet)
      ON CONFLICT(user_id, file_id) DO UPDATE SET name=excluded.name, opened_at=excluded.opened_at, snippet=excluded.snippet`);
    stmt.run({
      userId: row.userId,
      fileId: row.fileId,
      name: row.name ?? null,
      openedAt: row.openedAt,
      snippet: row.snippet ?? null,
    });
  }

  listRecent(userId: string, limit = 10): RecentItem[] {
    return this.db.prepare(`SELECT user_id as userId, file_id as fileId, name, snippet, opened_at as openedAt FROM recent_items WHERE user_id = ? ORDER BY opened_at DESC LIMIT ?`).all(userId, Math.max(1, Math.min(100, limit))) as RecentItem[];
  }

  // Subscriptions
  upsertSubscription(row: SubscriptionRow): void {
    const stmt = this.db.prepare(`INSERT INTO subscriptions (user_id, topic, criteria_json, created_at)
      VALUES (@userId, @topic, @criteriaJson, @createdAt)
      ON CONFLICT(user_id, topic) DO UPDATE SET criteria_json=excluded.criteria_json`);
    stmt.run({
      userId: row.userId,
      topic: row.topic,
      criteriaJson: row.criteriaJson ?? null,
      createdAt: row.createdAt,
    });
  }

  removeSubscription(userId: string, topic: string): number {
    return this.db.prepare(`DELETE FROM subscriptions WHERE user_id = ? AND topic = ?`).run(userId, topic).changes;
  }

  listSubscriptions(userId: string): SubscriptionRow[] {
    return this.db.prepare(`SELECT user_id as userId, topic, criteria_json as criteriaJson, created_at as createdAt FROM subscriptions WHERE user_id = ? ORDER BY created_at DESC`).all(userId) as SubscriptionRow[];
  }

  // Stage 7: Change events / Notifications / Digests

  insertChangeEvent(row: { fileId: string; changeId: string; hash: string; occurredAt: number; metaJson?: string | null }): void {
    const stmt = this.db.prepare(`INSERT OR IGNORE INTO change_events (file_id, change_id, hash, occurred_at, meta_json)
      VALUES (@fileId, @changeId, @hash, @occurredAt, @metaJson)`);
    stmt.run({
      fileId: row.fileId,
      changeId: row.changeId,
      hash: row.hash,
      occurredAt: row.occurredAt,
      metaJson: row.metaJson ?? null,
    });
  }

  listSubscriptionsByTopic(topic: string): SubscriptionRow[] {
    return this.db.prepare(`SELECT user_id as userId, topic, criteria_json as criteriaJson, created_at as createdAt FROM subscriptions WHERE topic = ?`).all(topic) as SubscriptionRow[];
  }

  upsertNotification(row: {
    userId: string;
    topic: string;
    fileId: string;
    changeId: string;
    hash: string;
    windowStart: number;
    windowEnd: number;
    status?: 'pending' | 'delivered' | 'failed';
    payloadJson?: string | null;
    createdAt: number;
    updatedAt: number;
  }): void {
    const stmt = this.db.prepare(`INSERT INTO notifications_queue (user_id, topic, file_id, change_id, hash, window_start, window_end, status, payload_json, created_at, updated_at)
      VALUES (@userId, @topic, @fileId, @changeId, @hash, @windowStart, @windowEnd, @status, @payloadJson, @createdAt, @updatedAt)
      ON CONFLICT(user_id, topic, file_id, change_id, hash, window_start)
      DO UPDATE SET payload_json=excluded.payload_json, updated_at=excluded.updated_at`);
    stmt.run({
      userId: row.userId,
      topic: row.topic,
      fileId: row.fileId,
      changeId: row.changeId,
      hash: row.hash,
      windowStart: row.windowStart,
      windowEnd: row.windowEnd,
      status: row.status ?? 'pending',
      payloadJson: row.payloadJson ?? null,
      createdAt: row.createdAt,
      updatedAt: row.updatedAt,
    });
  }

  listPendingNotificationsReady(nowMs: number, limit = 100): Array<{
    id: number; userId: string; topic: string; fileId: string; changeId: string; hash: string;
    windowStart: number; windowEnd: number; status: string; payloadJson?: string | null;
  }> {
    const stmt = this.db.prepare(`SELECT id, user_id as userId, topic, file_id as fileId, change_id as changeId, hash,
      window_start as windowStart, window_end as windowEnd, status, payload_json as payloadJson
      FROM notifications_queue WHERE status = 'pending' AND window_end < ? ORDER BY window_end ASC LIMIT ?`);
    return stmt.all(nowMs, Math.max(1, Math.min(1000, limit))) as any;
  }

  markNotificationsDelivered(ids: number[], deliveredAt: number): number {
    if (!ids.length) return 0;
    const stmt = this.db.prepare(`UPDATE notifications_queue SET status='delivered', delivered_at=?, updated_at=? WHERE id IN (${ids.map(() => '?').join(',')})`);
    const res = stmt.run(deliveredAt, deliveredAt, ...ids);
    return res.changes || 0;
  }

  createDigest(row: { userId: string; period: 'daily' | 'weekly'; windowStart: number; windowEnd: number; payloadJson: string; createdAt: number; deliveredAt?: number | null }): number {
    const stmt = this.db.prepare(`INSERT INTO digests (user_id, period, window_start, window_end, payload_json, created_at, delivered_at)
      VALUES (@userId, @period, @windowStart, @windowEnd, @payloadJson, @createdAt, @deliveredAt)`);
    const res = stmt.run({
      userId: row.userId,
      period: row.period,
      windowStart: row.windowStart,
      windowEnd: row.windowEnd,
      payloadJson: row.payloadJson,
      createdAt: row.createdAt,
      deliveredAt: row.deliveredAt ?? null,
    });
    return Number(res.lastInsertRowid);
  }

  listUserNotificationsInWindow(userId: string, windowStart: number, windowEnd: number): Array<{
    id: number; topic: string; fileId: string; payloadJson?: string | null; windowStart: number; windowEnd: number;
  }> {
    const stmt = this.db.prepare(`SELECT id, topic, file_id as fileId, payload_json as payloadJson, window_start as windowStart, window_end as windowEnd
      FROM notifications_queue WHERE user_id = ? AND window_start >= ? AND window_end <= ?`);
    return stmt.all(userId, windowStart, windowEnd) as any;
  }

  listAllSubscribers(): string[] {
    const rows = this.db.prepare(`SELECT DISTINCT user_id as userId FROM subscriptions`).all() as Array<{ userId: string }>;
    return rows.map(r => r.userId);
  }
}
