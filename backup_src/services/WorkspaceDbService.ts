import type { BotConfig } from '@/types';
import type { DriveListQuery } from '@/types/drive';
import type { GoogleService } from './GoogleService';
import { SqliteWorkspace } from '@/workspace/sqlite/SqliteWorkspace';

function nowMs(): number { return Date.now(); }

export type Favorite = { fileId: string; name?: string; tags?: string[]; addedAt: number };
export type SavedSearch = { name: string; filters: DriveListQuery; createdAt: number; updatedAt: number };
export type RecentItem = { fileId: string; name?: string; snippet?: string; openedAt: number };
export type Subscription = { topic: string; criteria?: unknown; createdAt: number };

export class WorkspaceDbService {
  private db: SqliteWorkspace;

  constructor(config: BotConfig) {
    // touch config to avoid TS6133 unused-parameter error while keeping signature stable
    void config;
    const dbPath = process.env['BOT_INDEX_DB_PATH'] || './data/search-index.db';
    this.db = new SqliteWorkspace({ dbPath });
  }

  async initialize(): Promise<void> { /* no-op */ }
  async shutdown(): Promise<void> { /* no-op */ }
  getStats(): Record<string, unknown> { return {}; }

  // Favorites
  async addFavorite(userId: string, fileId: string, name?: string, tags?: string[]): Promise<{ added: boolean; favorite: Favorite }>{
    const addedAt = nowMs();
    this.db.upsertStar({ userId, fileId, name: name ?? null, tags: Array.isArray(tags) ? JSON.stringify(tags) : null, addedAt });
    const favorite: Favorite = { fileId, addedAt, ...(name !== undefined ? { name } : {}), ...(tags !== undefined ? { tags } : {}) };
    return { added: true, favorite };
  }
  async removeFavorite(userId: string, fileId: string): Promise<boolean> { return this.db.removeStar(userId, fileId) > 0; }
  async listFavorites(userId: string): Promise<Favorite[]> {
    return this.db.listStars(userId).map(r => {
      const fav: Favorite = { fileId: r.fileId, addedAt: r.addedAt };
      if (r.name != null) (fav as any).name = r.name;
      if (r.tags) (fav as any).tags = JSON.parse(r.tags);
      return fav;
    });
  }

  // Saved searches
  async saveSearch(userId: string, name: string, filters: DriveListQuery): Promise<{ created: boolean; search: SavedSearch }>{
    const now = nowMs();
    const existing = this.db.getSavedSearch(userId, name);
    this.db.upsertSavedSearch({ userId, name, queryJson: JSON.stringify(filters), createdAt: existing?.createdAt ?? now, updatedAt: now });
    return { created: !existing, search: { name, filters, createdAt: existing?.createdAt ?? now, updatedAt: now } };
  }
  async removeSearch(userId: string, name: string): Promise<boolean> { return this.db.removeSavedSearch(userId, name) > 0; }
  async listSearches(userId: string): Promise<SavedSearch[]> {
    return this.db.listSavedSearches(userId).map(r => ({ name: r.name, filters: JSON.parse(r.queryJson), createdAt: r.createdAt, updatedAt: r.updatedAt }));
  }
  getSavedSearch(userId: string, name: string): SavedSearch | undefined {
    const r = this.db.getSavedSearch(userId, name);
    return r ? { name: r.name, filters: JSON.parse(r.queryJson), createdAt: r.createdAt, updatedAt: r.updatedAt } : undefined;
  }

  // Execute saved search through Google when SQLite is not used by caller
  async runSearch(userId: string, name: string, deps: { google: GoogleService; config: BotConfig }): Promise<any | undefined> {
    const saved = this.getSavedSearch(userId, name);
    if (!saved) return undefined;
    const base = { ...saved.filters } as Partial<DriveListQuery>;
    const cfg = deps.config.drive || {};
    if (Array.isArray(cfg.allowedMime) && cfg.allowedMime.length) {
      base.mimeIncludes = base.mimeIncludes && base.mimeIncludes.length
        ? base.mimeIncludes.filter(m => cfg.allowedMime.includes(m))
        : cfg.allowedMime;
    }
    if (Array.isArray(cfg.ownerAllowlist) && cfg.ownerAllowlist.length) {
      base.ownerAllowlist = cfg.ownerAllowlist;
    }
    if (!base.folderId) {
      base.folderId = deps.config.google?.driveFolderId ?? deps.config.drive?.folderId ?? undefined;
    }
    if (!base.folderId) return undefined;
    return deps.google.listDriveFiles(base as DriveListQuery);
  }

  // Recent
  async addRecent(userId: string, item: RecentItem): Promise<void> {
    this.db.upsertRecent({ userId, fileId: item.fileId, name: item.name ?? null, snippet: item.snippet ?? null, openedAt: item.openedAt });
  }
  async listRecent(userId: string, limit = 10): Promise<RecentItem[]> {
    return this.db.listRecent(userId, limit).map(r => {
      const item: RecentItem = { fileId: r.fileId, openedAt: r.openedAt } as any;
      if (r.name != null) (item as any).name = r.name;
      if (r.snippet != null) (item as any).snippet = r.snippet;
      return item;
    });
  }

  // Subscriptions
  async subscribe(userId: string, topic: string, criteria?: unknown): Promise<Subscription> {
    const createdAt = nowMs();
    this.db.upsertSubscription({ userId, topic, criteriaJson: criteria ? JSON.stringify(criteria) : null, createdAt });
    return { topic, criteria, createdAt };
  }
  async unsubscribe(userId: string, topic: string): Promise<boolean> { return this.db.removeSubscription(userId, topic) > 0; }
  async listSubscriptions(userId: string): Promise<Subscription[]> {
    return this.db.listSubscriptions(userId).map(r => ({ topic: r.topic, criteria: r.criteriaJson ? JSON.parse(r.criteriaJson) : undefined, createdAt: r.createdAt }));
  }

  // Stage 7: notifications and digests
  /**
   * Дедуп + коалесинг: сохраняет событие и ставит уведомления в очередь для всех подписчиков topic=file:<fileId>
   */
  async notifyChange(evt: { fileId: string; changeId: string; hash: string; occurredAt?: number; meta?: unknown }): Promise<void> {
    const occurredAt = evt.occurredAt ?? nowMs();
    // 1) дедуп события
    this.db.insertChangeEvent({ fileId: evt.fileId, changeId: evt.changeId, hash: evt.hash, occurredAt, metaJson: evt.meta ? JSON.stringify(evt.meta) : null });

    // 2) коалесинг по окну
    const windowMs = Number(process.env['WORKSPACE_NOTIF_WINDOW_MS'] || 600000); // 10 мин по умолчанию
    const windowStart = Math.floor(occurredAt / windowMs) * windowMs;
    const windowEnd = windowStart + windowMs - 1;
    const createdAt = nowMs();

    const topic = `file:${evt.fileId}`;
    const subs = this.db.listSubscriptionsByTopic(topic);
    const payload = { fileId: evt.fileId, changes: [{ changeId: evt.changeId, hash: evt.hash, occurredAt }] };
    const payloadJson = JSON.stringify(payload);
    for (const s of subs) {
      this.db.upsertNotification({
        userId: s.userId,
        topic,
        fileId: evt.fileId,
        changeId: evt.changeId,
        hash: evt.hash,
        windowStart,
        windowEnd,
        payloadJson,
        status: 'pending',
        createdAt,
        updatedAt: createdAt,
      });
    }
  }

  /**
   * Возвращает pending-уведомления, готовые к доставке, и помечает их доставленными.
   * Отправку наружу оставляем внешнему слою (SchedulerService), поэтому тут — только DB-операции.
   */
  async flushNotifications(now: number = nowMs()): Promise<Array<{ id: number; userId: string; topic: string; fileId: string; payload?: any }>> {
    const ready = this.db.listPendingNotificationsReady(now, 200);
    const result = ready.map(r => ({ id: r.id, userId: r.userId, topic: r.topic, fileId: r.fileId, payload: r.payloadJson ? JSON.parse(r.payloadJson) : undefined }));
    if (ready.length) this.db.markNotificationsDelivered(ready.map(r => r.id), now);
    return result;
  }

  /**
   * Сборка дайджеста за окно времени. Группируем по файлу, считаем кол-во изменений.
   * Возвращаем структуру, пригодную для Embed, без отправки.
   */
  async buildDigest(userId: string, _period: 'daily' | 'weekly', windowStart: number, windowEnd: number): Promise<{ items: Array<{ fileId: string; count: number }>; total: number }>{
    const notifs = this.db.listUserNotificationsInWindow(userId, windowStart, windowEnd);
    const map = new Map<string, number>();
    for (const n of notifs) {
      const payload = n.payloadJson ? JSON.parse(n.payloadJson) : undefined;
      const fileId = (payload && payload.fileId) || n.fileId;
      const prev = map.get(fileId) || 0;
      map.set(fileId, prev + 1);
    }
    const items = Array.from(map.entries()).map(([fileId, count]) => ({ fileId, count })).sort((a, b) => b.count - a.count).slice(0, 50);
    return { items, total: notifs.length };
  }

  createDigestRecord(userId: string, period: 'daily' | 'weekly', windowStart: number, windowEnd: number, payload: unknown, deliveredAt?: number | null): number {
    const payloadJson = JSON.stringify(payload ?? {});
    const id = this.db.createDigest({ userId, period, windowStart, windowEnd, payloadJson, createdAt: nowMs(), deliveredAt: deliveredAt ?? null });
    return id;
  }
}
