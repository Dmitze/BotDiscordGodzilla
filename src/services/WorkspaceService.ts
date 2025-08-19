import { ensureDir, pathExists, readJSON, writeJSON } from 'fs-extra';
import { dirname, resolve } from 'path';
import { tmpdir } from 'os';
import type { DriveListQuery, DriveListResult } from '@/types/drive';
import type { GoogleService } from './GoogleService';
import type { BotConfig } from '@/types';

export type Favorite = { fileId: string; addedAt: string };
export type SavedSearch = { name: string; filters: DriveListQuery; createdAt: string };

export interface WorkspaceSnapshot {
  favorites: Record<string, Favorite[]>; // by userId
  searches: Record<string, SavedSearch[]>; // by userId
}

function nowIso(): string { return new Date().toISOString(); }

/**
 * WorkspaceService: per-user favorites and saved searches
 * - In-memory storage with optional JSON persistence (single file)
 */
export class WorkspaceService {
  private persistPath: string | null = null;
  private favorites = new Map<string, Favorite[]>();
  private searches = new Map<string, SavedSearch[]>();

  constructor(persistPath?: string) {
    this.persistPath = persistPath ? resolve(persistPath) : null;
  }

  // Persistence helpers
  private async load(): Promise<void> {
    if (!this.persistPath) return;
    try {
      if (!(await pathExists(this.persistPath))) return;
      const json = (await readJSON(this.persistPath)) as WorkspaceSnapshot | undefined;
      if (!json) return;
      this.favorites = new Map(Object.entries(json.favorites || {}).map(([k, v]) => [k, Array.isArray(v) ? v : []]));
      this.searches = new Map(Object.entries(json.searches || {}).map(([k, v]) => [k, Array.isArray(v) ? v : []]));
    } catch {
      // ignore
    }
  }

  private async save(): Promise<void> {
    if (!this.persistPath) return;
    const snap: WorkspaceSnapshot = {
      favorites: Object.fromEntries(this.favorites),
      searches: Object.fromEntries(this.searches),
    };
    await ensureDir(dirname(this.persistPath));
    await writeJSON(this.persistPath, snap, { spaces: 2 });
  }

  /** Initialize from disk if persistPath provided */
  public async initialize(): Promise<void> {
    await this.load();
  }

  // Favorites
  public async addFavorite(userId: string, fileId: string): Promise<{ added: boolean; favorite: Favorite }>{
    const list = this.favorites.get(userId) || [];
    const exists = list.find(f => f.fileId === fileId);
    if (exists) return { added: false, favorite: exists };
    const fav: Favorite = { fileId, addedAt: nowIso() };
    list.push(fav);
    this.favorites.set(userId, list);
    await this.save();
    return { added: true, favorite: fav };
  }

  public async removeFavorite(userId: string, fileId: string): Promise<boolean> {
    const list = this.favorites.get(userId) || [];
    const next = list.filter(f => f.fileId !== fileId);
    const changed = next.length !== list.length;
    if (changed) {
      this.favorites.set(userId, next);
      await this.save();
    }
    return changed;
  }

  public async listFavorites(userId: string): Promise<Favorite[]> {
    return this.favorites.get(userId) || [];
  }

  // Saved searches
  public async saveSearch(userId: string, name: string, filters: DriveListQuery): Promise<{ created: boolean; search: SavedSearch }>{
    const list = this.searches.get(userId) || [];
    const exists = list.find(s => s.name.toLowerCase() === name.toLowerCase());
    if (exists) {
      // update filters
      exists.filters = filters;
      await this.save();
      return { created: false, search: exists };
    }
    const search: SavedSearch = { name, filters, createdAt: nowIso() };
    list.push(search);
    this.searches.set(userId, list);
    await this.save();
    return { created: true, search };
  }

  public async removeSearch(userId: string, name: string): Promise<boolean> {
    const list = this.searches.get(userId) || [];
    const lname = name.toLowerCase();
    const next = list.filter(s => s.name.toLowerCase() !== lname);
    const changed = next.length !== list.length;
    if (changed) {
      this.searches.set(userId, next);
      await this.save();
    }
    return changed;
  }

  public async listSearches(userId: string): Promise<SavedSearch[]> {
    return this.searches.get(userId) || [];
  }

  /** Execute saved search with policy enforcement */
  public async runSearch(
    userId: string,
    name: string,
    deps: { google: GoogleService; config: BotConfig }
  ): Promise<DriveListResult | undefined> {
    const list = this.searches.get(userId) || [];
    const item = list.find(s => s.name.toLowerCase() === name.toLowerCase());
    if (!item) return undefined;

    const base = { ...item.filters } as Partial<DriveListQuery>;
    // Enforce policies from config
    const cfg = deps.config.drive || {};
    if (Array.isArray(cfg.allowedMime) && cfg.allowedMime.length) {
      base.mimeIncludes = base.mimeIncludes && base.mimeIncludes.length
        ? base.mimeIncludes.filter(m => cfg.allowedMime!.includes(m))
        : cfg.allowedMime;
    }
    if (Array.isArray(cfg.ownerAllowlist) && cfg.ownerAllowlist.length) {
      base.ownerAllowlist = cfg.ownerAllowlist;
    }
    // Use default folder if not specified
    if (!base.folderId) {
      base.folderId = deps.config.google?.driveFolderId ?? deps.config.drive?.folderId ?? undefined;
    }
    if (!base.folderId) return undefined;
    return deps.google.listDriveFiles(base as DriveListQuery);
  }

  // Test helper
  public async __reset(): Promise<void> {
    this.favorites.clear();
    this.searches.clear();
    if (this.persistPath) {
      await ensureDir(dirname(this.persistPath));
      await writeJSON(this.persistPath, { favorites: {}, searches: {} }, { spaces: 2 });
    }
  }
}

// Default singleton for commands if container isn't wired yet
export const defaultWorkspaceService = new WorkspaceService(resolve(tmpdir(), 'workspace.json'));
(async () => { try { await defaultWorkspaceService.initialize(); } catch {} })();
