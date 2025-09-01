/**
 * UserWorkspaceService
 * In-memory + file-backed storage for per-user workspace items (files and searches).
 */

import { ensureDir, pathExists, readJSON, writeJSON } from 'fs-extra';
import { join } from 'path';
import { randomUUID } from 'crypto';

export type WorkspaceItemType = 'file' | 'query';

export interface WorkspaceItem {
  id: string;
  type: WorkspaceItemType;
  title: string;
  payload: { fileId?: string; query?: string };
  createdAt: string;
  tags?: string[];
}

export interface WorkspaceListFilter {
  type?: WorkspaceItemType;
}

const memStore = new Map<string, WorkspaceItem[]>();
const baseDir = join(process.cwd(), 'data', 'workspaces');

async function loadFromFile(userId: string): Promise<WorkspaceItem[]> {
  try {
    const file = join(baseDir, `${userId}.json`);
    if (!(await pathExists(file))) return [];
    const data = (await readJSON(file)) as WorkspaceItem[];
    return Array.isArray(data) ? data : [];
  } catch {
    return [];
  }
}

async function saveToFile(userId: string, items: WorkspaceItem[]): Promise<void> {
  const file = join(baseDir, `${userId}.json`);
  await ensureDir(baseDir);
  await writeJSON(file, items, { spaces: 2 });
}

async function getAll(userId: string): Promise<WorkspaceItem[]> {
  const cached = memStore.get(userId);
  if (cached) return cached;
  const loaded = await loadFromFile(userId);
  memStore.set(userId, loaded);
  return loaded;
}

export const UserWorkspaceService = {
  async addItem(
    userId: string,
    item: Omit<WorkspaceItem, 'id' | 'createdAt'> & Partial<Pick<WorkspaceItem, 'id' | 'createdAt'>>
  ): Promise<WorkspaceItem> {
    const now = new Date().toISOString();
    const newItemBase: Omit<WorkspaceItem, 'tags'> & Partial<Pick<WorkspaceItem, 'tags'>> = {
      id: item.id || randomUUID(),
      createdAt: item.createdAt || now,
      type: item.type,
      title: item.title,
      payload: item.payload,
    };
    const newItem: WorkspaceItem = ((): WorkspaceItem => {
      if (item.tags) {
        return { ...newItemBase, tags: item.tags } as WorkspaceItem;
      }
      return newItemBase as WorkspaceItem;
    })();

    const items = await getAll(userId);
    items.push(newItem);
    memStore.set(userId, items);
    await saveToFile(userId, items);
    return newItem;
  },

  async list(userId: string, filter?: WorkspaceListFilter): Promise<WorkspaceItem[]> {
    const items = await getAll(userId);
    if (!filter?.type) return items;
    return items.filter(i => i.type === filter.type);
  },

  async remove(userId: string, id: string): Promise<boolean> {
    const items = await getAll(userId);
    const next = items.filter(i => i.id !== id);
    const changed = next.length !== items.length;
    if (changed) {
      memStore.set(userId, next);
      await saveToFile(userId, next);
    }
    return changed;
  },

  async get(userId: string, id: string): Promise<WorkspaceItem | undefined> {
    const items = await getAll(userId);
    return items.find(i => i.id === id);
  },

  /** test-helper */
  async __reset(userId?: string): Promise<void> {
    if (userId) {
      memStore.delete(userId);
      await saveToFile(userId, []);
      return;
    }
    memStore.clear();
  },
};
