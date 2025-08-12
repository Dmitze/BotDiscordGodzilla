/**
 * SheetsContextService
 * Хранит выбор Spreadsheet/Sheet на пользователя/канал/guild с TTL.
 */

import { BaseService } from '@/core/BaseService';
import type { BotConfig } from '@/types';
import logger from '@/utils/logger';

interface ContextKey {
  userId?: string;
  channelId?: string;
  guildId?: string;
}

export interface SheetContext {
  spreadsheetId: string;
  sheetName?: string;
  updatedAt: number;
}

/** Простая запись с TTL для in-memory */
interface TTLRecord<T> {
  value: T;
  expiresAt: number;
}

export class SheetsContextService extends BaseService {
  private memoryStore: Map<string, TTLRecord<SheetContext>> = new Map();
  private defaultTTL: number;
  private cleanupTimer: NodeJS.Timeout | null = null;

  constructor(config: BotConfig) {
    super('SheetsContextService', config);
    // TTL из конфигурации производительности, fallback 30 минут
    this.defaultTTL = Math.max(60, Number(config?.performance?.cacheTTL) || 1800);
  }

  protected async onInitialize(): Promise<void> {
    // Периодическая очистка in-memory
    this.cleanupTimer = setInterval(() => {
      try {
        const now = Date.now();
        for (const [key, rec] of this.memoryStore.entries()) {
          if (rec.expiresAt <= now) this.memoryStore.delete(key);
        }
      } catch (e) {
        logger.warn('SheetsContextService: cleanup error', { component: 'SheetsContextService', error: String(e) });
      }
    }, 60_000);

    logger.info('✅ SheetsContextService инициализирован', { component: 'SheetsContextService', ttl: this.defaultTTL });
  }

  protected async onShutdown(): Promise<void> {
    if (this.cleanupTimer) {
      clearInterval(this.cleanupTimer);
      this.cleanupTimer = null;
    }
    this.memoryStore.clear();
  }

  protected async onHealthCheck() {
    return { healthy: true } as any;
  }

  protected onGetStats() {
    return { entries: this.memoryStore.size } as any;
  }

  private buildKeys(key: ContextKey): string[] {
    const keys: string[] = [];
    if (key.userId) keys.push(`user:${key.userId}`);
    if (key.channelId) keys.push(`channel:${key.channelId}`);
    if (key.guildId) keys.push(`guild:${key.guildId}`);
    return keys;
  }

  private primaryKey(key: ContextKey): string {
    if (key.userId) return `user:${key.userId}`;
    if (key.channelId) return `channel:${key.channelId}`;
    if (key.guildId) return `guild:${key.guildId}`;
    throw new Error('SheetsContextService: пустой ключ контекста');
  }

  public async setContext(key: ContextKey, ctx: Omit<SheetContext, 'updatedAt'>, ttlSec?: number): Promise<void> {
    const k = this.primaryKey(key);
    const ttl = Math.max(30, ttlSec ?? this.defaultTTL);
    this.memoryStore.set(k, { value: { ...ctx, updatedAt: Date.now() }, expiresAt: Date.now() + ttl * 1000 });
    logger.debug('SheetsContextService: контекст сохранён', { component: 'SheetsContextService', key: k, ttl });
  }

  public async getContext(key: ContextKey): Promise<SheetContext | null> {
    const now = Date.now();
    for (const k of this.buildKeys(key)) {
      const rec = this.memoryStore.get(k);
      if (rec && rec.expiresAt > now) return rec.value;
      if (rec && rec.expiresAt <= now) this.memoryStore.delete(k);
    }
    return null;
  }

  public async clearContext(key: ContextKey): Promise<boolean> {
    let removed = false;
    for (const k of this.buildKeys(key)) {
      removed = this.memoryStore.delete(k) || removed;
    }
    return removed;
  }
}

export default SheetsContextService;
