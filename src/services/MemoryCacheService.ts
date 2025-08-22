/**
 * In-memory Cache Service for tests and local dev
 * API-compatible subset of CacheService used by tests/commands
 */
import type { BotConfig, HealthStatus, ServiceStats } from '@/types';
import { BaseService as BaseServiceClass } from '@/core/BaseService';

interface Entry<T> { value: T; expiresAt?: number }

export class MemoryCacheService extends BaseServiceClass {
  private store = new Map<string, Entry<unknown>>();
  private stats: ServiceStats;

  constructor(config: BotConfig) {
    super('MemoryCacheService', config);
    this.stats = { service: 'MemoryCacheService', uptime: 0, requests: 0, errors: 0 };
  }

  protected override async onInitialize(): Promise<void> {
    // no-op
  }

  protected override async onShutdown(): Promise<void> {
    this.store.clear();
  }

  public async get<T = unknown>(key: string): Promise<T | null> {
    const e = this.store.get(key);
    if (!e) return null;
    if (e.expiresAt && Date.now() > e.expiresAt) { this.store.delete(key); return null; }
    return e.value as T;
    }

  public async set<T = unknown>(key: string, value: T, ttlSec?: number): Promise<void> {
    const expiresAt = ttlSec ? Date.now() + ttlSec * 1000 : undefined;
    // Avoid setting optional property to undefined (exactOptionalPropertyTypes)
    const entry: Entry<T> = expiresAt !== undefined ? { value, expiresAt } : { value };
    this.store.set(key, entry as Entry<unknown>);
  }

  public async delete(key: string): Promise<void> { this.store.delete(key); }

  public async clear(): Promise<void> { this.store.clear(); }

  protected override async onHealthCheck(): Promise<HealthStatus> {
    return { healthy: true, service: this.name } as HealthStatus;
  }

  protected override onGetStats(): ServiceStats {
    return {
      ...this.stats,
      service: this.name,
      uptime: this.stats.uptime,
    };
  }
}
