/**
 * Simple segment/text normalization cache keyed by contentHash.
 * In-memory with TTL to reduce redundant normalization costs.
 * Can be replaced by a SQLite-backed cache later.
 */

export interface SegmentCacheValue {
  text: string;
  normText?: string;
  updatedAt: number; // epoch ms
}

export interface SegmentCache {
  get(contentHash: string): SegmentCacheValue | undefined;
  set(contentHash: string, value: SegmentCacheValue, ttlMs?: number): void;
  delete(contentHash: string): void;
  clear(): void;
}

export class InMemorySegmentCache implements SegmentCache {
  private store = new Map<string, { v: SegmentCacheValue; exp?: number }>();

  constructor(private defaultTtlMs: number = 10 * 60 * 1000) {}

  get(contentHash: string): SegmentCacheValue | undefined {
    const e = this.store.get(contentHash);
    if (!e) return undefined;
    if (e.exp && e.exp < Date.now()) {
      this.store.delete(contentHash);
      return undefined;
    }
    return e.v;
  }

  set(contentHash: string, value: SegmentCacheValue, ttlMs?: number): void {
    const exp = (ttlMs ?? this.defaultTtlMs) > 0 ? Date.now() + (ttlMs ?? this.defaultTtlMs) : undefined;
    if (exp !== undefined) {
      this.store.set(contentHash, { v: value, exp });
    } else {
      this.store.set(contentHash, { v: value });
    }
  }

  delete(contentHash: string): void {
    this.store.delete(contentHash);
  }

  clear(): void {
    this.store.clear();
  }
}
