import logger from '@/utils/logger';

export interface UIStateEntry<T> {
  value: T;
  expiresAt: number; // epoch ms
}

/**
 * In-memory UI state with TTL. No timers to avoid open handles in tests.
 */
export class UIStateService {
  private store = new Map<string, UIStateEntry<unknown>>();

  set<T>(key: string, value: T, ttlSec = 300): void {
    const expiresAt = Date.now() + ttlSec * 1000;
    this.store.set(key, { value, expiresAt });
  }

  get<T>(key: string): T | null {
    const entry = this.store.get(key);
    if (!entry) return null;
    if (Date.now() > entry.expiresAt) {
      this.store.delete(key);
      return null;
    }
    return entry.value as T;
  }

  delete(key: string): void {
    this.store.delete(key);
  }

  cleanup(): void {
    const now = Date.now();
    let removed = 0;
    for (const [k, v] of this.store.entries()) {
      if (now > v.expiresAt) {
        this.store.delete(k);
        removed++;
      }
    }
    if (removed > 0) {
      logger.debug('UIStateService cleanup removed entries', { removed });
    }
  }

  makeKey(parts: { scope: string; userId: string; nonce: string }): string {
    return `${parts.scope}:${parts.userId}:${parts.nonce}`;
  }
}

export const uiState = new UIStateService();
