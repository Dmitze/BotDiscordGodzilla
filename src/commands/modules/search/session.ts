import type { PaginationState } from './types';

let externalMap: Map<string, PaginationState> | null = null;
const internal = new Map<string, PaginationState>();

export function bindSessionMap(map: Map<string, PaginationState>): void {
  externalMap = map;
}

function store(): Map<string, PaginationState> {
  return externalMap ?? internal;
}

export function getSession(id: string): PaginationState | undefined {
  return store().get(id);
}

export function setSession(id: string, state: PaginationState): void {
  store().set(id, state);
}

export function deleteSession(id: string): void {
  store().delete(id);
}

export function size(): number {
  return store().size;
}

export function cleanupExpired(ttlSec: number, nowSec: number = Math.floor(Date.now() / 1000)): number {
  const s = store();
  let removed = 0;
  for (const [sid, state] of s.entries()) {
    if (nowSec - state.timestamp > ttlSec) {
      s.delete(sid);
      removed++;
    }
  }
  return removed;
}
