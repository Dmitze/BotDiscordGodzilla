import logger from '@/utils/logger';

export async function withTimeout<T>(promise: Promise<T>, ms: number, label = 'operation'): Promise<T> {
  let timer: NodeJS.Timeout | undefined;
  const timeout = new Promise<never>((_, reject) => {
    timer = setTimeout(() => reject(new Error(`${label} timed out after ${ms}ms`)), ms);
  });
  try {
    const res = await Promise.race([promise, timeout]);
    return res as T;
  } finally {
    if (timer) clearTimeout(timer);
  }
}

export async function retry<T>(fn: () => Promise<T>, attempts = 3, delayMs = 200, label = 'retry'): Promise<T> {
  let lastErr: unknown;
  for (let i = 0; i < attempts; i++) {
    try {
      return await fn();
    } catch (e) {
      lastErr = e;
      if (i < attempts - 1) {
        logger.warn(`${label}: attempt ${i + 1} failed, retrying...`, {
          error: e instanceof Error ? e.message : String(e),
        });
        await new Promise(r => setTimeout(r, delayMs));
      }
    }
  }
  throw lastErr instanceof Error ? lastErr : new Error(String(lastErr));
}
