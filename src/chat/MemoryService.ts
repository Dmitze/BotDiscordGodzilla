import logger from '@/utils/logger';

export interface ChatTurn {
  role: 'user' | 'assistant' | 'system';
  content: string;
  ts: number;
}

export interface ChatContext {
  turns: ChatTurn[];
  tokenEstimate: number;
}

/**
 * Простейшая in-memory память диалога по ключу канал+пользователь.
 * Оценка токенов грубая: ~4 символа = 1 токен.
 */
export class MemoryService {
  private store = new Map<string, ChatContext>();
  private readonly maxTokens: number;
  private readonly summaryAfter: number;

  constructor(options?: { maxTokens?: number; summaryAfter?: number }) {
    this.maxTokens = Math.max(200, options?.maxTokens ?? 2000);
    this.summaryAfter = Math.max(100, options?.summaryAfter ?? 1500);
  }

  private key(channelId: string, userId: string): string {
    return `${channelId}:${userId}`;
  }

  getContext(channelId: string, userId: string): ChatContext {
    const k = this.key(channelId, userId);
    let ctx = this.store.get(k);
    if (!ctx) {
      ctx = { turns: [], tokenEstimate: 0 };
      this.store.set(k, ctx);
    }
    return ctx;
  }

  addTurn(channelId: string, userId: string, turn: ChatTurn): ChatContext {
    const ctx = this.getContext(channelId, userId);
    ctx.turns.push(turn);
    ctx.tokenEstimate += this.estimateTokens(turn.content);
    this.trim(ctx);
    return ctx;
  }

  reset(channelId: string, userId: string): void {
    this.store.delete(this.key(channelId, userId));
  }

  private estimateTokens(text: string): number {
    return Math.ceil((text || '').length / 4);
  }

  private trim(ctx: ChatContext): void {
    try {
      if (ctx.tokenEstimate <= this.maxTokens) return;
      // Простая стратегия: удаляем старые ходы, пока не уложимся.
      while (ctx.turns.length > 1 && ctx.tokenEstimate > this.summaryAfter) {
        const t = ctx.turns.shift();
        if (!t) break;
        ctx.tokenEstimate -= this.estimateTokens(t.content);
      }
      // Если всё ещё много — ужимаем содержимое
      if (ctx.tokenEstimate > this.maxTokens && ctx.turns.length) {
        const last = ctx.turns[0];
        last.content = last.content.slice(0, Math.max(0, this.maxTokens * 4 - 64));
        ctx.tokenEstimate = this.summaryAfter; // грубая коррекция
      }
    } catch (e) {
      logger.warn('memory_trim_failed', {
        type: 'chat',
        component: 'MemoryService',
        error: e instanceof Error ? e.message : String(e),
      });
    }
  }
}
