import logger from '@/utils/logger';

export type IntentType = 'ANALYZE_SHEET' | 'ANALYZE_FILE' | 'QNA_GENERAL' | 'HELP' | 'UNKNOWN';

export interface DetectedIntent {
  type: IntentType;
  params?: Record<string, string>;
  confidence: number; // 0..1
}

/**
 * Простой эвристический детектор интентов.
 * Можно заменить/усилить локальным LLM через AIService.classify в будущем.
 */
export class IntentDetector {
  detect(text: string): DetectedIntent {
    const q = (text || '').trim().toLowerCase();
    if (!q) return { type: 'UNKNOWN', confidence: 0 };

    try {
      if (this.isHelp(q)) return { type: 'HELP', confidence: 0.9 };

      const sheet = this.tryExtractSheet(q);
      if (sheet) return { type: 'ANALYZE_SHEET', confidence: 0.75, params: sheet };

      const file = this.tryExtractFile(q);
      if (file) return { type: 'ANALYZE_FILE', confidence: 0.7, params: file };

      if (this.isQna(q)) return { type: 'QNA_GENERAL', confidence: 0.6 };

      return { type: 'UNKNOWN', confidence: 0.2 };
    } catch (e) {
      logger.warn('intent_detect_failed', {
        type: 'chat',
        component: 'IntentDetector',
        error: e instanceof Error ? e.message : String(e),
      });
      return { type: 'UNKNOWN', confidence: 0 };
    }
  }

  private isHelp(q: string): boolean {
    return /^help$|^помощ/i.test(q);
  }

  private tryExtractSheet(q: string): Record<string, string> | undefined {
    if (/(проанализируй|анализ|подсчитай|сколько|сумма|средн|медиан|таблиц)/i.test(q)) {
      const m = q.match(/["'«](.+?)["'»]/);
      const name = m?.[1];
      return name ? { sheet: name } : {};
    }
    return undefined;
  }

  private tryExtractFile(q: string): Record<string, string> | undefined {
    if (/(файл|документ|drive|google drive|pdf|docx|doc)/i.test(q)) {
      const idMatch = q.match(/[a-z0-9_-]{20,}/i);
      return idMatch ? { id: idMatch[0] } : {};
    }
    return undefined;
  }

  private isQna(q: string): boolean {
    return /[?]|(что|как|почему|зачем|когда|где)\s/i.test(q);
  }
}
