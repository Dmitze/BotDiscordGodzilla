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
      // HELP
      if (/^help$|^помощ/i.test(q)) return { type: 'HELP', confidence: 0.9 };

      // ANALYZE_SHEET
      if (/(проанализируй|анализ|подсчитай|сколько|сумма|средн|медиан|таблиц)/i.test(q)) {
        // Пытаемся вытащить имя листа/таблицы в кавычках
        const m = q.match(/["'«](.+?)["'»]/);
        const name = m?.[1];
        return { type: 'ANALYZE_SHEET', confidence: 0.75, params: name ? { sheet: name } : {} };
      }

      // ANALYZE_FILE
      if (/(файл|документ|drive|google drive|pdf|docx|doc)/i.test(q)) {
        const idMatch = q.match(/[a-z0-9_-]{20,}/i);
        return { type: 'ANALYZE_FILE', confidence: 0.7, params: idMatch ? { id: idMatch[0] } : {} };
      }

      // QNA_GENERAL
      if (/[?]|(что|как|почему|зачем|когда|где)\s/i.test(q)) {
        return { type: 'QNA_GENERAL', confidence: 0.6 };
      }

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
}
