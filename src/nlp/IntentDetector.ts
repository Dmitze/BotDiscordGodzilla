import { normalizeText } from '@/nlp/normalize';
import { detectLanguage } from '@/nlp/LanguageDetector';
import logger from '@/utils/logger';

export type IntentKey = 'help' | 'search' | 'favorites' | 'workspace' | 'analytics' | 'ai';
export type Locale = 'uk' | 'en';

export interface IntentResult {
  intent?: IntentKey;
  confidence: number; // 0..1
  locale: Locale;
  matched?: string;
  source: 'rules' | 'ai' | 'none';
}

export interface ClassifyIntentOptions {
  locale: Locale;
  timeoutMs?: number;
  maxTokens?: number;
}

export type ClassifyIntentFn = (text: string, opts: ClassifyIntentOptions) => Promise<{ intent?: IntentKey; confidence?: number }>;

const INTENT_TRIGGERS: Record<IntentKey, { triggers: Record<Locale, string[]> }> = {
  help: {
    triggers: {
      uk: ['допомога', 'поміч', 'help'],
      en: ['help', 'support', 'assist'],
    },
  },
  search: {
    triggers: {
      uk: ['пошук', 'знайди', 'знайти'],
      en: ['search', 'find', 'lookup'],
    },
  },
  favorites: {
    triggers: {
      uk: ['вибране', 'улюблене'],
      en: ['favorites', 'stars', 'bookmarks'],
    },
  },
  workspace: {
    triggers: {
      uk: ['простір', 'workspace', 'робочий'],
      en: ['workspace', 'space', 'project'],
    },
  },
  analytics: {
    triggers: {
      uk: ['аналітика', 'статистика'],
      en: ['analytics', 'stats', 'statistics'],
    },
  },
  ai: {
    triggers: {
      uk: ['штучний інтелект', 'ai', 'бот'],
      en: ['ai', 'assistant', 'bot'],
    },
  },
};

function bestRuleMatch(text: string, locale: Locale): { intent?: IntentKey; confidence: number; matched?: string } {
  let best: { intent?: IntentKey; confidence: number; matched?: string } = { confidence: 0 };
  for (const key of Object.keys(INTENT_TRIGGERS) as IntentKey[]) {
    const list = INTENT_TRIGGERS[key].triggers[locale] || [];
    for (const phrase of list) {
      if (!phrase) continue;
      const idx = text.indexOf(phrase);
      if (idx >= 0) {
        const conf = Math.min(1, Math.max(0.6, phrase.length / Math.max(8, text.length)));
        if (conf > best.confidence) {
          best = { intent: key, confidence: conf, matched: phrase };
        }
      }
    }
  }
  return best;
}

export async function detectIntent(
  rawText: string,
  options: {
    classifyIntent?: ClassifyIntentFn;
    timeoutMs?: number;
    maxTokens?: number;
    defaultLocale?: Locale;
  } = {}
): Promise<IntentResult> {
  try {
    const normalized = normalizeText(rawText, { maxLength: 768 });
    const lang = detectLanguage(normalized);
    const locale: Locale = (lang === 'unknown' ? options.defaultLocale || 'uk' : (lang as Locale));

    // Rule-based first
    const rule = bestRuleMatch(normalized, locale);
    if (rule.intent && rule.confidence >= 0.6) {
      const res: IntentResult = { intent: rule.intent, confidence: rule.confidence, locale, source: 'rules' };
      if (rule.matched) {
        res.matched = rule.matched;
      }
      logger.debug('intent_rule', res as any);
      return res;
    }

    // AI fallback
    const classify = options.classifyIntent;
    if (!classify) {
      return { confidence: 0, locale, source: 'none' };
    }

    const timeoutMs = options.timeoutMs ?? 2000;
    const maxTokens = options.maxTokens ?? 128;

    const aiPromise = classify(normalized, { locale, timeoutMs, maxTokens });
    const timeout = new Promise<IntentResult>((resolve) => {
      setTimeout(() => resolve({ confidence: 0, locale, source: 'none' }), timeoutMs);
    });

    const aiRes = await Promise.race([
      aiPromise.then((r) => {
        const base: IntentResult = {
          confidence: typeof r.confidence === 'number' ? r.confidence : 0.7,
          locale,
          source: 'ai',
        };
        if (r.intent) base.intent = r.intent;
        return base;
      }),
      timeout,
    ]);

    logger.debug('intent_ai', aiRes as any);
    return aiRes;
  } catch (e) {
    logger.debug('intent_detect_failed', { error: e instanceof Error ? e.message : String(e) });
    return { confidence: 0, locale: 'uk', source: 'none' };
  }
}
