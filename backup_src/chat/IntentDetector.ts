import logger from '@/utils/logger';
import LanguageDetector, { type LanguageCode } from '@/chat/LanguageDetector';
import type { AIService } from '@/services/AIService';

export type IntentType = 'SEARCH' | 'ANALYZE_SHEET' | 'ANALYZE_FILE' | 'QNA_GENERAL' | 'HELP' | 'UNKNOWN';

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
  constructor(
    ai?: AIService,
    opts: { aiTimeoutMs?: number; aiMaxTokens?: number } = {}
  ) {
    if (ai) this.ai = ai; // Guard assignment for exactOptionalPropertyTypes
    this.opts = opts;
  }

  private ai?: AIService;
  private opts: { aiTimeoutMs?: number; aiMaxTokens?: number } = {};

  // Allow late injection from Bot/ServiceManager
  public setAI(ai?: AIService): void {
    if (ai) this.ai = ai;
  }

  public setOptions(opts: { aiTimeoutMs?: number; aiMaxTokens?: number }): void {
    this.opts = { ...this.opts, ...opts };
  }

  detect(text: string): DetectedIntent {
    const raw = (text || '').trim();
    if (!raw) return { type: 'UNKNOWN', confidence: 0 };

    const norm = this.normalize(raw);
    const lang: LanguageCode = LanguageDetector.detectLanguage(raw);

    try {
      // help
      if (this.isHelp(norm, lang)) return { type: 'HELP', confidence: 0.9 };

      // search
      const search = this.tryExtractSearch(norm, lang);
      if (search) return { type: 'SEARCH', confidence: 0.8, params: search };

      // analyze sheet
      const sheet = this.tryExtractSheet(norm, lang);
      if (sheet) return { type: 'ANALYZE_SHEET', confidence: 0.75, params: sheet };

      // analyze file
      const file = this.tryExtractFile(norm, lang);
      if (file) return { type: 'ANALYZE_FILE', confidence: 0.7, params: file };

      // general QnA
      if (this.isQna(norm, lang)) return { type: 'QNA_GENERAL', confidence: 0.6 };

      // optional AI fallback
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

  async detectWithAI(text: string): Promise<DetectedIntent> {
    const local = this.detect(text);
    if (local.type !== 'UNKNOWN' && local.confidence >= 0.5) return local;
    if (!this.ai) return local;

    try {
      const prompt = this.buildClassificationPrompt(text);
      const timeoutMs = this.opts.aiTimeoutMs ?? 3000;
      let timer: NodeJS.Timeout | undefined;
      const timeoutPromise = new Promise<never>((_, reject) => {
        timer = setTimeout(() => reject(new Error('intent_ai_timeout')), timeoutMs);
      });
      const ai = this.ai; // narrowed by guard above
      const res = (await Promise.race([
        ai.generateResponse(prompt, { maxTokens: this.opts.aiMaxTokens ?? 128 }),
        timeoutPromise,
      ])) as { content: string };
      if (timer) clearTimeout(timer);
      const parsed = this.parseAIResult(res.content);
      return parsed ?? local;
    } catch (e) {
      logger.warn('intent_ai_fallback_failed', {
        error: e instanceof Error ? e.message : String(e),
      });
      return local;
    }
  }

  // --- helpers ---
  private normalize(input: string): string {
    return input
      .replace(/[\u200B-\u200D\uFEFF]/g, '') // zero-width chars
      .replace(/[\n\r\t]+/g, ' ')
      .replace(/[\p{P}\p{S}]+/gu, ' ') // punctuation/symbols to space
      .toLowerCase()
      .replace(/\s{2,}/g, ' ')
      .trim();
  }

  private isHelp(q: string, lang: LanguageCode): boolean {
    const patterns: Record<LanguageCode, RegExp> = {
      uk: /^(допомога|help)$/i,
      en: /^(help)$/i,
    };
    return (patterns[lang] || patterns.uk).test(q);
  }

  private tryExtractSearch(q: string, lang: LanguageCode): Record<string, string> | undefined {
    const re = lang === 'en'
      ? /(search|find)/i
      : /(пошук|знайд[ий]|знайти)/i;
    if (re.test(q)) {
      const quoted = q.match(/["'«](.+?)["'»]/);
      if (quoted?.[1]) return { query: quoted[1] };
      const m = q.match(new RegExp(`${re.source}\\s+(.+)`, 'i'));
      if (m?.[1]) return { query: m[1].trim() };
      return { query: q };
    }
    return undefined;
  }

  private tryExtractSheet(q: string, lang: LanguageCode): Record<string, string> | undefined {
    const re = lang === 'en'
      ? /(analy[sz]e|sum|average|median|table|sheet)/i
      : /(проаналізуй|аналіз|підрахуй|скільки|сума|середн|медіан|таблиц|аркуш|лист)/i;
    if (re.test(q)) {
      const m = q.match(/["'«](.+?)["'»]/);
      const name = m?.[1];
      return name ? { sheet: name } : {};
    }
    return undefined;
  }

  private tryExtractFile(q: string, _lang: LanguageCode): Record<string, string> | undefined {
    if (/(file|document|docx|pdf|drive|google drive|документ|файл)/i.test(q)) {
      const idMatch = q.match(/[a-z0-9_-]{20,}/i);
      return idMatch ? { id: idMatch[0] } : {};
    }
    return undefined;
  }

  private isQna(q: string, lang: LanguageCode): boolean {
    if (/[?]/.test(q)) return true;
    const re = lang === 'en'
      ? /(what|how|why|when|where)\s/i
      : /(що|як|чому|навіщо|коли|де)\s/i;
    return re.test(q);
  }

  private buildClassificationPrompt(text: string): string {
    return (
      'Classify the user input into one of intents: SEARCH, ANALYZE_SHEET, ANALYZE_FILE, QNA_GENERAL, HELP, UNKNOWN. '
      + 'Respond strictly as compact JSON: {"type":"INTENT","confidence":0.0,"params":{}}. '
      + 'If extracting params, use keys: query (for search), sheet (for sheet name), id (for file id). '
      + 'User input: ' + JSON.stringify(text)
    );
  }

  private parseAIResult(s: string): DetectedIntent | null {
    try {
      const jsonStart = s.indexOf('{');
      const jsonEnd = s.lastIndexOf('}');
      const raw = jsonStart >= 0 && jsonEnd > jsonStart ? s.slice(jsonStart, jsonEnd + 1) : s;
      const obj = JSON.parse(raw) as Partial<DetectedIntent & { type: string }>;
      const type = (obj.type || 'UNKNOWN') as IntentType;
      const confidence = typeof obj.confidence === 'number' ? obj.confidence : 0.5;
      const params = (obj.params && typeof obj.params === 'object') ? (obj.params) : undefined;
      const base: DetectedIntent = { type, confidence };
      return params ? { ...base, params } : base;
    } catch {
      return null;
    }
  }
}
