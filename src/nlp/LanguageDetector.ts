import logger from '@/utils/logger';

export type DetectedLang = 'uk' | 'en' | 'unknown';

const UK_STOP = new Set(['і', 'та', 'в', 'на', 'що', 'як', 'це', 'цею', 'цей', 'ця', 'чи', 'але', 'або']);
const EN_STOP = new Set(['the', 'and', 'in', 'on', 'is', 'it', 'this', 'that', 'or', 'but', 'to', 'of']);

const CYRILLIC_RE = /[\u0400-\u04FF]/; // Cyrillic block
const LATIN_RE = /[A-Za-z]/;

export function detectLanguage(text: string): DetectedLang {
  try {
    if (!text) return 'unknown';
    const sample = text.slice(0, 512);

    const hasCyr = CYRILLIC_RE.test(sample);
    const hasLat = LATIN_RE.test(sample);

    if (hasCyr && !hasLat) return 'uk';
    if (hasLat && !hasCyr) return 'en';

    // Mixed: use stop-words heuristic
    const tokens = sample
      .toLowerCase()
      .replace(/[^\p{L}\p{N}\s]/gu, ' ')
      .split(/\s+/)
      .filter(Boolean)
      .slice(0, 50);

    let ukScore = 0;
    let enScore = 0;
    for (const t of tokens) {
      if (UK_STOP.has(t)) ukScore += 1;
      if (EN_STOP.has(t)) enScore += 1;
    }

    const lang: DetectedLang = ukScore > enScore ? 'uk' : enScore > ukScore ? 'en' : 'unknown';
    // Safely call logStructured if available
    if (logger && typeof logger.logStructured === 'function') {
      logger.logStructured('debug', 'language_detect', { ukScore, enScore, lang });
    }
    return lang;
  } catch (e) {
    // Safely call logStructured if available
    if (logger && typeof logger.logStructured === 'function') {
      logger.logStructured('debug', 'language_detect_failed', { error: e instanceof Error ? e.message : String(e) });
    }
    return 'unknown';
  }
}