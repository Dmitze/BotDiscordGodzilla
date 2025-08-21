export type LanguageCode = 'uk' | 'en';

/**
 * Простий детектор мови на основі евристик.
 * Перевага — українська (вимога проєкту), якщо чітких сигналів немає.
 */
export class LanguageDetector {
  static detectLanguage(text: string | undefined | null): LanguageCode {
    const s = (text ?? '').trim();
    if (!s) return 'uk';

    // Евристики:
    // - наявність специфічних українських літер
    // - частотні слова для uk/en
    const hasUASpecific = /[ієїґІЄЇҐ]/.test(s);
    if (hasUASpecific) return 'uk';

    const uaWords = /(будь ласка|допомога|пошук|знайди|таблиц|аркуш|лист|документ|файл|проаналізуй)/i;
    const enWords = /(please|help|search|find|sheet|table|document|file|analy[sz]e)/i;

    const uaScore = uaWords.test(s) ? 1 : 0;
    const enScore = enWords.test(s) ? 1 : 0;

    if (uaScore > enScore) return 'uk';
    if (enScore > uaScore) return 'en';

    // За замовчуванням — українська
    return 'uk';
  }
}

export default LanguageDetector;
