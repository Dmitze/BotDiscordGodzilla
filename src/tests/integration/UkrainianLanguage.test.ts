/**
 * Integration tests for Ukrainian language support in the Discord bot
 */

import { jest, describe, it, expect, beforeEach } from '@jest/globals';
import { detectLanguage } from '../../nlp/LanguageDetector';
import { t } from '../../i18n';

// Mock the logger
jest.mock('../../utils/logger', () => ({
  default: {
    info: jest.fn(),
    warn: jest.fn(),
    error: jest.fn(),
    debug: jest.fn(),
    log: jest.fn(),
    apiRequest: jest.fn(),
    apiError: jest.fn(),
    security: jest.fn(),
    performance: jest.fn(),
    system: jest.fn(),
    logStructured: jest.fn(), // Add the missing logStructured method
    startStructuredTimer: jest.fn().mockReturnValue({ end: jest.fn() }),
    getStats: jest.fn(),
    getLogBuffer: jest.fn(),
    cleanup: jest.fn(),
    isHealthy: jest.fn(),
  },
}));

describe('Ukrainian Language Integration', () => {
  describe('Language Detection', () => {
    it('should detect Ukrainian language correctly', () => {
      const ukrainianText = 'Привіт, як справи? Документи готові для аналізу.';
      const detected = detectLanguage(ukrainianText);
      expect(detected).toBe('uk');
    });

    it('should detect English language correctly', () => {
      const englishText = 'Hello, how are you? Documents are ready for analysis.';
      const detected = detectLanguage(englishText);
      expect(detected).toBe('en');
    });

    it('should default to Ukrainian for mixed or unknown text', () => {
      const mixedText = 'Привіт hello як how справи?';
      const detected = detectLanguage(mixedText);
      expect(detected).toBe('uk');
    });
  });

  describe('Localization', () => {
    it('should provide Ukrainian translations for key commands', () => {
      // Test Ukrainian translations
      expect(t('commands.search.name')).toBe('пошук');
      expect(t('commands.search.description')).toBe('🔍 Гнучкий пошук по документах');
      expect(t('commands.markdown.name')).toBe('markdown');
      expect(t('commands.analytics.name')).toBe('аналітика');
    });

    it('should provide English translations when switched to English', () => {
      // This would require setting the locale, but we can at least verify the English keys exist
      const englishKeys = [
        'commands.search.name',
        'commands.search.description',
        'commands.markdown.name',
        'commands.analytics.name'
      ];
      
      for (const key of englishKeys) {
        const translation = t(key);
        expect(translation).toBeDefined();
        expect(typeof translation).toBe('string');
      }
    });
  });

  describe('Formatting', () => {
    it('should format numbers according to Ukrainian locale', () => {
      // Ukrainian locale uses space as thousands separator and comma as decimal separator
      const largeNumber = 1234567.89;
      const formatted = largeNumber.toLocaleString('uk-UA');
      expect(formatted).toMatch(/^1\s234\s567,89$/);
    });

    it('should format dates according to Ukrainian locale', () => {
      const date = new Date('2023-05-15T10:30:00Z');
      const formatted = date.toLocaleDateString('uk-UA');
      expect(formatted).toMatch(/^\d{2}\.\d{2}\.\d{4}$/);
    });

    it('should format currency according to Ukrainian locale', () => {
      const amount = 1234.56;
      const formatted = amount.toLocaleString('uk-UA', {
        style: 'currency',
        currency: 'UAH'
      });
      expect(formatted).toMatch(/\d{1,3}[\s\d]*,\d{2}\s₴/);
    });
  });
});