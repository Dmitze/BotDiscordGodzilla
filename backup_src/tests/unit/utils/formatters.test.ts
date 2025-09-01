/**
 * Unit тесты для утилиты formatters
 */

import { describe, it, expect } from '@jest/globals';

describe('Formatters Utils', () => {
  describe('formatDate', () => {
    it('should format date correctly', () => {
      const date = new Date('2024-01-15T10:30:00Z');
      const formatted = formatDate(date);
      
      expect(formatted).toMatch(/^\d{2}\.\d{2}\.\d{4}$/);
    });

    it('should format current date', () => {
      const now = new Date();
      const formatted = formatDate(now);
      
      expect(formatted).toMatch(/^\d{2}\.\d{2}\.\d{4}$/);
    });

    it('should handle invalid date', () => {
      const invalidDate = new Date('invalid');
      const formatted = formatDate(invalidDate);
      
      expect(formatted).toBe('Н/Д');
    });
  });

  describe('formatDateTime', () => {
    it('should format date and time correctly', () => {
      const date = new Date('2024-01-15T10:30:00Z');
      const formatted = formatDateTime(date);
      
      // uk-UA: date, space then time with seconds, separated by comma
      expect(formatted).toMatch(/^\d{2}\.\d{2}\.\d{4}, \d{2}:\d{2}:\d{2}$/);
    });

    it('should format current date and time', () => {
      const now = new Date();
      const formatted = formatDateTime(now);
      
      expect(formatted).toMatch(/^\d{2}\.\d{2}\.\d{4}, \d{2}:\d{2}:\d{2}$/);
    });
  });

  describe('formatNumber', () => {
    it('should format large numbers with separators', () => {
      // uk-UA uses space (incl. NBSP) as thousands separator
      expect(formatNumber(1234567)).toMatch(/^1[ \u00A0]234[ \u00A0]567$/);
      expect(formatNumber(1000000)).toMatch(/^1[ \u00A0]000[ \u00A0]000$/);
    });

    it('should format decimal numbers', () => {
      // uk-UA uses comma as decimal separator
      expect(formatNumber(1234.56)).toMatch(/^1[ \u00A0]234,56$/);
      expect(formatNumber(0.123)).toBe('0,123');
    });

    it('should handle zero', () => {
      expect(formatNumber(0)).toBe('0');
    });

    it('should handle negative numbers', () => {
      expect(formatNumber(-1234)).toMatch(/^-1[ \u00A0]234$/);
    });
  });

  describe('formatFileSize', () => {
    it('should format bytes', () => {
      expect(formatFileSize(512)).toBe('512 B');
      expect(formatFileSize(1023)).toBe('1023 B');
    });

    it('should format kilobytes', () => {
      expect(formatFileSize(1024)).toBe('1 KB');
      expect(formatFileSize(1536)).toBe('1.5 KB');
    });

    it('should format megabytes', () => {
      expect(formatFileSize(1048576)).toBe('1 MB');
      expect(formatFileSize(1572864)).toBe('1.5 MB');
    });

    it('should format gigabytes', () => {
      expect(formatFileSize(1073741824)).toBe('1 GB');
      expect(formatFileSize(1610612736)).toBe('1.5 GB');
    });

    it('should handle zero', () => {
      expect(formatFileSize(0)).toBe('0 B');
    });
  });

  describe('formatDuration', () => {
    it('should format seconds', () => {
      expect(formatDuration(30)).toBe('30с');
      expect(formatDuration(59)).toBe('59с');
    });

    it('should format minutes', () => {
      expect(formatDuration(60)).toBe('1хв');
      expect(formatDuration(90)).toBe('1хв 30с');
      // current implementation prefers hours over 60 minutes
      expect(formatDuration(3600)).toBe('1год');
    });

    it('should format hours', () => {
      // 3601s -> 1 hour and 1 second, minutes remain 0
      expect(formatDuration(3601)).toBe('1год 1с');
      expect(formatDuration(3661)).toBe('1год 1хв 1с');
    });

    it('should format days', () => {
      expect(formatDuration(86400)).toBe('1д');
      expect(formatDuration(90000)).toBe('1д 1год');
    });

    it('should handle zero', () => {
      expect(formatDuration(0)).toBe('0с');
    });
  });

  describe('truncateText', () => {
    it('should truncate long text', () => {
      const longText = 'This is a very long text that needs to be truncated';
      const truncated = truncateText(longText, 20);
      
      expect(truncated.length).toBeLessThanOrEqual(23); // 20 + '...'
      expect(truncated.endsWith('...')).toBe(true);
    });

    it('should not truncate short text', () => {
      const shortText = 'Short text';
      const result = truncateText(shortText, 20);
      
      expect(result).toBe(shortText);
    });

    it('should handle empty string', () => {
      expect(truncateText('', 10)).toBe('');
    });

    it('should handle null and undefined', () => {
      expect(truncateText(null, 10)).toBe('');
      expect(truncateText(undefined, 10)).toBe('');
    });
  });

  describe('capitalizeFirst', () => {
    it('should capitalize first letter', () => {
      expect(capitalizeFirst('hello')).toBe('Hello');
      expect(capitalizeFirst('world')).toBe('World');
    });

    it('should handle already capitalized', () => {
      expect(capitalizeFirst('Hello')).toBe('Hello');
    });

    it('should handle empty string', () => {
      expect(capitalizeFirst('')).toBe('');
    });

    it('should handle single character', () => {
      expect(capitalizeFirst('a')).toBe('A');
    });
  });

  describe('formatPercentage', () => {
    it('should format percentage correctly', () => {
      expect(formatPercentage(0.75)).toBe('75.0%');
      expect(formatPercentage(0.5)).toBe('50.0%');
      expect(formatPercentage(1)).toBe('100.0%');
    });

    it('should handle decimal percentages', () => {
      expect(formatPercentage(0.123)).toBe('12.3%');
      expect(formatPercentage(0.001)).toBe('0.1%');
    });

    it('should handle zero', () => {
      expect(formatPercentage(0)).toBe('0.0%');
    });

    it('should handle values greater than 1', () => {
      expect(formatPercentage(1.5)).toBe('150.0%');
    });
  });

  describe('formatCurrency', () => {
    it('should format currency correctly', () => {
      // uk-UA: space (or NBSP) for thousands, comma for decimals
      expect(formatCurrency(1234.56)).toMatch(/^1[ \u00A0]234,56 \u20B4$/);
      expect(formatCurrency(1000000)).toMatch(/^1[ \u00A0]000[ \u00A0]000,00 \u20B4$/);
    });

    it('should handle zero', () => {
      expect(formatCurrency(0)).toBe('0,00 ₴');
    });

    it('should handle negative values', () => {
      expect(formatCurrency(-1234.56)).toMatch(/^-1[ \u00A0]234,56 \u20B4$/);
    });

    it('should handle custom currency', () => {
      expect(formatCurrency(1234.56, 'USD')).toMatch(/^1[ \u00A0]234,56 USD$/);
    });
  });
});

// Мок функции форматирования (замените на реальные импорты)
function formatDate(date: Date): string {
  if (isNaN(date.getTime())) return 'Н/Д';
  return date.toLocaleDateString('uk-UA');
}

function formatDateTime(date: Date): string {
  if (isNaN(date.getTime())) return 'Н/Д';
  return date.toLocaleString('uk-UA');
}

function formatNumber(num: number): string {
  return num.toLocaleString('uk-UA');
}

function formatFileSize(bytes: number): string {
  if (bytes === 0) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
}

function formatDuration(seconds: number): string {
  if (seconds === 0) return '0с';
  
  const days = Math.floor(seconds / 86400);
  const hours = Math.floor((seconds % 86400) / 3600);
  const minutes = Math.floor((seconds % 3600) / 60);
  const secs = seconds % 60;
  
  const parts = [];
  if (days > 0) parts.push(`${days}д`);
  if (hours > 0) parts.push(`${hours}год`);
  if (minutes > 0) parts.push(`${minutes}хв`);
  if (secs > 0 || parts.length === 0) parts.push(`${secs}с`);
  
  return parts.join(' ');
}

function truncateText(text: string | null | undefined, maxLength: number): string {
  if (!text) return '';
  if (text.length <= maxLength) return text;
  return text.substring(0, maxLength) + '...';
}

function capitalizeFirst(str: string): string {
  if (!str) return str;
  return str.charAt(0).toUpperCase() + str.slice(1);
}

function formatPercentage(value: number): string {
  return (value * 100).toFixed(1) + '%';
}

function formatCurrency(amount: number, currency: string = '₴'): string {
  return amount.toLocaleString('uk-UA', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }) + ' ' + currency;
} 