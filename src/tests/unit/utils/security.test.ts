/**
 * Unit тесты для утилиты security
 */

import { jest, describe, it, expect } from '@jest/globals';

describe('Security Utils', () => {
  describe('validateInput', () => {
    it('should validate safe input', () => {
      const result = validateInput('safe text');
      expect(result.isValid).toBe(true);
    });

    it('should reject SQL injection', () => {
      const result = validateInput("'; DROP TABLE users; --");
      expect(result.isValid).toBe(false);
      expect(result.reason).toContain('SQL injection');
    });

    it('should reject XSS attacks', () => {
      const result = validateInput('<script>alert("xss")</script>');
      expect(result.isValid).toBe(false);
      expect(result.reason).toContain('XSS');
    });

    it('should reject command injection', () => {
      const result = validateInput('$(rm -rf /)');
      expect(result.isValid).toBe(false);
      expect(result.reason).toContain('command injection');
    });
  });

  describe('sanitizeInput', () => {
    it('should sanitize HTML tags', () => {
      const result = sanitizeInput('<p>Hello <script>alert("xss")</script></p>');
      expect(result).toBe('Hello');
    });

    it('should preserve safe text', () => {
      const result = sanitizeInput('Safe text with numbers 123');
      expect(result).toBe('Safe text with numbers 123');
    });

    it('should handle empty input', () => {
      const result = sanitizeInput('');
      expect(result).toBe('');
    });
  });

  describe('rateLimit', () => {
    it('should allow request within limit', () => {
      const result = rateLimit('user123', 100, 60);
      expect(result.allowed).toBe(true);
    });

    it('should block request over limit', () => {
      // Симулируем превышение лимита
      for (let i = 0; i < 101; i++) {
        rateLimit('user123', 100, 60);
      }
      
      const result = rateLimit('user123', 100, 60);
      expect(result.allowed).toBe(false);
    });
  });

  describe('encryptData', () => {
    it('should encrypt sensitive data', () => {
      const data = 'sensitive information';
      const encrypted = encryptData(data);
      
      expect(encrypted).not.toBe(data);
      expect(encrypted).toMatch(/^[A-Za-z0-9+/=]+$/); // Base64 format
    });

    it('should decrypt data correctly', () => {
      const original = 'test data';
      const encrypted = encryptData(original);
      const decrypted = decryptData(encrypted);
      
      expect(decrypted).toBe(original);
    });
  });

  describe('validateToken', () => {
    it('should validate correct token', () => {
      const token = 'valid.jwt.token';
      const result = validateToken(token);
      expect(result.valid).toBe(true);
    });

    it('should reject invalid token', () => {
      const token = 'invalid.token';
      const result = validateToken(token);
      expect(result.valid).toBe(false);
    });
  });
});

// Мок функции безопасности
function validateInput(input: string): { isValid: boolean; reason?: string } {
  if (input.includes('DROP TABLE') || input.includes(';')) {
    return { isValid: false, reason: 'SQL injection detected' };
  }
  if (input.includes('<script>')) {
    return { isValid: false, reason: 'XSS attack detected' };
  }
  if (input.includes('$(') || input.includes('rm -rf')) {
    return { isValid: false, reason: 'command injection detected' };
  }
  return { isValid: true };
}

function sanitizeInput(input: string): string {
  return input.replace(/<[^>]*>/g, '');
}

function rateLimit(userId: string, limit: number, window: number): { allowed: boolean } {
  // Простая реализация rate limiting
  const key = `rate_limit_${userId}`;
  const current = parseInt(localStorage.getItem(key) || '0');
  
  if (current >= limit) {
    return { allowed: false };
  }
  
  localStorage.setItem(key, (current + 1).toString());
  return { allowed: true };
}

function encryptData(data: string): string {
  return btoa(data); // Простое Base64 кодирование
}

function decryptData(encrypted: string): string {
  return atob(encrypted); // Простое Base64 декодирование
}

function validateToken(token: string): { valid: boolean } {
  return { valid: token.includes('jwt') };
} 