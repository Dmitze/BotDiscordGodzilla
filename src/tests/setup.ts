/* eslint-disable no-console */
// Setup файл для Jest тестів

import { config } from 'dotenv';
// ВАЖЛИВО: Мокаємо SecurityManager до імпорту тестованих модулів, щоб уникнути setInterval
jest.mock('../utils/security', () => {
  // Легковагові моки без таймерів/interval'ів
  const validateInput = (input: string) => ({
    isValid: true,
    sanitizedValue: input,
    errors: [],
    warnings: [],
  });
  const sanitizeInput = (input: string) => input;
  const checkRateLimit = (_userId: string) => ({
    allowed: true,
    remaining: 10,
    resetTime: Date.now() + 60_000,
  });
  const validateUrl = (url: string) => ({
    isValid: true,
    sanitizedValue: url,
    errors: [],
    warnings: [],
  });
  class SecurityManagerMock {
    public initialize(): void { /* no-op */ }
    public cleanup(): void { /* no-op */ }
    public validateInput = validateInput;
    public checkRateLimit = checkRateLimit;
    public validateUrl = validateUrl;
    public getStats = () => ({ totalValidations: 0 } as any);
    public getSuspiciousActivities = () => [] as any[];
  }
  const securityManager = new SecurityManagerMock();
  return {
    SecurityManager: SecurityManagerMock,
    securityManager,
    validateInput,
    checkRateLimit,
    validateUrl,
    getSecurityStats: () => securityManager.getStats(),
    getSuspiciousActivities: () => securityManager.getSuspiciousActivities(),
    cleanupSecurityManager: () => securityManager.cleanup(),
    sanitizeInput,
    default: securityManager,
  };
});

// Завантаження змінних середовища
config({ path: '.env.test' });

// Мок для process.env
process.env['NODE_ENV'] = 'test';

// Придушення консольних логів під час тестів (можна ввімкнути через TEST_VERBOSE_LOGS=true)
const VERBOSE = process.env['TEST_VERBOSE_LOGS'] === 'true';
if (!VERBOSE) {
  jest.spyOn(console, 'log').mockImplementation(() => {});
  jest.spyOn(console, 'info').mockImplementation(() => {});
  jest.spyOn(console, 'warn').mockImplementation(() => {});
  jest.spyOn(console, 'error').mockImplementation(() => {});
}

// Базові налаштування для тестів (лог лише у VERBOSE режимі)
if (VERBOSE) {
  console.log('🧪 Тестове середовище ініціалізовано');
}

// Глобальна очистка ресурсів логера після всіх тестів (динамічний імпорт щоб уникнути ранньої ініціалізації)
afterAll(async () => {
  try {
    const { default: logger } = await import('../utils/logger');
    await logger.cleanup();
  } catch (e) {
    // ignore
  }
});
