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
  const sanitizeInput = (input: string, inputType?: 'command' | 'message' | 'url' | 'file') => {
    if (inputType) {
      return {
        isValid: true,
        sanitizedValue: input,
        errors: [],
        warnings: [],
      };
    }
    return input;
  };
  // Простий, чистий хелпер для маскування PII у тестах
  // Не створює таймерів, сумісний з юніт/інтеграційними тестами
  const maskPII = (
    input: string,
    opts?: { email?: boolean; phone?: boolean }
  ): string => {
    if (!input) return input;
    let out = input;
    const enableEmail = opts?.email !== false; // default true
    const enablePhone = opts?.phone !== false; // default true
    if (enableEmail) {
      // Маскуємо email: залишаємо перший символ локальної частини, решту замінюємо на * до домену
      const emailRegex = /([a-zA-Z0-9._%+-])([a-zA-Z0-9._%+-]*)(@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/g;
      out = out.replace(emailRegex, (_m, first: string, middle: string, domain: string) => {
        const maskedMiddle = middle.length > 0 ? '*'.repeat(Math.min(middle.length, 6)) : '***';
        return `${first}${maskedMiddle}${domain}`;
      });
    }
    if (enablePhone) {
      // Маскуємо телефони: залишаємо останні 4 цифри, решту замінюємо на *
      const phoneRegex = /(?<!\d)([+]?\d[\d\s().-]{6,}\d)(?!\d)/g;
      out = out.replace(phoneRegex, (match: string) => {
        const digits = match.replace(/\D/g, '');
        if (digits.length < 7) return match;
        return '*'.repeat(Math.max(0, digits.length - 4)) + digits.slice(-4);
      });
    }
    return out;
  };
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
    maskPII,
    default: securityManager,
  };
});

// Глобальний мок Redis: блокує реальні підключення та прибирає шумні логи у всіх тестах
jest.mock('redis', () => {
  const noop = async () => undefined;
  const syncNoop = () => undefined;
  return {
    createClient: jest.fn(() => ({
      on: syncNoop,
      off: syncNoop,
      connect: noop,
      disconnect: noop,
      get: jest.fn(async () => null),
      set: jest.fn(async () => 'OK'),
      del: jest.fn(async () => 1),
      exists: jest.fn(async () => 0),
      keys: jest.fn(async () => []),
      flushDb: jest.fn(async () => 'OK'),
      ping: jest.fn(async () => 'PONG'),
    })),
  };
});

// Умовне приглушення модулю логера у тестах (використовуємо реальний логер лише у VERBOSE режимі)
jest.mock('../utils/logger', () => {
  const VERBOSE = process.env['TEST_VERBOSE_LOGS'] === 'true';
  if (VERBOSE) {
    // Повертаємо реальний модуль, щоб бачити повні логи, коли це потрібно
    // eslint-disable-next-line @typescript-eslint/no-var-requires
    return jest.requireActual('../utils/logger');
  }
  const buffer: Array<{ level: string; message: unknown; meta?: Record<string, unknown> }> = [];
  const SECRET_KEYS = new Set([
    'token', 'apikey', 'api_key', 'password', 'pass', 'secret', 'clientsecret',
    'authorization', 'auth', 'bearer', 'session', 'cookie', 'cookies'
  ]);
  const redact = (key: string, value: unknown, seen: WeakSet<object>): unknown => {
    if (value == null) return value as unknown;
    if (SECRET_KEYS.has(key.toLowerCase())) return '[REDACTED]';
    if (typeof value === 'string') return value;
    if (typeof value === 'object') {
      if (seen.has(value as object)) return '[CIRCULAR]';
      seen.add(value as object);
      if (Array.isArray(value)) return value.map(v => redact(key, v, seen));
      const out: Record<string, unknown> = {};
      for (const [k, v] of Object.entries(value as Record<string, unknown>)) {
        out[k] = redact(k, v, seen);
      }
      return out;
    }
    return value;
  };
  const sanitizeMeta = (meta?: Record<string, unknown>) => {
    if (!meta) return meta;
    const seen = new WeakSet<object>();
    const out: Record<string, unknown> = {};
    for (const [k, v] of Object.entries(meta)) out[k] = redact(k, v, seen);
    return out;
  };
  const push = (level: string, message?: unknown, meta?: Record<string, unknown>) => {
    buffer.push({ level, message, meta: sanitizeMeta(meta) });
  };
  const makeFn = (level: string) => jest.fn((message?: unknown, meta?: Record<string, unknown>) => push(level, message, meta));
  const info = makeFn('info');
  const warn = makeFn('warn');
  const error = makeFn('error');
  const debug = makeFn('debug');
  const security = makeFn('security');
  const performance = makeFn('performance');
  const commands = makeFn('commands');
  const api = makeFn('api');
  const system = makeFn('system');
  const getLogBuffer = () => buffer;
  const clearLogBuffer = () => { buffer.length = 0; };
  const setLogLevel = jest.fn((_level: string) => {});
  const startTimer = jest.fn(() => ({ start: Date.now() }));
  const endTimer = jest.fn((_ctx?: Record<string, unknown>) => {});
  const withTimer = jest.fn(async <T>(op: string, fn: () => Promise<T> | T) => {
    startTimer();
    const res = await fn();
    endTimer({ operation: op });
    return res;
  });
  const child = jest.fn(() => mockLogger);
  const toJSON = jest.fn(() => ({ size: buffer.length }));
  const cleanup = jest.fn(async () => undefined);
  const mockLogger = {
    info,
    warn,
    error,
    debug,
    security,
    performance,
    commands,
    api,
    system,
    getLogBuffer,
    clearLogBuffer,
    setLogLevel,
    startTimer,
    endTimer,
    withTimer,
    child,
    toJSON,
    cleanup,
  };
  return { __esModule: true, default: mockLogger, mockLogger, ...mockLogger };
});

// Завантаження змінних середовища
config({ path: '.env.test' });

// Мок для process.env
process.env['NODE_ENV'] = 'test';
// Відключити фонові таймери/метрики; AI fast-path-и виставляються точково у перформанс/лоад тестах
process.env['DISABLE_CRON'] = process.env['DISABLE_CRON'] ?? 'true';
process.env['DISABLE_AI_TIMERS'] = process.env['DISABLE_AI_TIMERS'] ?? 'true';
process.env['DISABLE_AI_HEALTHCHECK'] = process.env['DISABLE_AI_HEALTHCHECK'] ?? 'true';
process.env['METRICS_ENABLE'] = process.env['METRICS_ENABLE'] ?? '0';
process.env['EMBEDDINGS_ENABLE'] = process.env['EMBEDDINGS_ENABLE'] ?? '0';

// ТЕСТОВИЙ ШИМ: автоматично викликати .unref() для таймерів Node, щоб уникнути open handle leaks у Jest
// Має бути якнайраніше після конфігурації env
{
  const _setTimeout = global.setTimeout.bind(global);
  const _setInterval = global.setInterval.bind(global);
  const _setImmediate = global.setImmediate.bind(global);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const tryUnref = (t: any) => { try { if (t && typeof t.unref === 'function') t.unref(); } catch { /* ignore */ } return t; };
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (global as any).setTimeout = ((handler: TimerHandler, timeout?: number, ...args: unknown[]) => tryUnref(_setTimeout(handler as any, timeout as any, ...args as any))) as typeof setTimeout;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (global as any).setInterval = ((handler: TimerHandler, timeout?: number, ...args: unknown[]) => tryUnref(_setInterval(handler as any, timeout as any, ...args as any))) as typeof setInterval;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (global as any).setImmediate = ((handler: (...args: unknown[]) => void, ...args: unknown[]) => tryUnref(_setImmediate(handler as any, ...args as any))) as typeof setImmediate;
}

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
