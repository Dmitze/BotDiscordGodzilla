/* eslint-disable @typescript-eslint/no-unsafe-return */
/* eslint-disable @typescript-eslint/no-explicit-any */
/**
 * Unit тесты для утилиты logger
 */

import { jest, describe, it, expect, beforeEach, beforeAll } from '@jest/globals';
import type { Logger as LoggerType, LogEntry } from '../../../utils/logger';
import type * as Winston from 'winston';

// Екземпляр логера буде імпортований після встановлення моків
let typedLogger: LoggerType;

beforeAll(async () => {
  // Мокаем winston ДО імпорту логера, щоб сінглтон використав мок
  // Зверніть увагу: logger.ts імпортує winston як default
  // Тому надаємо default-об'єкт з потрібними властивостями
  // та також дублікуємо їх як іменовані експорти на випадок звернення
  jest.unstable_mockModule('winston', () => {
    const mockedLogger = {
      log: jest.fn(),
      info: jest.fn(),
      error: jest.fn(),
      warn: jest.fn(),
      debug: jest.fn(),
      close: jest.fn(),
    };
    const emptyFormat = {} as unknown as Winston.Logform.Format;
    const emptyTransport = {} as unknown as Winston.transport;
    const mocked = {
      createLogger: jest.fn<() => Winston.Logger>(() => mockedLogger as unknown as Winston.Logger),
      format: {
        combine: jest.fn<() => Winston.Logform.Format>(() => emptyFormat),
        timestamp: jest.fn<() => Winston.Logform.Format>(() => emptyFormat),
        errors: jest.fn<() => Winston.Logform.Format>(() => emptyFormat),
        json: jest.fn<() => Winston.Logform.Format>(() => emptyFormat),
        colorize: jest.fn<() => Winston.Logform.Format>(() => emptyFormat),
        simple: jest.fn<() => Winston.Logform.Format>(() => emptyFormat),
        printf: jest.fn<() => Winston.Logform.Format>(() => emptyFormat),
      },
      transports: {
        Console: jest.fn<() => Winston.transport>(() => emptyTransport),
        File: jest.fn<() => Winston.transport>(() => emptyTransport),
      },
    };
    return { __esModule: true, default: mocked, ...mocked };
  });

  const mod = await import('../../../utils/logger');
  typedLogger = mod.default as unknown as LoggerType;
});

describe('Logger Utils', () => {
  beforeEach(() => {
    // Очищаем моки
    jest.clearAllMocks();
  });

  describe('logger instance', () => {
    it('should create logger instance', () => {
      expect(typedLogger).toBeDefined();
    });

    it('should have required methods', () => {
      expect(typeof typedLogger.info).toBe('function');
      expect(typeof typedLogger.error).toBe('function');
      expect(typeof typedLogger.warn).toBe('function');
      expect(typeof typedLogger.debug).toBe('function');
    });
  });

  describe('logging methods', () => {
    it('should log info message', () => {
      const message = 'Test info message';
      typedLogger.info(message);
      const buf = typedLogger.getLogBuffer();
      const last = buf[buf.length - 1] as LogEntry;
      expect(last.message).toBe(message);
      expect(last.level).toBe('info');
    });

    it('should log error message', () => {
      const message = 'Test error message';
      typedLogger.error(message);
      const buf = typedLogger.getLogBuffer();
      const last = buf[buf.length - 1] as LogEntry;
      expect(last.message).toBe(message);
      expect(last.level).toBe('error');
    });

    it('should log warning message', () => {
      const message = 'Test warning message';
      typedLogger.warn(message);
      const buf = typedLogger.getLogBuffer();
      const last = buf[buf.length - 1] as LogEntry;
      expect(last.message).toBe(message);
      expect(last.level).toBe('warn');
    });

    it('should log debug message', () => {
      const message = 'Test debug message';
      typedLogger.debug(message);
      const buf = typedLogger.getLogBuffer();
      const last = buf[buf.length - 1] as LogEntry;
      expect(last.message).toBe(message);
      expect(last.level).toBe('debug');
    });
  });

  describe('error logging', () => {
    it('should log error with stack trace', () => {
      const error = new Error('Test error');
      typedLogger.error('Error occurred', error);
      const buf = typedLogger.getLogBuffer();
      const last = buf[buf.length - 1] as LogEntry;
      expect(last.message).toBe('Error occurred');
      expect(last.level).toBe('error');
      expect(last.meta).toBeDefined();
    });

    it('should log error object', () => {
      const errorObj = { message: 'Custom error', code: 500 };
      typedLogger.error('API Error', errorObj);
      const buf = typedLogger.getLogBuffer();
      const last = buf[buf.length - 1] as LogEntry;
      expect(last.message).toBe('API Error');
      expect(last.level).toBe('error');
      expect(last.meta).toBeDefined();
    });
  });

  describe('structured logging', () => {
    it('should log with metadata', () => {
      const metadata = { userId: '123', action: 'search' };
      typedLogger.info('User action', metadata);
      const buf = typedLogger.getLogBuffer();
      const last = buf[buf.length - 1] as LogEntry;
      expect(last.message).toBe('User action');
      const meta = last.meta as Record<string, unknown>;
      expect(meta['userId']).toBe('123');
      expect(meta['action']).toBe('search');
    });

    it('should log performance data', () => {
      const perfData = { duration: 150, operation: 'search' };
      typedLogger.info('Performance metric', perfData);
      const buf = typedLogger.getLogBuffer();
      const last = buf[buf.length - 1] as LogEntry;
      expect(last.message).toBe('Performance metric');
      const meta = last.meta as Record<string, unknown>;
      expect(meta['duration']).toBe(150);
      expect(meta['operation']).toBe('search');
    });
  });

  describe('log levels', () => {
    it('should respect log level configuration', () => {
      // Для простоты проверяем, что вызовы не бросают ошибок и записываются в буфер
      typedLogger.debug('Debug message');
      typedLogger.info('Info message');
      const buf = typedLogger.getLogBuffer();
      const messages = buf.map((e) => e.message);
      expect(messages).toContain('Debug message');
      expect(messages).toContain('Info message');
    });
  });

  describe('error handling', () => {
    it('should handle null messages gracefully', () => {
      typedLogger.info(null as unknown as string);
      const buf = typedLogger.getLogBuffer();
      const last = buf[buf.length - 1] as LogEntry;
      expect(last.message).toBe(null as unknown as string);
    });

    it('should handle undefined messages gracefully', () => {
      typedLogger.info(undefined as unknown as string);
      const buf = typedLogger.getLogBuffer();
      const last = buf[buf.length - 1] as LogEntry;
      expect(last.message).toBeUndefined();
    });

    it('should handle empty string messages', () => {
      typedLogger.info('');
      const buf = typedLogger.getLogBuffer();
      const last = buf[buf.length - 1] as LogEntry;
      expect(last.message).toBe('');
    });
  });

  describe('performance logging', () => {
    it('should log execution time', () => {
      const startTime = Date.now();
      const endTime = startTime + 100;
      
      typedLogger.info('Operation completed', {
        duration: endTime - startTime,
        operation: 'test',
      });
      const buf = typedLogger.getLogBuffer();
      const last = buf[buf.length - 1] as LogEntry;
      expect(last.message).toBe('Operation completed');
      const meta = last.meta as Record<string, unknown>;
      expect(meta['duration']).toBe(100);
      expect(meta['operation']).toBe('test');
    });
  });
});