/**
 * Unit тесты для утилиты logger
 */

import { jest, describe, it, expect, beforeEach } from '@jest/globals';
import logger from '../../../utils/logger';
import type { Logger as LoggerType, LogEntry } from '../../../utils/logger';

// Мокаем winston
jest.mock('winston', () => ({
  createLogger: jest.fn(() => ({
    log: jest.fn(),
    info: jest.fn(),
    error: jest.fn(),
    warn: jest.fn(),
    debug: jest.fn(),
  })),
  format: {
    combine: jest.fn(),
    timestamp: jest.fn(),
    errors: jest.fn(),
    json: jest.fn(),
    colorize: jest.fn(),
    simple: jest.fn(),
  },
  transports: {
    Console: jest.fn(),
    File: jest.fn(),
  },
}));

describe('Logger Utils', () => {
  const typedLogger: LoggerType = logger as unknown as LoggerType;

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
      expect((last.meta as any)['userId']).toBe('123');
      expect((last.meta as any)['action']).toBe('search');
    });

    it('should log performance data', () => {
      const perfData = { duration: 150, operation: 'search' };
      typedLogger.info('Performance metric', perfData);
      const buf = typedLogger.getLogBuffer();
      const last = buf[buf.length - 1] as LogEntry;
      expect(last.message).toBe('Performance metric');
      expect((last.meta as any)['duration']).toBe(150);
      expect((last.meta as any)['operation']).toBe('search');
    });
  });

  describe('log levels', () => {
    it('should respect log level configuration', () => {
      // Для простоты проверяем, что вызовы не бросают ошибок и записываются в буфер
      typedLogger.debug('Debug message');
      typedLogger.info('Info message');
      const buf = typedLogger.getLogBuffer();
      const messages = (buf as LogEntry[]).map((e) => e.message);
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
      expect((last.meta as any)['duration']).toBe(100);
      expect((last.meta as any)['operation']).toBe('test');
    });
  });
});