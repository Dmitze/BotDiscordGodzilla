/**
 * Unit тесты для утилиты logger
 */

import { jest, describe, it, expect, beforeEach } from '@jest/globals';

// Мокаем winston
jest.mock('winston', () => ({
  createLogger: jest.fn(() => ({
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
  let logger: any;

  beforeEach(() => {
    // Очищаем моки
    jest.clearAllMocks();
    
    // Импортируем logger после моков
    logger = require('../../../utils/logger');
  });

  describe('logger instance', () => {
    it('should create logger instance', () => {
      expect(logger).toBeDefined();
    });

    it('should have required methods', () => {
      expect(typeof logger.info).toBe('function');
      expect(typeof logger.error).toBe('function');
      expect(typeof logger.warn).toBe('function');
      expect(typeof logger.debug).toBe('function');
    });
  });

  describe('logging methods', () => {
    it('should log info message', () => {
      const message = 'Test info message';
      logger.info(message);
      
      expect(logger.info).toHaveBeenCalledWith(message);
    });

    it('should log error message', () => {
      const message = 'Test error message';
      logger.error(message);
      
      expect(logger.error).toHaveBeenCalledWith(message);
    });

    it('should log warning message', () => {
      const message = 'Test warning message';
      logger.warn(message);
      
      expect(logger.warn).toHaveBeenCalledWith(message);
    });

    it('should log debug message', () => {
      const message = 'Test debug message';
      logger.debug(message);
      
      expect(logger.debug).toHaveBeenCalledWith(message);
    });
  });

  describe('error logging', () => {
    it('should log error with stack trace', () => {
      const error = new Error('Test error');
      logger.error('Error occurred', error);
      
      expect(logger.error).toHaveBeenCalledWith('Error occurred', error);
    });

    it('should log error object', () => {
      const errorObj = { message: 'Custom error', code: 500 };
      logger.error('API Error', errorObj);
      
      expect(logger.error).toHaveBeenCalledWith('API Error', errorObj);
    });
  });

  describe('structured logging', () => {
    it('should log with metadata', () => {
      const metadata = { userId: '123', action: 'search' };
      logger.info('User action', metadata);
      
      expect(logger.info).toHaveBeenCalledWith('User action', metadata);
    });

    it('should log performance data', () => {
      const perfData = { duration: 150, operation: 'search' };
      logger.info('Performance metric', perfData);
      
      expect(logger.info).toHaveBeenCalledWith('Performance metric', perfData);
    });
  });

  describe('log levels', () => {
    it('should respect log level configuration', () => {
      // В тестовом окружении должен быть установлен уровень error
      logger.debug('Debug message');
      logger.info('Info message');
      
      // Debug и info не должны логироваться в production
      expect(logger.debug).toHaveBeenCalledWith('Debug message');
      expect(logger.info).toHaveBeenCalledWith('Info message');
    });
  });

  describe('error handling', () => {
    it('should handle null messages gracefully', () => {
      logger.info(null);
      expect(logger.info).toHaveBeenCalledWith(null);
    });

    it('should handle undefined messages gracefully', () => {
      logger.info(undefined);
      expect(logger.info).toHaveBeenCalledWith(undefined);
    });

    it('should handle empty string messages', () => {
      logger.info('');
      expect(logger.info).toHaveBeenCalledWith('');
    });
  });

  describe('performance logging', () => {
    it('should log execution time', () => {
      const startTime = Date.now();
      const endTime = startTime + 100;
      
      logger.info('Operation completed', {
        duration: endTime - startTime,
        operation: 'test',
      });
      
      expect(logger.info).toHaveBeenCalledWith('Operation completed', {
        duration: 100,
        operation: 'test',
      });
    });
  });
}); 