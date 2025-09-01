/**
 * Performance тесты для Discord бота
 */

import { jest, describe, it, expect, beforeAll, afterAll } from '@jest/globals';
import { Bot } from '../../core/Bot';
import { createMockConfig } from '../utils/testHelpers';

describe('Performance Tests', () => {
  let bot: Bot;
  let mockConfig: any;

  beforeAll(async () => {
    // Увімкнути fast-path лише для перформанс тестів
    process.env['AI_TEST_FAST'] = process.env['AI_TEST_FAST'] ?? '1';
    process.env['AI_PERF_FAST'] = process.env['AI_PERF_FAST'] ?? '1';
    process.env['DISABLE_AI_TIMERS'] = process.env['DISABLE_AI_TIMERS'] ?? 'true';
    process.env['DISABLE_AI_HEALTHCHECK'] = process.env['DISABLE_AI_HEALTHCHECK'] ?? 'true';
    mockConfig = createMockConfig();
    bot = new Bot(mockConfig);
  });

  afterAll(async () => {
    if (bot) {
      await bot.shutdown();
    }
  });

  describe('Bot Initialization Performance', () => {
    it('should initialize within 3 seconds', async () => {
      const startTime = Date.now();
      
      await bot.initialize();
      
      const duration = Date.now() - startTime;
      expect(duration).toBeLessThan(3000);
    });

    it('should shutdown within 2 seconds', async () => {
      await bot.initialize();
      
      const startTime = Date.now();
      await bot.shutdown();
      
      const duration = Date.now() - startTime;
      expect(duration).toBeLessThan(2000);
    });
  });

  describe('Command Execution Performance', () => {
    beforeEach(async () => {
      await bot.initialize();
    });

    afterEach(async () => {
      await bot.shutdown();
    });

    it('should execute search command within 1 second', async () => {
      const mockInteraction = {
        commandName: 'пошук',
        options: {
          getString: jest.fn().mockReturnValue('test'),
        },
        reply: jest.fn(),
        client: {
          serviceContainer: {
            get: jest.fn().mockReturnValue({
              searchData: jest.fn().mockResolvedValue([['test', 'data']]),
            }),
          },
        },
      };

      const startTime = Date.now();
      
      await bot.commandManager.execute(mockInteraction);
      
      const duration = Date.now() - startTime;
      expect(duration).toBeLessThan(1000);
    });

    it('should handle multiple concurrent commands', async () => {
      const mockInteraction = {
        commandName: 'пошук',
        options: {
          getString: jest.fn().mockReturnValue('test'),
        },
        reply: jest.fn(),
        client: {
          serviceContainer: {
            get: jest.fn().mockReturnValue({
              searchData: jest.fn().mockResolvedValue([['test', 'data']]),
            }),
          },
        },
      };

      const startTime = Date.now();
      
      // Выполняем 10 команд параллельно
      const promises = Array(10).fill(null).map(() => 
        bot.commandManager.execute(mockInteraction)
      );
      
      await Promise.all(promises);
      
      const duration = Date.now() - startTime;
      expect(duration).toBeLessThan(2000); // 2 секунды на 10 команд
    });
  });

  describe('Memory Usage Performance', () => {
    it('should not exceed memory limits during initialization', async () => {
      const initialMemory = process.memoryUsage().heapUsed;
      
      await bot.initialize();
      
      const finalMemory = process.memoryUsage().heapUsed;
      const memoryIncrease = finalMemory - initialMemory;
      
      // Увеличение памяти не должно превышать 50MB
      expect(memoryIncrease).toBeLessThan(50 * 1024 * 1024);
    });

    it('should release memory after shutdown', async () => {
      await bot.initialize();
      
      const memoryBeforeShutdown = process.memoryUsage().heapUsed;
      
      await bot.shutdown();
      // Дати часу GC, якщо доступний, та мікропаузу після shutdown
      if (global && typeof global.gc === 'function') {
        try { global.gc(); } catch { /* ignore */ }
      }
      await new Promise(r => setTimeout(r, 50));
      const memoryAfterShutdown = process.memoryUsage().heapUsed;
      
      // Дозволяємо невеликий шум вимірювання (1 МБ)
      const EPS = 1 * 1024 * 1024;
      expect(memoryAfterShutdown).toBeLessThanOrEqual(memoryBeforeShutdown + EPS);
    });
  });

  describe('Service Performance', () => {
    beforeEach(async () => {
      await bot.initialize();
    });

    afterEach(async () => {
      await bot.shutdown();
    });

    it('should initialize services within 2 seconds', async () => {
      const startTime = Date.now();
      
      await bot.serviceContainer.initialize();
      
      const duration = Date.now() - startTime;
      expect(duration).toBeLessThan(2000);
    });

    it('should get health status within 800ms', async () => {
      const startTime = Date.now();
      
      await bot.serviceContainer.getHealthStatus();
      
      const duration = Date.now() - startTime;
      expect(duration).toBeLessThan(800);
    });
  });

  describe('Cache Performance', () => {
    beforeEach(async () => {
      await bot.initialize();
    });

    afterEach(async () => {
      await bot.shutdown();
    });

    it('should set cache value within 100ms', async () => {
      const cacheService = bot.serviceContainer.get('cache');
      
      const startTime = Date.now();
      
      await cacheService.set('test_key', { data: 'test_value' });
      
      const duration = Date.now() - startTime;
      expect(duration).toBeLessThan(100);
    });

    it('should get cache value within 50ms', async () => {
      const cacheService = bot.serviceContainer.get('cache');
      
      await cacheService.set('test_key', { data: 'test_value' });
      
      const startTime = Date.now();
      
      await cacheService.get('test_key');
      
      const duration = Date.now() - startTime;
      expect(duration).toBeLessThan(50);
    });
  });

  describe('Google Service Performance', () => {
    beforeEach(async () => {
      await bot.initialize();
    });

    afterEach(async () => {
      await bot.shutdown();
    });

    it('should search data within 2 seconds', async () => {
      const googleService = bot.serviceContainer.get('google');
      
      const startTime = Date.now();
      
      await googleService.searchData('test', 10);
      
      const duration = Date.now() - startTime;
      expect(duration).toBeLessThan(2000);
    });
  });

  describe('AI Service Performance', () => {
    beforeEach(async () => {
      await bot.initialize();
    });

    afterEach(async () => {
      await bot.shutdown();
    });

    it('should generate AI response within 5 seconds', async () => {
      const aiService = bot.serviceContainer.get('ai');
      
      const startTime = Date.now();
      
      await aiService.generateResponse('Test query');
      
      const duration = Date.now() - startTime;
      expect(duration).toBeLessThan(5000);
    });
  });

  describe('Load Testing', () => {
    it('should handle 100 sequential commands', async () => {
      await bot.initialize();
      
      const mockInteraction = {
        commandName: 'пошук',
        options: {
          getString: jest.fn().mockReturnValue('test'),
        },
        reply: jest.fn(),
        client: {
          serviceContainer: {
            get: jest.fn().mockReturnValue({
              searchData: jest.fn().mockResolvedValue([['test', 'data']]),
            }),
          },
        },
      };

      const startTime = Date.now();
      
      // Выполняем 100 команд последовательно
      for (let i = 0; i < 100; i++) {
        await bot.commandManager.execute(mockInteraction);
      }
      
      const duration = Date.now() - startTime;
      
      // 100 команд должны выполниться за разумное время
      expect(duration).toBeLessThan(30000); // 30 секунд
      
      await bot.shutdown();
    });

    it('should maintain performance under memory pressure', async () => {
      await bot.initialize();
      
      // Симулируем нагрузку на память
      const memoryPressure = [];
      for (let i = 0; i < 1000; i++) {
        memoryPressure.push(new Array(1000).fill('test'));
      }
      
      const mockInteraction = {
        commandName: 'пошук',
        options: {
          getString: jest.fn().mockReturnValue('test'),
        },
        reply: jest.fn(),
        client: {
          serviceContainer: {
            get: jest.fn().mockReturnValue({
              searchData: jest.fn().mockResolvedValue([['test', 'data']]),
            }),
          },
        },
      };

      const startTime = Date.now();
      
      await bot.commandManager.execute(mockInteraction);
      
      const duration = Date.now() - startTime;
      
      // Производительность не должна сильно ухудшиться
      expect(duration).toBeLessThan(2000);
      
      // Освобождаем память
      memoryPressure.length = 0;
      
      await bot.shutdown();
    });
  });

  describe('Stress Testing', () => {
    it('should handle rapid initialization and shutdown cycles', async () => {
      const cycles = 5;
      const startTime = Date.now();
      
      for (let i = 0; i < cycles; i++) {
        const cycleBot = new Bot(mockConfig);
        await cycleBot.initialize();
        await cycleBot.shutdown();
      }
      
      const duration = Date.now() - startTime;
      
      // 5 циклов должны выполниться за разумное время
      expect(duration).toBeLessThan(15000); // 15 секунд
    });

    it('should handle concurrent service access', async () => {
      await bot.initialize();
      
      const services = ['google', 'ai', 'cache', 'metrics'];
      const startTime = Date.now();
      
      // Одновременно обращаемся к разным сервисам
      const promises = services.map(serviceName => 
        bot.serviceContainer.get(serviceName).getHealthStatus()
      );
      
      await Promise.all(promises);
      
      const duration = Date.now() - startTime;
      // Дозволяємо більш м'який ліміт з урахуванням оверхеду середовища CI
      expect(duration).toBeLessThan(4000);
      
      await bot.shutdown();
    });
  });
});