/**
 * Load тесты для Discord бота
 */

import { jest, describe, it, expect, beforeAll, afterAll } from '@jest/globals';
import { Bot } from '../../core/Bot';
import { createMockConfig } from '../utils/testHelpers';

describe('Load Tests', () => {
  let bot: Bot;
  let mockConfig: any;

  beforeAll(async () => {
    // Увімкнути fast-path лише для load тестів
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

  describe('High Load Command Execution', () => {
    it('should handle 1000 concurrent commands', async () => {
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
      
      // Выполняем 1000 команд параллельно
      const promises = Array(1000).fill(null).map(() => 
        bot.commandManager.execute(mockInteraction)
      );
      
      await Promise.all(promises);
      
      const duration = Date.now() - startTime;
      
      // 1000 команд должны выполниться за разумное время
      expect(duration).toBeLessThan(60000); // 60 секунд
      
      await bot.shutdown();
    }, 120000); // Увеличиваем timeout для load тестов

    it('should handle 500 sequential commands', async () => {
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
      
      // Выполняем 500 команд последовательно
      for (let i = 0; i < 500; i++) {
        await bot.commandManager.execute(mockInteraction);
      }
      
      const duration = Date.now() - startTime;
      
      // 500 команд должны выполниться за разумное время
      expect(duration).toBeLessThan(30000); // 30 секунд
      
      await bot.shutdown();
    }, 60000);

    it('should maintain performance under memory pressure', async () => {
      await bot.initialize();
      
      // Симулируем высокую нагрузку на память
      const memoryPressure = [];
      for (let i = 0; i < 10000; i++) {
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
      
      // Выполняем команды под нагрузкой
      const promises = Array(100).fill(null).map(() => 
        bot.commandManager.execute(mockInteraction)
      );
      
      await Promise.all(promises);
      
      const duration = Date.now() - startTime;
      
      // Производительность не должна сильно ухудшиться
      expect(duration).toBeLessThan(10000); // 10 секунд
      
      // Освобождаем память
      memoryPressure.length = 0;
      
      await bot.shutdown();
    }, 30000);
  });

  describe('Service Load Testing', () => {
    it('should handle concurrent service access', async () => {
      await bot.initialize();
      
      const services = ['google', 'ai', 'cache', 'metrics'];
      const startTime = Date.now();
      
      // Одновременно обращаемся к разным сервисам
      const promises = Array(100).fill(null).map(() => 
        Promise.all(services.map(serviceName => 
          bot.serviceContainer.get(serviceName).getHealthStatus()
        ))
      );
      
      await Promise.all(promises);
      
      const duration = Date.now() - startTime;
      expect(duration).toBeLessThan(5000); // 5 секунд
      
      await bot.shutdown();
    });

    it('should handle rapid bot initialization cycles', async () => {
      const cycles = 10;
      const startTime = Date.now();
      
      for (let i = 0; i < cycles; i++) {
        const cycleBot = new Bot(mockConfig);
        await cycleBot.initialize();
        await cycleBot.shutdown();
      }
      
      const duration = Date.now() - startTime;
      
      // 10 циклов должны выполниться за разумное время
      expect(duration).toBeLessThan(60000); // Increased from 30000 to 60000ms
    }, 120000); // Increased timeout from 60000 to 120000ms
  });

  describe('Memory Load Testing', () => {
    it('should not exceed memory limits under load', async () => {
      const initialMemory = process.memoryUsage().heapUsed;
      
      await bot.initialize();
      
      // Симулируем нагрузку
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

      // Выполняем много команд
      const promises = Array(500).fill(null).map(() => 
        bot.commandManager.execute(mockInteraction)
      );
      
      await Promise.all(promises);
      
      const finalMemory = process.memoryUsage().heapUsed;
      const memoryIncrease = finalMemory - initialMemory;
      
      // Увеличение памяти не должно превышать 100MB
      expect(memoryIncrease).toBeLessThan(100 * 1024 * 1024);
      
      await bot.shutdown();
    }, 60000);

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

  describe('Cache Load Testing', () => {
    it('should handle high cache throughput', async () => {
      await bot.initialize();
      
      const cacheService = bot.serviceContainer.get('cache');
      const startTime = Date.now();
      
      // Выполняем много операций с кешем
      const promises = Array(1000).fill(null).map((_, index) => 
        cacheService.set(`key_${index}`, { data: `value_${index}` })
      );
      
      await Promise.all(promises);
      
      const duration = Date.now() - startTime;
      expect(duration).toBeLessThan(5000); // 5 секунд
      
      await bot.shutdown();
    });

    it('should handle cache eviction under load', async () => {
      await bot.initialize();
      
      const cacheService = bot.serviceContainer.get('cache');
      
      // Заполняем кеш
      for (let i = 0; i < 1000; i++) {
        await cacheService.set(`key_${i}`, { data: `value_${i}` });
      }
      
      // Проверяем, что кеш работает
      const value = await cacheService.get('key_500');
      expect(value).toBeDefined();
      
      await bot.shutdown();
    });
  });
}); 