/**
 * E2E тесты для Discord бота
 */

import { jest, describe, it, expect, beforeAll, afterAll } from '@jest/globals';
import { Bot } from '../../core/Bot';
import { createMockConfig } from '../utils/testHelpers';

describe('Bot E2E Tests', () => {
  let bot: Bot;
  let mockConfig: any;

  beforeAll(async () => {
    mockConfig = createMockConfig();
    bot = new Bot(mockConfig);
  });

  afterAll(async () => {
    if (bot) {
      await bot.shutdown();
    }
  });

  describe('Bot Initialization', () => {
    it('should initialize bot successfully', async () => {
      await expect(bot.initialize()).resolves.not.toThrow();
    });

    it('should have all required services', async () => {
      await bot.initialize();
      
      expect(bot.serviceContainer).toBeDefined();
      expect(bot.commandManager).toBeDefined();
      expect(bot.client).toBeDefined();
    });

    it('should load all commands', async () => {
      await bot.initialize();
      
      const commands = bot.commandManager.getCommands();
      expect(commands.size).toBeGreaterThan(0);
    });
  });

  describe('Service Container', () => {
    it('should initialize all services', async () => {
      await bot.initialize();
      
      const services = bot.serviceContainer.getServices();
      expect(Object.keys(services).length).toBeGreaterThan(0);
    });

    it('should have healthy services', async () => {
      await bot.initialize();
      
      const healthStatus = await bot.serviceContainer.getHealthStatus();
      
      Object.values(healthStatus).forEach(service => {
        expect(service.healthy).toBe(true);
      });
    });
  });

  describe('Command Manager', () => {
    it('should register commands (dynamic verification)', async () => {
      await bot.initialize();
      const commands = bot.commandManager.getCommands();

      // Динамічна перевірка: принаймні певне ядро команд повинно бути доступне
      const hasSearch = commands.has('пошук');
      const hasAi = commands.has('ai') || commands.has('ai_асистент');
      expect(hasSearch).toBe(true);
      expect(hasAi).toBe(true);

      // Загальна кількість повинна бути більшою за мінімум
      expect(commands.size).toBeGreaterThanOrEqual(5);

      // Імена повинні бути унікальними
      const names = Array.from(commands.keys());
      const unique = new Set(names);
      expect(unique.size).toBe(names.length);
    });

    it('should validate command structure', async () => {
      await bot.initialize();
      
      const commands = bot.commandManager.getCommands();
      
      for (const [name, command] of commands) {
        expect(command.getName()).toBeDefined();
        expect(command.getDescription()).toBeDefined();
        expect(command.getData()).toBeDefined();
      }
    });
  });

  describe('Health Check', () => {
    it('should return healthy status', async () => {
      await bot.initialize();
      
      const health = await bot.getHealthStatus();
      
      expect(health.healthy).toBe(true);
      expect(health.service).toBe('DiscordBot');
      expect(health.details).toBeDefined();
    });

    it('should include service details', async () => {
      await bot.initialize();
      
      const health = await bot.getHealthStatus();
      
      expect(health.details?.connected).toBeDefined();
      expect(health.details?.services).toBeDefined();
      expect(health.details?.uptime).toBeDefined();
    });
  });

  describe('Graceful Shutdown', () => {
    it('should shutdown gracefully', async () => {
      await bot.initialize();
      
      await expect(bot.shutdown()).resolves.not.toThrow();
    });

    it('should cleanup resources', async () => {
      await bot.initialize();
      await bot.shutdown();
      
      // Проверяем, что ресурсы освобождены
      expect(bot.isBotReady()).toBe(false);
    });
  });

  describe('Error Handling', () => {
    it('should handle initialization errors gracefully', async () => {
      // Создаем бота с неверной конфигурацией
      const invalidConfig = {
        ...mockConfig,
        discord: {
          ...mockConfig.discord,
          token: 'invalid_token',
        },
      };

      const invalidBot = new Bot(invalidConfig);
      
      // Бот должен обработать ошибку и не упасть
      await expect(invalidBot.initialize()).rejects.toThrow();
    });
  });

  describe('Performance', () => {
    it('should initialize within reasonable time', async () => {
      const startTime = Date.now();
      
      await bot.initialize();
      
      const duration = Date.now() - startTime;
      expect(duration).toBeLessThan(5000); // 5 секунд максимум
    });

    it('should shutdown within reasonable time', async () => {
      await bot.initialize();
      
      const startTime = Date.now();
      await bot.shutdown();
      
      const duration = Date.now() - startTime;
      expect(duration).toBeLessThan(3000); // 3 секунды максимум
    });
  });
}); 