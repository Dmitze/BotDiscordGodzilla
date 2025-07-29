/**
 * Навантажувальні тести для Discord AI Assistant Bot
 * Оновлено: 28.07.2025
 */

const { jest } = require('@jest/globals');
const { performance } = require('perf_hooks');

// Мокаємо Discord.js для тестування
jest.mock('discord.js', () => ({
  SlashCommandBuilder: jest.fn().mockImplementation(() => ({
    setName: jest.fn().mockReturnThis(),
    setDescription: jest.fn().mockReturnThis(),
    addStringOption: jest.fn().mockReturnThis(),
    addIntegerOption: jest.fn().mockReturnThis(),
    addSubcommand: jest.fn().mockReturnThis(),
  })),
  EmbedBuilder: jest.fn().mockImplementation(() => ({
    setColor: jest.fn().mockReturnThis(),
    setTitle: jest.fn().mockReturnThis(),
    setDescription: jest.fn().mockReturnThis(),
    addFields: jest.fn().mockReturnThis(),
    setTimestamp: jest.fn().mockReturnThis(),
  })),
}));

// Мокаємо сервіси
jest.mock('../../services/AIService', () => {
  return jest.fn().mockImplementation(() => ({
    initialize: jest.fn().mockResolvedValue(),
    generateResponse: jest.fn().mockResolvedValue('AI response'),
    isActive: () => true,
  }));
});

jest.mock('../../services/GoogleService', () => {
  return jest.fn().mockImplementation(() => ({
    initialize: jest.fn().mockResolvedValue(),
    getSheetData: jest.fn().mockResolvedValue(
      Array(1000)
        .fill()
        .map((_, i) => [`Item ${i}`, `Description ${i}`, `Type ${i % 5}`])
    ),
    isActive: () => true,
  }));
});

jest.mock('../../services/CacheService', () => {
  return jest.fn().mockImplementation(() => ({
    initialize: jest.fn().mockResolvedValue(),
    get: jest.fn().mockResolvedValue(null),
    set: jest.fn().mockResolvedValue(),
    isActive: () => true,
  }));
});

jest.mock('../../utils/logger', () => ({
  info: jest.fn(),
  error: jest.fn(),
  warn: jest.fn(),
  debug: jest.fn(),
}));

class LoadTestRunner {
  constructor() {
    this.results = {
      totalRequests: 0,
      successfulRequests: 0,
      failedRequests: 0,
      totalResponseTime: 0,
      averageResponseTime: 0,
      minResponseTime: Infinity,
      maxResponseTime: 0,
      responseTimes: [],
      errors: [],
    };
  }

  /**
   * Генерація випадкового запиту
   */
  generateRandomQuery() {
    const queries = [
      'особовий склад',
      'техніка',
      'матеріали',
      'операції',
      'накази',
      'звіти',
      'плани',
      'інструкції',
    ];
    return queries[Math.floor(Math.random() * queries.length)];
  }

  /**
   * Створення мок interaction
   */
  createMockInteraction(query, commandType = 'search') {
    return {
      options: {
        getString: jest.fn().mockImplementation(param => {
          switch (param) {
            case 'запит':
              return query;
            case 'тип_документа':
              return 'all';
            case 'дата_від':
              return null;
            case 'дата_до':
              return null;
            case 'підрозділ':
              return null;
            case 'пріоритет':
              return 'all';
            case 'контекст':
              return null;
            case 'режим':
              return 'general';
            case 'дія':
              return 'search';
            default:
              return null;
          }
        }),
        getInteger: jest.fn().mockReturnValue(20),
        getSubcommand: jest.fn().mockReturnValue('особовий-склад'),
      },
      user: {
        tag: `testuser${Math.floor(Math.random() * 1000)}#1234`,
        id: Math.floor(Math.random() * 1000000).toString(),
      },
      deferReply: jest.fn().mockResolvedValue(),
      editReply: jest.fn().mockResolvedValue(),
      reply: jest.fn().mockResolvedValue(),
      deferred: false,
      replied: false,
    };
  }

  /**
   * Створення мок бота
   */
  createMockBot() {
    const AIService = require('../../services/AIService');
    const GoogleService = require('../../services/GoogleService');
    const CacheService = require('../../services/CacheService');

    return {
      getService: jest.fn(name => {
        const services = {
          ai: new AIService(),
          google: new GoogleService(),
          cache: new CacheService(),
        };
        return services[name];
      }),
      handleError: jest.fn().mockResolvedValue({
        handled: true,
        message: 'Error handled',
      }),
    };
  }

  /**
   * Виконання одного тесту
   */
  async executeSingleTest(command, interaction, bot) {
    const startTime = performance.now();

    try {
      await command.execute(interaction, bot);
      const endTime = performance.now();
      const responseTime = endTime - startTime;

      this.results.successfulRequests++;
      this.results.totalResponseTime += responseTime;
      this.results.minResponseTime = Math.min(this.results.minResponseTime, responseTime);
      this.results.maxResponseTime = Math.max(this.results.maxResponseTime, responseTime);
      this.results.responseTimes.push(responseTime);

      return { success: true, responseTime };
    } catch (error) {
      this.results.failedRequests++;
      this.results.errors.push({
        error: error.message,
        timestamp: new Date().toISOString(),
      });

      return { success: false, error: error.message };
    } finally {
      this.results.totalRequests++;
    }
  }

  /**
   * Виконання послідовних тестів
   */
  async runSequentialTests(numTests = 100) {
    console.log(`🚀 Запуск ${numTests} послідовних тестів...`);

    const bot = this.createMockBot();
    const searchCommand = require('../../commands/search');
    const aiCommand = require('../../commands/aiAssistant');
    const documentsCommand = require('../../commands/documents');

    const commands = [searchCommand, aiCommand, documentsCommand];

    for (let i = 0; i < numTests; i++) {
      const query = this.generateRandomQuery();
      const interaction = this.createMockInteraction(query);
      const command = commands[i % commands.length];

      await this.executeSingleTest(command, interaction, bot);

      // Прогрес кожні 10 тестів
      if ((i + 1) % 10 === 0) {
        console.log(`✅ Виконано ${i + 1}/${numTests} тестів`);
      }
    }

    this.calculateResults();
    return this.results;
  }

  /**
   * Виконання паралельних тестів
   */
  async runParallelTests(numTests = 100, concurrency = 10) {
    console.log(`🚀 Запуск ${numTests} паралельних тестів з конкурентністю ${concurrency}...`);

    const bot = this.createMockBot();
    const searchCommand = require('../../commands/search');
    const aiCommand = require('../../commands/aiAssistant');
    const documentsCommand = require('../../commands/documents');

    const commands = [searchCommand, aiCommand, documentsCommand];

    const batches = [];
    for (let i = 0; i < numTests; i += concurrency) {
      const batch = [];
      for (let j = 0; j < concurrency && i + j < numTests; j++) {
        const query = this.generateRandomQuery();
        const interaction = this.createMockInteraction(query);
        const command = commands[(i + j) % commands.length];

        batch.push(this.executeSingleTest(command, interaction, bot));
      }
      batches.push(batch);
    }

    for (let i = 0; i < batches.length; i++) {
      await Promise.all(batches[i]);

      // Прогрес кожні 10 батчів
      if ((i + 1) % 10 === 0) {
        console.log(`✅ Виконано ${(i + 1) * concurrency}/${numTests} тестів`);
      }
    }

    this.calculateResults();
    return this.results;
  }

  /**
   * Виконання стресового тесту
   */
  async runStressTest(duration = 60000, requestsPerSecond = 10) {
    console.log(
      `🚀 Запуск стресового тесту на ${duration / 1000} секунд з ${requestsPerSecond} запитів/сек...`
    );

    const bot = this.createMockBot();
    const searchCommand = require('../../commands/search');
    const aiCommand = require('../../commands/aiAssistant');
    const documentsCommand = require('../../commands/documents');

    const commands = [searchCommand, aiCommand, documentsCommand];
    const startTime = Date.now();
    const interval = 1000 / requestsPerSecond;

    return new Promise(resolve => {
      const intervalId = setInterval(async () => {
        if (Date.now() - startTime >= duration) {
          clearInterval(intervalId);
          this.calculateResults();
          resolve(this.results);
          return;
        }

        const query = this.generateRandomQuery();
        const interaction = this.createMockInteraction(query);
        const command = commands[Math.floor(Math.random() * commands.length)];

        await this.executeSingleTest(command, interaction, bot);
      }, interval);
    });
  }

  /**
   * Розрахунок результатів
   */
  calculateResults() {
    this.results.averageResponseTime =
      this.results.totalResponseTime / this.results.successfulRequests;

    // Розрахунок медіани
    const sortedTimes = [...this.results.responseTimes].sort((a, b) => a - b);
    this.results.medianResponseTime = sortedTimes[Math.floor(sortedTimes.length / 2)];

    // Розрахунок 95-го перцентиля
    this.results.p95ResponseTime = sortedTimes[Math.floor(sortedTimes.length * 0.95)];

    // Розрахунок 99-го перцентиля
    this.results.p99ResponseTime = sortedTimes[Math.floor(sortedTimes.length * 0.99)];

    // Розрахунок успішності
    this.results.successRate = (this.results.successfulRequests / this.results.totalRequests) * 100;
  }

  /**
   * Виведення результатів
   */
  printResults() {
    console.log('\n📊 РЕЗУЛЬТАТИ НАВАНТАЖУВАЛЬНИХ ТЕСТІВ');
    console.log('=====================================');
    console.log(`📈 Загальна кількість запитів: ${this.results.totalRequests}`);
    console.log(`✅ Успішних запитів: ${this.results.successfulRequests}`);
    console.log(`❌ Невдалих запитів: ${this.results.failedRequests}`);
    console.log(`📊 Успішність: ${this.results.successRate.toFixed(2)}%`);
    console.log('\n⏱️ ЧАС ВІДПОВІДІ:');
    console.log(`   Середній: ${this.results.averageResponseTime.toFixed(2)}ms`);
    console.log(`   Медіана: ${this.results.medianResponseTime.toFixed(2)}ms`);
    console.log(`   Мінімальний: ${this.results.minResponseTime.toFixed(2)}ms`);
    console.log(`   Максимальний: ${this.results.maxResponseTime.toFixed(2)}ms`);
    console.log(`   95-й перцентиль: ${this.results.p95ResponseTime.toFixed(2)}ms`);
    console.log(`   99-й перцентиль: ${this.results.p99ResponseTime.toFixed(2)}ms`);

    if (this.results.errors.length > 0) {
      console.log('\n❌ ПОМИЛКИ:');
      this.results.errors.slice(0, 5).forEach((error, index) => {
        console.log(`   ${index + 1}. ${error.error} (${error.timestamp})`);
      });
      if (this.results.errors.length > 5) {
        console.log(`   ... та ще ${this.results.errors.length - 5} помилок`);
      }
    }

    console.log('\n🎯 ВИСНОВКИ:');
    if (this.results.successRate >= 95) {
      console.log('   ✅ Система показує високу стабільність');
    } else if (this.results.successRate >= 90) {
      console.log('   ⚠️ Система показує задовільну стабільність');
    } else {
      console.log('   ❌ Система потребує покращення стабільності');
    }

    if (this.results.averageResponseTime <= 1000) {
      console.log('   ✅ Система показує високу продуктивність');
    } else if (this.results.averageResponseTime <= 3000) {
      console.log('   ⚠️ Система показує задовільну продуктивність');
    } else {
      console.log('   ❌ Система потребує оптимізації продуктивності');
    }
  }
}

// Тести
describe('Load Tests', () => {
  let loadTestRunner;

  beforeEach(() => {
    loadTestRunner = new LoadTestRunner();
  });

  test('Sequential Load Test - 50 requests', async () => {
    const results = await loadTestRunner.runSequentialTests(50);

    expect(results.totalRequests).toBe(50);
    expect(results.successRate).toBeGreaterThan(90);
    expect(results.averageResponseTime).toBeLessThan(5000);
  }, 30000);

  test('Parallel Load Test - 100 requests with concurrency 10', async () => {
    const results = await loadTestRunner.runParallelTests(100, 10);

    expect(results.totalRequests).toBe(100);
    expect(results.successRate).toBeGreaterThan(90);
    expect(results.averageResponseTime).toBeLessThan(5000);
  }, 60000);

  test('Stress Test - 30 seconds with 5 requests per second', async () => {
    const results = await loadTestRunner.runStressTest(30000, 5);

    expect(results.totalRequests).toBeGreaterThan(100);
    expect(results.successRate).toBeGreaterThan(85);
    expect(results.averageResponseTime).toBeLessThan(10000);
  }, 90000);

  test('Performance Benchmark', async () => {
    const results = await loadTestRunner.runSequentialTests(200);

    loadTestRunner.printResults();

    // Перевірка продуктивності
    expect(results.successRate).toBeGreaterThan(95);
    expect(results.averageResponseTime).toBeLessThan(2000);
    expect(results.p95ResponseTime).toBeLessThan(5000);
    expect(results.p99ResponseTime).toBeLessThan(10000);
  }, 120000);
});

// Експорт для використання в інших тестах
module.exports = LoadTestRunner;
