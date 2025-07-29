/**
 * Базовий клас для всіх сервісів
 * Забезпечує єдиний інтерфейс та базову функціональність
 */

const logger = require('../utils/logger');

class BaseService {
  constructor(name, config = {}) {
    this.name = name;
    this.config = config;
    this.isInitialized = false;
    this.isHealthy = true;
    this.stats = {
      startTime: null,
      totalRequests: 0,
      successfulRequests: 0,
      failedRequests: 0,
      averageResponseTime: 0,
      totalResponseTime: 0,
      lastError: null,
      lastErrorTime: null,
    };
  }

  /**
   * Ініціалізація сервісу
   */
  async initialize() {
    try {
      logger.info(`🔧 Ініціалізація сервісу ${this.name}...`);
      
      this.stats.startTime = new Date();
      await this.onInitialize();
      
      this.isInitialized = true;
      logger.info(`✅ Сервіс ${this.name} ініціалізовано`);
    } catch (error) {
      logger.error(`❌ Помилка ініціалізації сервісу ${this.name}:`, error);
      this.isHealthy = false;
      this.stats.lastError = error.message;
      this.stats.lastErrorTime = new Date();
      throw error;
    }
  }

  /**
   * Перевірка стану сервісу
   */
  isHealthy() {
    return this.isHealthy && this.isInitialized;
  }

  /**
   * Отримання статистики сервісу
   */
  getStats() {
    const uptime = this.stats.startTime 
      ? Date.now() - this.stats.startTime.getTime()
      : 0;

    return {
      name: this.name,
      isInitialized: this.isInitialized,
      isHealthy: this.isHealthy,
      uptime,
      ...this.stats,
      successRate: this.stats.totalRequests > 0 
        ? (this.stats.successfulRequests / this.stats.totalRequests) * 100 
        : 0,
    };
  }

  /**
   * Оновлення статистики
   */
  updateStats(success, responseTime = 0) {
    this.stats.totalRequests++;
    this.stats.totalResponseTime += responseTime;
    
    if (success) {
      this.stats.successfulRequests++;
    } else {
      this.stats.failedRequests++;
    }

    // Розрахунок середнього часу відповіді
    this.stats.averageResponseTime = this.stats.totalResponseTime / this.stats.totalRequests;
  }

  /**
   * Обробка помилки
   */
  handleError(error, context = {}) {
    this.isHealthy = false;
    this.stats.lastError = error.message;
    this.stats.lastErrorTime = new Date();
    
    logger.error(`❌ Помилка в сервісі ${this.name}:`, {
      error: error.message,
      context,
      service: this.name,
    });

    return {
      success: false,
      error: error.message,
      service: this.name,
      timestamp: new Date(),
    };
  }

  /**
   * Скидання стану помилки
   */
  resetError() {
    this.isHealthy = true;
    this.stats.lastError = null;
    this.stats.lastErrorTime = null;
  }

  /**
   * Завершення роботи сервісу
   */
  async shutdown() {
    try {
      logger.info(`🛑 Завершення роботи сервісу ${this.name}...`);
      
      await this.onShutdown();
      
      this.isInitialized = false;
      this.isHealthy = false;
      
      logger.info(`✅ Сервіс ${this.name} завершено`);
    } catch (error) {
      logger.error(`❌ Помилка завершення сервісу ${this.name}:`, error);
      throw error;
    }
  }

  /**
   * Перезапуск сервісу
   */
  async restart() {
    try {
      logger.info(`🔄 Перезапуск сервісу ${this.name}...`);
      
      await this.shutdown();
      await this.initialize();
      
      logger.info(`✅ Сервіс ${this.name} перезапущено`);
    } catch (error) {
      logger.error(`❌ Помилка перезапуску сервісу ${this.name}:`, error);
      throw error;
    }
  }

  /**
   * Health check
   */
  async healthCheck() {
    try {
      const result = await this.onHealthCheck();
      this.isHealthy = result.healthy;
      
      if (!result.healthy) {
        this.stats.lastError = result.error || 'Health check failed';
        this.stats.lastErrorTime = new Date();
      } else {
        this.resetError();
      }
      
      return result;
    } catch (error) {
      this.handleError(error, { context: 'healthCheck' });
      return {
        healthy: false,
        error: error.message,
        service: this.name,
      };
    }
  }

  /**
   * Віртуальні методи для перевизначення в нащадках
   */

  /**
   * Ініціалізація сервісу (перевизначається в нащадках)
   */
  async onInitialize() {
    // Базова реалізація - нічого не робить
  }

  /**
   * Завершення роботи сервісу (перевизначається в нащадках)
   */
  async onShutdown() {
    // Базова реалізація - нічого не робить
  }

  /**
   * Health check сервісу (перевизначається в нащадках)
   */
  async onHealthCheck() {
    // Базова реалізація - повертає успішний результат
    return {
      healthy: this.isInitialized,
      service: this.name,
      timestamp: new Date(),
    };
  }

  /**
   * Валідація конфігурації (перевизначається в нащадках)
   */
  validateConfig() {
    // Базова реалізація - завжди успішна
    return { valid: true };
  }

  /**
   * Отримання інформації про сервіс
   */
  getInfo() {
    return {
      name: this.name,
      version: this.config.version || '1.0.0',
      description: this.config.description || 'Base service',
      isInitialized: this.isInitialized,
      isHealthy: this.isHealthy,
      uptime: this.stats.startTime 
        ? Date.now() - this.stats.startTime.getTime()
        : 0,
    };
  }
}

module.exports = BaseService; 