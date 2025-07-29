/**
 * Базовий клас для всіх сервісів
 * Надає спільну функціональність та інтерфейс
 */

import type { BaseService as IBaseService, BotConfig, HealthStatus, ServiceStats } from '@/types';

export abstract class BaseService implements IBaseService {
  public readonly name: string;
  public readonly config: BotConfig;
  protected isInitialized = false;
  protected startTime: number;

  constructor(name: string, config: BotConfig) {
    this.name = name;
    this.config = config;
    this.startTime = Date.now();
  }

  /**
   * Ініціалізація сервісу
   */
  public async initialize(): Promise<void> {
    if (this.isInitialized) {
      throw new Error(`Сервіс ${this.name} вже ініціалізовано`);
    }

    try {
      await this.onInitialize();
      this.isInitialized = true;
    } catch (error) {
      throw new Error(`Помилка ініціалізації сервісу ${this.name}: ${error}`);
    }
  }

  /**
   * Завершення роботи сервісу
   */
  public async shutdown(): Promise<void> {
    if (!this.isInitialized) {
      return;
    }

    try {
      await this.onShutdown();
      this.isInitialized = false;
    } catch (error) {
      throw new Error(`Помилка завершення сервісу ${this.name}: ${error}`);
    }
  }

  /**
   * Перевірка здоров'я сервісу
   */
  public async healthCheck(): Promise<HealthStatus> {
    try {
      const health = await this.onHealthCheck();
      return {
        healthy: health.healthy,
        service: this.name,
        ...(health.error && { error: health.error }),
        ...(health.details && { details: health.details }),
      };
    } catch (error) {
      return {
        healthy: false,
        service: this.name,
        error: `Помилка health check: ${error}`,
      };
    }
  }

  /**
   * Отримання статистики сервісу
   */
  public getStats(): ServiceStats {
    const baseStats: ServiceStats = {
      service: this.name,
      uptime: Date.now() - this.startTime,
      requests: 0,
      errors: 0,
    };

    const serviceStats = this.onGetStats();
    return { ...baseStats, ...serviceStats };
  }

  /**
   * Перевірка чи сервіс ініціалізовано
   */
  protected checkInitialized(): void {
    if (!this.isInitialized) {
      throw new Error(`Сервіс ${this.name} не ініціалізовано`);
    }
  }

  /**
   * Абстрактні методи для реалізації в нащадках
   */
  protected abstract onInitialize(): Promise<void>;
  protected abstract onShutdown(): Promise<void>;
  protected abstract onHealthCheck(): Promise<HealthStatus>;
  protected abstract onGetStats(): Partial<ServiceStats>;
} 