/**
 * Контейнер сервісів з Dependency Injection
 * Централізоване управління всіма сервісами
 */

import type { BaseService, BotConfig, HealthStatus } from '@/types';
import { BaseService as BaseServiceClass } from './BaseService';

export class ServiceContainer {
  private services = new Map<string, BaseService>();
  private readonly config: BotConfig;

  constructor(config: BotConfig) {
    this.config = config;
  }

  /**
   * Реєстрація сервісу
   */
  public register<T extends BaseService>(name: string, service: T): void {
    if (this.services.has(name)) {
      throw new Error(`Сервіс ${name} вже зареєстрований`);
    }

    this.services.set(name, service);
  }

  /**
   * Отримання сервісу
   */
  public get<T extends BaseService>(name: string): T {
    const service = this.services.get(name);
    if (!service) {
      throw new Error(`Сервіс ${name} не знайдено`);
    }

    return service as T;
  }

  /**
   * Перевірка чи сервіс існує
   */
  public has(name: string): boolean {
    return this.services.has(name);
  }

  /**
   * Отримання всіх сервісів
   */
  public getAll(): Map<string, BaseService> {
    return new Map(this.services);
  }

  /**
   * Ініціалізація всіх сервісів
   */
  public async initialize(): Promise<void> {
    const initPromises: Promise<void>[] = [];

    for (const [name, service] of this.services.entries()) {
      try {
        initPromises.push(service.initialize());
      } catch (error) {
        throw new Error(`Помилка ініціалізації сервісу ${name}: ${error}`);
      }
    }

    await Promise.all(initPromises);
  }

  /**
   * Завершення роботи всіх сервісів
   */
  public async shutdown(): Promise<void> {
    const shutdownPromises: Promise<void>[] = [];

    for (const [name, service] of this.services.entries()) {
      try {
        shutdownPromises.push(service.shutdown());
      } catch (error) {
        console.error(`Помилка завершення сервісу ${name}:`, error);
      }
    }

    await Promise.all(shutdownPromises);
  }

  /**
   * Health check всіх сервісів
   */
  public async getHealthStatus(): Promise<Record<string, HealthStatus>> {
    const healthStatus: Record<string, HealthStatus> = {};

    for (const [name, service] of this.services.entries()) {
      try {
        healthStatus[name] = await service.healthCheck();
      } catch (error) {
        healthStatus[name] = {
          healthy: false,
          service: name,
          error: `Помилка health check: ${error}`,
        };
      }
    }

    return healthStatus;
  }

  /**
   * Отримання статистики всіх сервісів
   */
  public getAllStats(): Record<string, unknown> {
    const stats: Record<string, unknown> = {};

    for (const [name, service] of this.services.entries()) {
      try {
        stats[name] = service.getStats();
      } catch (error) {
        stats[name] = {
          error: `Помилка отримання статистики: ${error}`,
        };
      }
    }

    return stats;
  }

  /**
   * Видалення сервісу
   */
  public remove(name: string): boolean {
    return this.services.delete(name);
  }

  /**
   * Очищення всіх сервісів
   */
  public clear(): void {
    this.services.clear();
  }

  /**
   * Отримання кількості сервісів
   */
  public get size(): number {
    return this.services.size;
  }
} 