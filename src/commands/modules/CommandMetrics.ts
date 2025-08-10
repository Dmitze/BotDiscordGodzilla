/**
 * Система метрик та статистики для команд Discord бота
 * Збір, аналіз та звітність по використанню команд
 * Версія 1.0.0 - Виокремлено з BaseCommand
 */

import type { ChatInputCommandInteraction } from 'discord.js';
import type { CommandStats, LogMeta } from '@/types';
import logger from '@/utils/logger';

export interface CommandMetrics {
  commandName: string;
  executionCount: number;
  successCount: number;
  errorCount: number;
  averageExecutionTime: number;
  totalExecutionTime: number;
  lastExecuted: number;
  slowExecutions: number;
  cacheHits: number;
  cacheMisses: number;
  userCount: number;
  retries: number;
  cooldownHits: number;
}

export interface ExecutionMetric {
  commandName: string;
  userId: string;
  executionTime: number;
  success: boolean;
  error?: string;
  timestamp: number;
  fromCache: boolean;
  retryCount: number;
}

export interface PerformanceThresholds {
  slowExecutionMs: number;
  verySlowExecutionMs: number;
  maxExecutionTime: number;
  warningErrorRate: number; // %
  criticalErrorRate: number; // %
}

export class CommandMetricsCollector {
  private static instance: CommandMetricsCollector | null = null;
  private metrics = new Map<string, CommandMetrics>();
  private executionHistory: ExecutionMetric[] = [];
  private readonly maxHistorySize = 10000;
  private readonly thresholds: PerformanceThresholds;

  constructor(thresholds?: Partial<PerformanceThresholds>) {
    if (CommandMetricsCollector.instance) {
      return CommandMetricsCollector.instance;
    }

    this.thresholds = {
      slowExecutionMs: 3000,
      verySlowExecutionMs: 10000,
      maxExecutionTime: 30000,
      warningErrorRate: 10,
      criticalErrorRate: 25,
      ...thresholds
    };

    CommandMetricsCollector.instance = this;
    this.startPeriodicReporting();
  }

  /**
   * Записати метрику виконання команди
   */
  public recordExecution(
    commandName: string,
    userId: string,
    executionTime: number,
    success: boolean,
    options: {
      error?: string;
      fromCache?: boolean;
      retryCount?: number;
    } = {}
  ): void {
    try {
      // Отримати або створити метрики для команди
      let commandMetrics = this.metrics.get(commandName);
      if (!commandMetrics) {
        commandMetrics = this.createEmptyMetrics(commandName);
        this.metrics.set(commandName, commandMetrics);
      }

      // Оновити метрики команди
      this.updateCommandMetrics(commandMetrics, executionTime, success, options);

      // Додати в історію виконань
      const executionMetric: ExecutionMetric = {
        commandName,
        userId,
        executionTime,
        success,
        error: options.error,
        timestamp: Date.now(),
        fromCache: options.fromCache || false,
        retryCount: options.retryCount || 0
      };

      this.addToHistory(executionMetric);

      // Логування повільних виконань
      if (executionTime > this.thresholds.slowExecutionMs) {
        logger.warn('🐌 Повільне виконання команди', {
          command: commandName,
          userId,
          executionTime: `${executionTime}ms`,
          threshold: `${this.thresholds.slowExecutionMs}ms`
        } as LogMeta);
      }

      // Логування критично повільних виконань
      if (executionTime > this.thresholds.verySlowExecutionMs) {
        logger.error('🚨 Критично повільне виконання команди', {
          command: commandName,
          userId,
          executionTime: `${executionTime}ms`,
          threshold: `${this.thresholds.verySlowExecutionMs}ms`
        } as LogMeta);
      }

    } catch (error) {
      logger.error('❌ Помилка запису метрик команди:', error);
    }
  }

  /**
   * Створити порожні метрики для нової команди
   */
  private createEmptyMetrics(commandName: string): CommandMetrics {
    return {
      commandName,
      executionCount: 0,
      successCount: 0,
      errorCount: 0,
      averageExecutionTime: 0,
      totalExecutionTime: 0,
      lastExecuted: 0,
      slowExecutions: 0,
      cacheHits: 0,
      cacheMisses: 0,
      userCount: 0,
      retries: 0,
      cooldownHits: 0
    };
  }

  /**
   * Оновити метрики команди
   */
  private updateCommandMetrics(
    metrics: CommandMetrics,
    executionTime: number,
    success: boolean,
    options: {
      error?: string;
      fromCache?: boolean;
      retryCount?: number;
    }
  ): void {
    metrics.executionCount++;
    metrics.lastExecuted = Date.now();

    if (success) {
      metrics.successCount++;
    } else {
      metrics.errorCount++;
    }

    // Оновлення часу виконання
    metrics.totalExecutionTime += executionTime;
    metrics.averageExecutionTime = metrics.totalExecutionTime / metrics.executionCount;

    // Повільні виконання
    if (executionTime > this.thresholds.slowExecutionMs) {
      metrics.slowExecutions++;
    }

    // Кеш метрики
    if (options.fromCache) {
      metrics.cacheHits++;
    } else {
      metrics.cacheMisses++;
    }

    // Повтори
    if (options.retryCount) {
      metrics.retries += options.retryCount;
    }
  }

  /**
   * Додати виконання в історію
   */
  private addToHistory(execution: ExecutionMetric): void {
    this.executionHistory.push(execution);

    // Обмеження розміру історії
    if (this.executionHistory.length > this.maxHistorySize) {
      this.executionHistory = this.executionHistory.slice(-this.maxHistorySize * 0.8);
    }
  }

  /**
   * Отримати метрики команди
   */
  public getCommandMetrics(commandName: string): CommandMetrics | undefined {
    return this.metrics.get(commandName);
  }

  /**
   * Отримати всі метрики
   */
  public getAllMetrics(): CommandMetrics[] {
    return Array.from(this.metrics.values());
  }

  /**
   * Отримати топ команд за використанням
   */
  public getTopCommands(limit: number = 10): CommandMetrics[] {
    return Array.from(this.metrics.values())
      .sort((a, b) => b.executionCount - a.executionCount)
      .slice(0, limit);
  }

  /**
   * Отримати команди з найбільшою кількістю помилок
   */
  public getCommandsWithMostErrors(limit: number = 10): CommandMetrics[] {
    return Array.from(this.metrics.values())
      .filter(m => m.errorCount > 0)
      .sort((a, b) => b.errorCount - a.errorCount)
      .slice(0, limit);
  }

  /**
   * Отримати найповільніші команди
   */
  public getSlowestCommands(limit: number = 10): CommandMetrics[] {
    return Array.from(this.metrics.values())
      .sort((a, b) => b.averageExecutionTime - a.averageExecutionTime)
      .slice(0, limit);
  }

  /**
   * Аналіз трендів використання
   */
  public analyzeTrends(timeframe: number = 24 * 60 * 60 * 1000): {
    totalExecutions: number;
    errorRate: number;
    averageResponseTime: number;
    topCommands: string[];
    trendDirection: 'up' | 'down' | 'stable';
  } {
    const now = Date.now();
    const recentExecutions = this.executionHistory.filter(
      e => now - e.timestamp < timeframe
    );

    const totalExecutions = recentExecutions.length;
    const errors = recentExecutions.filter(e => !e.success).length;
    const errorRate = totalExecutions > 0 ? (errors / totalExecutions) * 100 : 0;

    const totalTime = recentExecutions.reduce((sum, e) => sum + e.executionTime, 0);
    const averageResponseTime = totalExecutions > 0 ? totalTime / totalExecutions : 0;

    // Топ команди
    const commandCounts: Record<string, number> = {};
    recentExecutions.forEach(e => {
      commandCounts[e.commandName] = (commandCounts[e.commandName] || 0) + 1;
    });

    const topCommands = Object.entries(commandCounts)
      .sort(([,a], [,b]) => b - a)
      .slice(0, 5)
      .map(([command]) => command);

    // Аналіз тренду (порівняння з попереднім періодом)
    const previousTimeframe = this.executionHistory.filter(
      e => now - e.timestamp >= timeframe && now - e.timestamp < timeframe * 2
    );

    let trendDirection: 'up' | 'down' | 'stable' = 'stable';
    if (previousTimeframe.length > 0) {
      const previousCount = previousTimeframe.length;
      const changePercent = ((totalExecutions - previousCount) / previousCount) * 100;
      
      if (changePercent > 20) trendDirection = 'up';
      else if (changePercent < -20) trendDirection = 'down';
    }

    return {
      totalExecutions,
      errorRate,
      averageResponseTime,
      topCommands,
      trendDirection
    };
  }

  /**
   * Генерація звіту про продуктивність
   */
  public generatePerformanceReport(): {
    summary: {
      totalCommands: number;
      totalExecutions: number;
      overallErrorRate: number;
      averageResponseTime: number;
    };
    alerts: Array<{
      level: 'warning' | 'critical';
      message: string;
      command?: string;
      metric?: string;
      value?: number;
    }>;
    recommendations: string[];
  } {
    const allMetrics = this.getAllMetrics();
    const totalExecutions = allMetrics.reduce((sum, m) => sum + m.executionCount, 0);
    const totalErrors = allMetrics.reduce((sum, m) => sum + m.errorCount, 0);
    const totalTime = allMetrics.reduce((sum, m) => sum + m.totalExecutionTime, 0);

    const overallErrorRate = totalExecutions > 0 ? (totalErrors / totalExecutions) * 100 : 0;
    const averageResponseTime = totalExecutions > 0 ? totalTime / totalExecutions : 0;

    const alerts: Array<{
      level: 'warning' | 'critical';
      message: string;
      command?: string;
      metric?: string;
      value?: number;
    }> = [];

    const recommendations: string[] = [];

    // Аналіз на проблеми
    allMetrics.forEach(metric => {
      const errorRate = metric.executionCount > 0 ? (metric.errorCount / metric.executionCount) * 100 : 0;

      // Критичний рівень помилок
      if (errorRate > this.thresholds.criticalErrorRate) {
        alerts.push({
          level: 'critical',
          message: `Критичний рівень помилок`,
          command: metric.commandName,
          metric: 'errorRate',
          value: errorRate
        });
        recommendations.push(`Терміново перевірити команду ${metric.commandName} - ${errorRate.toFixed(1)}% помилок`);
      }
      // Попереджувальний рівень помилок
      else if (errorRate > this.thresholds.warningErrorRate) {
        alerts.push({
          level: 'warning',
          message: `Високий рівень помилок`,
          command: metric.commandName,
          metric: 'errorRate',
          value: errorRate
        });
      }

      // Повільні команди
      if (metric.averageExecutionTime > this.thresholds.slowExecutionMs) {
        alerts.push({
          level: 'warning',
          message: `Повільна команда`,
          command: metric.commandName,
          metric: 'averageExecutionTime',
          value: metric.averageExecutionTime
        });
        recommendations.push(`Оптимізувати продуктивність команди ${metric.commandName} - ${metric.averageExecutionTime.toFixed(0)}ms`);
      }

      // Низька ефективність кешу
      const totalCacheRequests = metric.cacheHits + metric.cacheMisses;
      if (totalCacheRequests > 10) {
        const cacheHitRate = (metric.cacheHits / totalCacheRequests) * 100;
        if (cacheHitRate < 50) {
          recommendations.push(`Покращити стратегію кешування для команди ${metric.commandName} - ${cacheHitRate.toFixed(1)}% hits`);
        }
      }
    });

    // Загальні рекомендації
    if (overallErrorRate > this.thresholds.warningErrorRate) {
      recommendations.push('Провести загальний аудит обробки помилок');
    }

    if (averageResponseTime > this.thresholds.slowExecutionMs) {
      recommendations.push('Оптимізувати загальну продуктивність системи');
    }

    return {
      summary: {
        totalCommands: allMetrics.length,
        totalExecutions,
        overallErrorRate,
        averageResponseTime
      },
      alerts,
      recommendations
    };
  }

  /**
   * Записати cooldown hit
   */
  public recordCooldownHit(commandName: string): void {
    const metrics = this.metrics.get(commandName);
    if (metrics) {
      metrics.cooldownHits++;
    }
  }

  /**
   * Очистити метрики (для тестування)
   */
  public clearMetrics(): void {
    this.metrics.clear();
    this.executionHistory = [];
  }

  /**
   * Періодичне звітування
   */
  private startPeriodicReporting(): void {
    setInterval(() => {
      try {
        const report = this.generatePerformanceReport();
        
        // Логування критичних алертів
        const criticalAlerts = report.alerts.filter(a => a.level === 'critical');
        if (criticalAlerts.length > 0) {
          logger.error('🚨 Критичні проблеми продуктивності команд:', {
            alerts: criticalAlerts,
            recommendations: report.recommendations
          } as LogMeta);
        }

        // Логування статистики
        logger.info('📊 Періодичний звіт команд:', {
          summary: report.summary,
          alertsCount: report.alerts.length,
          recommendationsCount: report.recommendations.length
        } as LogMeta);

      } catch (error) {
        logger.error('❌ Помилка періодичного звітування метрик:', error);
      }
    }, 15 * 60 * 1000); // Кожні 15 хвилин
  }

  /**
   * Експорт метрик для Prometheus
   */
  public exportPrometheusMetrics(): string {
    const metrics = this.getAllMetrics();
    let output = '';

    // Загальна кількість виконань
    output += '# TYPE discord_bot_command_executions_total counter\n';
    metrics.forEach(m => {
      output += `discord_bot_command_executions_total{command="${m.commandName}"} ${m.executionCount}\n`;
    });

    // Помилки
    output += '# TYPE discord_bot_command_errors_total counter\n';
    metrics.forEach(m => {
      output += `discord_bot_command_errors_total{command="${m.commandName}"} ${m.errorCount}\n`;
    });

    // Час виконання
    output += '# TYPE discord_bot_command_duration_seconds histogram\n';
    metrics.forEach(m => {
      const avgSeconds = m.averageExecutionTime / 1000;
      output += `discord_bot_command_duration_seconds{command="${m.commandName}"} ${avgSeconds}\n`;
    });

    return output;
  }
}

export default CommandMetricsCollector;