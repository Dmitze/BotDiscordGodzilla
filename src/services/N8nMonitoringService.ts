/**
 * N8n Monitoring Service
 * Моніторинг та сповіщення для n8n робочих процесів
 */

import type { BotConfig, ServiceStats } from '@/types';
import { BaseService } from '@/core/BaseService';
import logger from '@/utils/logger';
import { Counter, Gauge } from 'prom-client';

interface N8nWorkflowExecution {
  workflowId: string;
  workflowName: string;
  status: 'success' | 'failed' | 'running';
  startTime: number;
  endTime?: number;
  duration?: number;
  error?: string;
  nodeId?: string;
}

interface N8nMonitoringStats extends ServiceStats {
  totalExecutions: number;
  successfulExecutions: number;
  failedExecutions: number;
  activeWorkflows: number;
  averageExecutionTime: number;
}

interface N8nMetricsCollection {
  workflowExecutionsTotal: Counter<string>;
  workflowExecutionDuration: Gauge<string>;
  activeWorkflows: Gauge<string>;
  workflowFailuresTotal: Counter<string>;
  workflowSuccessRate: Gauge<string>;
}

export class N8nMonitoringService extends BaseService {
  private stats: N8nMonitoringStats;
  private metrics: N8nMetricsCollection | null = null;
  private activeExecutions: Map<string, N8nWorkflowExecution> = new Map();
  private executionHistory: N8nWorkflowExecution[] = [];
  private maxHistorySize: number = 1000;
  private alertThresholds: {
    failureRate: number; // percentage
    executionTime: number; // milliseconds
    consecutiveFailures: number;
  };

  private metricsService: any = null;

  constructor(config: BotConfig) {
    super('N8nMonitoringService', config);
    
    this.stats = {
      service: 'N8nMonitoringService',
      uptime: 0,
      requests: 0,
      errors: 0,
      startTime: Date.now(),
      totalExecutions: 0,
      successfulExecutions: 0,
      failedExecutions: 0,
      activeWorkflows: 0,
      averageExecutionTime: 0,
    };

    this.alertThresholds = {
      failureRate: (config as any).n8n?.alertThresholds?.failureRate || 10, // 10%
      executionTime: (config as any).n8n?.alertThresholds?.executionTime || 300000, // 5 minutes
      consecutiveFailures: (config as any).n8n?.alertThresholds?.consecutiveFailures || 3,
    };
  }

  /**
   * Ініціалізація залежностей сервісу
   * Викликається ServiceManager після створення всіх сервісів
   */
  public initializeServices(metricsService?: any): void {
    this.metricsService = metricsService;
    // Створення метрик, якщо доступний MetricsService
    this.createMetrics(this.metricsService).catch(error => {
      logger.error('Помилка створення метрик після ініціалізації залежностей:', {
        type: 'n8n_monitoring_service',
        event: 'metrics_init_failed',
        component: 'N8nMonitoringService',
        errorName: error instanceof Error ? error.name : undefined,
        errorMessage: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });
    });
  }

  /**
   * Ініціалізація сервісу моніторингу n8n
   */
  protected override async onInitialize(): Promise<void> {
    try {
      logger.info('📊 Ініціалізація N8nMonitoring сервісу...', {
        type: 'n8n_monitoring_service',
        event: 'init',
        component: 'N8nMonitoringService',
      });

      // Створення метрик, якщо доступний MetricsService
      // MetricsService буде переданий через initializeServices якщо доступний
      await this.createMetrics();

      logger.info('✅ N8nMonitoring сервіс ініціалізовано', {
        type: 'n8n_monitoring_service',
        event: 'init_success',
        component: 'N8nMonitoringService',
      });
    } catch (error) {
      logger.error('❌ Помилка ініціалізації N8nMonitoring сервісу:', {
        type: 'n8n_monitoring_service',
        event: 'init_failed',
        component: 'N8nMonitoringService',
        errorName: error instanceof Error ? error.name : undefined,
        errorMessage: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });
      throw error;
    }
  }

  /**
   * Створення метрик для Prometheus
   */
  private async createMetrics(metricsService?: any): Promise<void> {
    try {
      // Отримуємо MetricsService якщо він доступний
      if (!metricsService) {
        logger.debug('MetricsService недоступний, метрики не будуть створені', {
          type: 'n8n_monitoring_service',
          event: 'metrics_not_available',
          component: 'N8nMonitoringService',
        });
        return;
      }

      // Створюємо метрики через реєстр MetricsService
      const registry = (metricsService as any).registry;
      if (!registry) {
        logger.debug('Prometheus registry недоступний, метрики не будуть створені', {
          type: 'n8n_monitoring_service',
          event: 'registry_not_available',
          component: 'N8nMonitoringService',
        });
        return;
      }

      this.metrics = {
        workflowExecutionsTotal: new Counter({
          name: 'n8n_workflow_executions_total',
          help: 'Загальна кількість виконань n8n робочих процесів',
          labelNames: ['workflow_name', 'status'],
          registers: [registry],
        }),

        workflowExecutionDuration: new Gauge({
          name: 'n8n_workflow_execution_duration_seconds',
          help: 'Час виконання n8n робочих процесів в секундах',
          labelNames: ['workflow_name'],
          registers: [registry],
        }),

        activeWorkflows: new Gauge({
          name: 'n8n_active_workflows',
          help: 'Кількість активних n8n робочих процесів',
          registers: [registry],
        }),

        workflowFailuresTotal: new Counter({
          name: 'n8n_workflow_failures_total',
          help: 'Загальна кількість збоїв n8n робочих процесів',
          labelNames: ['workflow_name', 'error_type'],
          registers: [registry],
        }),

        workflowSuccessRate: new Gauge({
          name: 'n8n_workflow_success_rate',
          help: 'Відсоток успішних виконань n8n робочих процесів',
          labelNames: ['workflow_name'],
          registers: [registry],
        }),
      };

      logger.debug('✅ Метрики N8nMonitoring сервісу створено', {
        type: 'n8n_monitoring_service',
        event: 'metrics_created',
        component: 'N8nMonitoringService',
      });
    } catch (error) {
      logger.error('Помилка створення метрик N8nMonitoring сервісу:', {
        type: 'n8n_monitoring_service',
        event: 'metrics_create_failed',
        component: 'N8nMonitoringService',
        errorName: error instanceof Error ? error.name : undefined,
        errorMessage: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });
    }
  }

  /**
   * Початок виконання робочого процесу
   */
  public startWorkflowExecution(workflowId: string, workflowName: string): void {
    const execution: N8nWorkflowExecution = {
      workflowId,
      workflowName,
      status: 'running',
      startTime: Date.now(),
    };

    this.activeExecutions.set(workflowId, execution);
    this.stats.activeWorkflows = this.activeExecutions.size;
    
    // Оновлюємо метрики
    if (this.metrics) {
      this.metrics.activeWorkflows.set(this.stats.activeWorkflows);
    }

    logger.debug(`🚀 Початок виконання робочого процесу: ${workflowName}`, {
      type: 'n8n_monitoring_service',
      event: 'workflow_started',
      component: 'N8nMonitoringService',
      workflowId,
      workflowName,
    });
  }

  /**
   * Успішне завершення виконання робочого процесу
   */
  public completeWorkflowExecution(workflowId: string, nodeId?: string): void {
    const execution = this.activeExecutions.get(workflowId);
    if (!execution) {
      logger.warn(`Не знайдено активне виконання для workflowId: ${workflowId}`, {
        type: 'n8n_monitoring_service',
        event: 'execution_not_found',
        component: 'N8nMonitoringService',
        workflowId,
      });
      return;
    }

    const endTime = Date.now();
    const duration = endTime - execution.startTime;

    execution.status = 'success';
    execution.endTime = endTime;
    execution.duration = duration;
    if (nodeId !== undefined) {
      execution.nodeId = nodeId;
    }

    // Видаляємо з активних виконань
    this.activeExecutions.delete(workflowId);
    this.stats.activeWorkflows = this.activeExecutions.size;

    // Оновлюємо статистику
    this.stats.totalExecutions++;
    this.stats.successfulExecutions++;
    this.updateAverageExecutionTime(duration);

    // Додаємо до історії
    this.addToHistory(execution);

    // Оновлюємо метрики
    if (this.metrics) {
      this.metrics.workflowExecutionsTotal.inc({ workflow_name: execution.workflowName, status: 'success' }, 1);
      this.metrics.workflowExecutionDuration.set({ workflow_name: execution.workflowName }, duration / 1000);
      this.metrics.activeWorkflows.set(this.stats.activeWorkflows);
      this.updateSuccessRateMetric(execution.workflowName);
    }

    logger.debug(`✅ Успішне завершення робочого процесу: ${execution.workflowName}`, {
      type: 'n8n_monitoring_service',
      event: 'workflow_completed',
      component: 'N8nMonitoringService',
      workflowId,
      workflowName: execution.workflowName,
      duration,
    });

    // Перевіряємо наявність сповіщень
    this.checkForAlerts(execution);
  }

  /**
   * Помилка виконання робочого процесу
   */
  public failWorkflowExecution(workflowId: string, error: string, nodeId?: string): void {
    const execution = this.activeExecutions.get(workflowId);
    if (!execution) {
      logger.warn(`Не знайдено активне виконання для workflowId: ${workflowId}`, {
        type: 'n8n_monitoring_service',
        event: 'execution_not_found',
        component: 'N8nMonitoringService',
        workflowId,
      });
      return;
    }

    const endTime = Date.now();
    const duration = endTime - execution.startTime;

    execution.status = 'failed';
    execution.endTime = endTime;
    execution.duration = duration;
    execution.error = error;
    if (nodeId !== undefined) {
      execution.nodeId = nodeId;
    }

    // Видаляємо з активних виконань
    this.activeExecutions.delete(workflowId);
    this.stats.activeWorkflows = this.activeExecutions.size;

    // Оновлюємо статистику
    this.stats.totalExecutions++;
    this.stats.failedExecutions++;
    this.updateAverageExecutionTime(duration);

    // Додаємо до історії
    this.addToHistory(execution);

    // Оновлюємо метрики
    if (this.metrics) {
      this.metrics.workflowExecutionsTotal.inc({ workflow_name: execution.workflowName, status: 'failed' }, 1);
      this.metrics.workflowFailuresTotal.inc({ workflow_name: execution.workflowName, error_type: 'general' }, 1);
      this.metrics.workflowExecutionDuration.set({ workflow_name: execution.workflowName }, duration / 1000);
      this.metrics.activeWorkflows.set(this.stats.activeWorkflows);
      this.updateSuccessRateMetric(execution.workflowName);
    }

    logger.error(`❌ Помилка виконання робочого процесу: ${execution.workflowName}`, {
      type: 'n8n_monitoring_service',
      event: 'workflow_failed',
      component: 'N8nMonitoringService',
      workflowId,
      workflowName: execution.workflowName,
      error,
      duration,
    });

    // Перевіряємо наявність сповіщень
    this.checkForAlerts(execution);
  }

  /**
   * Додавання виконання до історії
   */
  private addToHistory(execution: N8nWorkflowExecution): void {
    this.executionHistory.push(execution);
    
    // Обмежуємо розмір історії
    if (this.executionHistory.length > this.maxHistorySize) {
      this.executionHistory.shift();
    }
  }

  /**
   * Оновлення середнього часу виконання
   */
  private updateAverageExecutionTime(duration: number): void {
    const totalExecutions = this.stats.successfulExecutions + this.stats.failedExecutions;
    if (totalExecutions > 0) {
      this.stats.averageExecutionTime = 
        (this.stats.averageExecutionTime * (totalExecutions - 1) + duration) / totalExecutions;
    }
  }

  /**
   * Оновлення метрики відсотка успішних виконань
   */
  private updateSuccessRateMetric(workflowName: string): void {
    if (!this.metrics) return;

    const workflowExecutions = this.executionHistory.filter(
      exec => exec.workflowName === workflowName
    );
    
    if (workflowExecutions.length > 0) {
      const successful = workflowExecutions.filter(exec => exec.status === 'success').length;
      const successRate = (successful / workflowExecutions.length) * 100;
      this.metrics.workflowSuccessRate.set({ workflow_name: workflowName }, successRate);
    }
  }

  /**
   * Перевірка наявності сповіщень
   */
  private checkForAlerts(execution: N8nWorkflowExecution): void {
    // Перевірка на надмірний час виконання
    if (execution.duration && execution.duration > this.alertThresholds.executionTime) {
      this.sendAlert(
        'warning',
        `Робочий процес "${execution.workflowName}" виконується занадто довго`,
        `Час виконання: ${execution.duration}ms, поріг: ${this.alertThresholds.executionTime}ms`
      );
    }

    // Перевірка на високий відсоток помилок
    const totalExecutions = this.stats.successfulExecutions + this.stats.failedExecutions;
    if (totalExecutions > 10) { // Потрібно мінімум 10 виконань для статистики
      const failureRate = (this.stats.failedExecutions / totalExecutions) * 100;
      if (failureRate > this.alertThresholds.failureRate) {
        this.sendAlert(
          'critical',
          'Високий відсоток помилок робочих процесів n8n',
          `Відсоток помилок: ${failureRate.toFixed(2)}%, поріг: ${this.alertThresholds.failureRate}%`
        );
      }
    }

    // Перевірка на послідовні помилки одного робочого процесу
    this.checkConsecutiveFailures(execution.workflowName);
  }

  /**
   * Перевірка на послідовні помилки одного робочого процесу
   */
  private checkConsecutiveFailures(workflowName: string): void {
    const recentExecutions = this.executionHistory
      .filter(exec => exec.workflowName === workflowName)
      .slice(-this.alertThresholds.consecutiveFailures);

    if (recentExecutions.length >= this.alertThresholds.consecutiveFailures) {
      const allFailed = recentExecutions.every(exec => exec.status === 'failed');
      if (allFailed) {
        this.sendAlert(
          'critical',
          `Послідовні помилки робочого процесу "${workflowName}"`,
          `Кількість послідовних помилок: ${this.alertThresholds.consecutiveFailures}`
        );
      }
    }
  }

  /**
   * Відправка сповіщення
   */
  private async sendAlert(level: 'warning' | 'critical', title: string, message: string): Promise<void> {
    try {
      // Логуємо сповіщення
      logger[level === 'critical' ? 'error' : 'warn'](`🔔 Алерт n8n: ${title}`, {
        type: 'n8n_monitoring_service',
        event: 'alert',
        component: 'N8nMonitoringService',
        level,
        title,
        message,
      });

      logger.debug('Сповіщення про помилки N8n буде інтегровано пізніше', { workflowId });
      // Це може бути реалізовано в майбутньому
    } catch (error) {
      logger.error('Помилка відправки алерту:', {
        type: 'n8n_monitoring_service',
        event: 'alert_send_failed',
        component: 'N8nMonitoringService',
        errorName: error instanceof Error ? error.name : undefined,
        errorMessage: error instanceof Error ? error.message : String(error),
      });
    }
  }

  /**
   * Зупинка сервісу
   */
  protected override async onShutdown(): Promise<void> {
    logger.info('🛑 N8nMonitoring сервіс зупинено', {
      type: 'n8n_monitoring_service',
      event: 'shutdown',
      component: 'N8nMonitoringService',
    });
  }

  /**
   * Перевірка стану здоров'я сервісу
   */
  protected override async onHealthCheck(): Promise<any> {
    try {
      // Basic health check - service is running
      return {
        healthy: true,
        service: 'N8nMonitoringService',
        message: 'N8nMonitoring service is running'
      };
    } catch (error) {
      return {
        healthy: false,
        service: 'N8nMonitoringService',
        error: error instanceof Error ? error.message : 'Unknown error'
      };
    }
  }

  /**
   * Отримання статистики сервісу
   */
  protected override onGetStats(): Partial<ServiceStats> {
    return {
      requests: this.stats.requests,
      errors: this.stats.errors,
      totalExecutions: this.stats.totalExecutions,
      successfulExecutions: this.stats.successfulExecutions,
      failedExecutions: this.stats.failedExecutions,
      activeWorkflows: this.stats.activeWorkflows,
      averageExecutionTime: this.stats.averageExecutionTime,
    };
  }

  /**
   * Отримання історії виконань
   */
  public getExecutionHistory(limit: number = 50): N8nWorkflowExecution[] {
    return this.executionHistory.slice(-limit);
  }

  /**
   * Отримання активних виконань
   */
  public getActiveExecutions(): N8nWorkflowExecution[] {
    return Array.from(this.activeExecutions.values());
  }

  /**
   * Очищення історії
   */
  public clearHistory(): void {
    this.executionHistory = [];
    logger.info('Історія виконань n8n очищена', {
      type: 'n8n_monitoring_service',
      event: 'history_cleared',
      component: 'N8nMonitoringService',
    });
  }

  /**
   * Зупинка сервісу
   */
  protected async onStop(): Promise<void> {
    logger.info('Зупинка N8nMonitoring сервісу...', {
      type: 'n8n_monitoring_service',
      event: 'stopping',
      component: 'N8nMonitoringService',
    });

    // Очищуємо активні виконання
    this.activeExecutions.clear();
    this.executionHistory = [];

    logger.info('✅ N8nMonitoring сервіс зупинено', {
      type: 'n8n_monitoring_service',
      event: 'stopped',
      component: 'N8nMonitoringService',
    });
  }

  /**
   * Get service statistics
   */
  public override getStats(): ServiceStats {
    const baseStats = super.getStats();
    return {
      ...baseStats,
      uptime: Date.now() - ((this.stats as any).startTime || Date.now()),
      requests: this.stats.totalExecutions,
      errors: this.stats.errors,
      activeWorkflows: this.stats.activeWorkflows,
      totalExecutions: this.stats.totalExecutions,
      successfulExecutions: this.stats.successfulExecutions,
      failedExecutions: this.stats.failedExecutions,
      averageExecutionTime: this.stats.averageExecutionTime,
    };
  }
}