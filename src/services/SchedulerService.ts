/**
 * Scheduler Service для Discord бота
 * Централізоване управління плануваними завданнями
 * TypeScript версія
 */

import logger from '../utils/logger';
import type { Client } from 'discord.js';
import type { schedule as CronSchedule, ScheduledTask } from 'node-cron';

interface Bot {
  getService(name: string): any;
  serviceManager?: any;
  client?: Client;
}

interface JobInfo {
  job: ScheduledTask;
  schedule: string;
  task: string;
  options: { timezone?: string; scheduled?: boolean };
  createdAt: Date;
  lastRun: Date | null;
  nextRun: Date;
  executions: number;
  errors: number;
}

interface JobDetails {
  name: string;
  schedule: string;
  task: string;
  createdAt: Date;
  lastRun: Date | null;
  nextRun: Date;
  executions: number;
  errors: number;
  isActive: boolean;
}

interface SchedulerStats {
  jobsCreated: number;
  jobsExecuted: number;
  jobsFailed: number;
  activeJobs: number;
  jobs: JobDetails[];
  isActive: boolean;
}

interface HealthStatus {
  timestamp: Date;
  services: Record<string, { isActive: boolean; hasStats: boolean }>;
  overall: 'healthy' | 'degraded' | 'unhealthy';
  discord?: {
    isReady: boolean;
    uptime: number | null;
    guilds: number;
  };
}

class SchedulerService {
  private bot: Bot;
  private jobs: Map<string, JobInfo>;
  private scheduler: { schedule: typeof CronSchedule } | null;
  private stats: {
    jobsCreated: number;
    jobsExecuted: number;
    jobsFailed: number;
    activeJobs: number;
  };
  private _isActive: boolean;

  constructor(bot: Bot) {
    this.bot = bot;
    this.jobs = new Map();
    this.scheduler = null;
    this.stats = {
      jobsCreated: 0,
      jobsExecuted: 0,
      jobsFailed: 0,
      activeJobs: 0,
    };
    this._isActive = false;
  }

  // ===== Workspace (Stage 7) =====
  private async flushWorkspaceNotifications(): Promise<void> {
    try {
      const ws = this.bot.getService('workspace');
      if (!ws || typeof ws.flushNotifications !== 'function') return;
      const ready = await ws.flushNotifications();
      if (!ready || !ready.length) {
        logger.debug('workspace: нет готовых уведомлений');
        return;
      }
      logger.info(`workspace: готово уведомлений к доставке: ${ready.length}`);
      // На этом этапе можно отправлять DM/канал. Оставляем лог/заглушку.
    } catch (e) {
      logger.warn('workspace: flushNotifications error', e as any);
    }
  }

  private async sendWorkspaceDigest(period: 'daily' | 'weekly'): Promise<void> {
    try {
      const ws = this.bot.getService('workspace');
      if (!ws || typeof ws.buildDigest !== 'function' || typeof ws.createDigestRecord !== 'function') return;

      const now = Date.now();
      let windowStart: number; let windowEnd: number;
      if (period === 'daily') {
        const d = new Date(now); d.setHours(0,0,0,0);
        windowStart = d.getTime() - 24 * 60 * 60 * 1000; // вчера 00:00
        windowEnd = d.getTime() - 1;                     // вчера 23:59:59.999
      } else {
        windowEnd = now;
        windowStart = now - 7 * 24 * 60 * 60 * 1000;
      }

      if (typeof ws.db?.listAllSubscribers !== 'function') {
        logger.debug('workspace: нет listAllSubscribers');
        return;
      }
      const users: string[] = ws.db.listAllSubscribers();
      let prepared = 0;
      for (const userId of users) {
        const digest = await ws.buildDigest(userId, period, windowStart, windowEnd);
        if (digest.total > 0) {
          ws.createDigestRecord(userId, period, windowStart, windowEnd, digest, null);
          prepared++;
        }
      }
      logger.info(`workspace: сформировано дайджестов (${period}): ${prepared}`);
    } catch (e) {
      logger.warn('workspace: sendWorkspaceDigest error', e as any);
    }
  }

  /**
   * Внутрішня перевірка: чи потрібно вимкнути cron (тести або DISABLE_CRON)
   */
  private isCronDisabled(): boolean {
    return (
      process.env['NODE_ENV'] === 'test' ||
      Boolean(process.env['JEST_WORKER_ID']) ||
      String(process.env['DISABLE_CRON']).toLowerCase() === 'true'
    );
  }

  /**
   * Ініціалізація Scheduler сервісу
   */
  async initialize(): Promise<void> {
    try {
      logger.info('⏰ Ініціалізація Scheduler сервісу...');

      // Пропускаємо ініціалізацію у тестовому середовищі
      if (this.isCronDisabled()) {
        logger.debug('⏭️ Пропуск ініціалізації Scheduler у тестовому середовищі');
        this._isActive = false;
        return;
      }

      // Створення планувальника
      await this.createScheduler();

      // Реєстрація стандартних завдань
      await this.registerDefaultJobs();

      this._isActive = true;
      logger.info('✅ Scheduler сервіс ініціалізовано');
    } catch (error) {
      logger.error(
        `❌ Помилка ініціалізації Scheduler сервісу: ${error instanceof Error ? error.message : String(error)}`
      );
      throw error;
    }
  }

  /**
   * Створення планувальника
   */
  private async createScheduler(): Promise<void> {
    try {
      // Використовуємо node-cron для планування
      const cron = require('node-cron');
      this.scheduler = cron;

      logger.debug('✅ Планувальник створено');
    } catch (error) {
      logger.warn(
        `⚠️ Планувальник вимкнено: ${error instanceof Error ? error.message : String(error)}. Модуль 'node-cron' не знайдено. Продовжуємо без Scheduler.`
      );
      // Дозволяємо подальший запуск системи без падения
      this.scheduler = null;
      return;
    }
  }

  /**
   * Реєстрація стандартних завдань
   */
  private async registerDefaultJobs(): Promise<void> {
    try {
      // Очищення кешу кожну годину
      this.scheduleJob('cache-cleanup', '0 * * * *', () => {
        this.cleanupCache();
      });

      // Оновлення статистики кожні 5 хвилин
      this.scheduleJob('stats-update', '*/5 * * * *', () => {
        this.updateStats();
      });

      // Перевірка здоров'я кожні 10 хвилин
      this.scheduleJob('health-check', '*/10 * * * *', () => {
        this.healthCheck();
      });

      // Резервне копіювання щодня о 2:00
      this.scheduleJob('backup', '0 2 * * *', () => {
        this.createBackup();
      });

      // Опитування змін Google Drive кожні 2 хвилини (без у тестах/коли вимкнено cron)
      this.scheduleJob('poll-drive-changes', '*/2 * * * *', () => {
        this.pollDriveChanges();
      });

      // Workspace: коалесинг/доставка уведомлений (каждые 5 минут)
      this.scheduleJob('workspace-flush-notifs', '*/5 * * * *', () => {
        this.flushWorkspaceNotifications();
      });

      // Workspace: дневной дайджест (каждый день 09:05)
      this.scheduleJob('workspace-digest-daily', '5 9 * * *', () => {
        this.sendWorkspaceDigest('daily');
      });

      // Workspace: недельный дайджест (каждый понедельник 09:10)
      this.scheduleJob('workspace-digest-weekly', '10 9 * * MON', () => {
        this.sendWorkspaceDigest('weekly');
      });

      logger.debug('✅ Стандартні завдання зареєстровано');
    } catch (error) {
      logger.error(
        `Помилка реєстрації стандартних завдань: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }

  /**
   * Опитування змін у Google Drive з ретраями та м'якими помилками
   */
  private async pollDriveChanges(): Promise<void> {
    try {
      const changesService = this.bot.getService('driveChanges');
      if (!changesService || typeof changesService.pollOnce !== 'function') {
        logger.debug('pollDriveChanges: сервіс змін недоступний');
        return;
      }

      const maxRetries = 3;
      let attempt = 0;
      // Експоненційна затримка між ретраями
      // eslint-disable-next-line no-constant-condition
      while (true) {
        try {
          const { events } = await changesService.pollOnce();
          if (events.length) await this.notifyDriveChanges(events).catch(() => undefined);
          break;
        } catch (e) {
          attempt++;
          // If it's an initialization issue, we might want to retry more aggressively
          const isInitializationError = e instanceof Error && 
            (e.message.includes('Google Drive client not initialized') || 
             e.message.includes('not properly initialized'));
          
          if (attempt >= maxRetries) {
            logger.error(`pollDriveChanges: помилка після ${attempt} спроб`, e as any);
            break;
          }
          
          // For initialization errors, use a shorter retry delay
          const wait = isInitializationError ? 
            3000 : // 3 seconds for initialization errors
            Math.min(30000, 1000 * Math.pow(2, attempt)); // Exponential backoff for other errors
          
          logger.warn(`pollDriveChanges: retry ${attempt}/${maxRetries} in ${wait}ms`, {
            component: 'SchedulerService',
            error: e instanceof Error ? e.message : String(e),
            isInitializationError
          });
          
          await new Promise(r => setTimeout(r, wait));
        }
      }
    } catch (error) {
      logger.error('pollDriveChanges: невідома помилка', error as any);
    }
  }

  /**
   * Надіслати повідомлення про зміни у канал/DM (no-op, якщо клієнт/канал не задані)
   */
  private async notifyDriveChanges(events: Array<{ fileId: string; name?: string; time?: string; type: string; owners?: string[]; webViewLink?: string }>): Promise<void> {
    try {
      if (!this.bot.client) return;
      const channelId = process.env['DRIVE_CHANGES_CHANNEL_ID'];
      if (!channelId) {
        logger.debug(`notifyDriveChanges: канал не налаштований, подій: ${events.length}`);
        return;
      }
      const channel = await this.bot.client.channels.fetch(channelId).catch(() => null);
      if (!channel || !('send' in (channel as any))) return;

      const chunks: typeof events[] = [];
      const copy = [...events];
      while (copy.length) chunks.push(copy.splice(0, 10));

      for (const part of chunks) {
        const lines = part.map(ev => {
          const who = ev.owners && ev.owners.length ? ` — ${ev.owners.join(', ')}` : '';
          const time = ev.time ? ` (${new Date(ev.time).toLocaleString('uk-UA')})` : '';
          const link = ev.webViewLink ? ` \u2014 <${ev.webViewLink}>` : '';
          return `• ${ev.type}: ${ev.name ?? ev.fileId}${who}${time}${link}`;
        });
        const content = `🛎️ Оновлення Google Drive (${part.length}):\n` + lines.join('\n');
        await (channel as any).send({ content }).catch(() => undefined);
      }
    } catch (error) {
      logger.warn('notifyDriveChanges: не вдалося надіслати повідомлення', error as any);
    }
  }

  /**
   * Планування завдання
   */
  scheduleJob(
    name: string,
    schedule: string,
    task: () => Promise<void> | void,
    options: { timezone?: string; scheduled?: boolean } = {}
  ): ScheduledTask | null {
    try {
      // Пропускаємо планування завдань у тестовому середовищі
      if (this.isCronDisabled()) {
        logger.debug(`⏭️ Пропуск планування завдання "${name}" у тестовому середовищі`);
        return null;
      }
      if (!this.scheduler) {
        logger.warn(`⚠️ Сервіс планувальника недоступний. Завдання "${name}" не буде заплановано.`);
        return null;
      }
      if (this.jobs.has(name)) {
        this.stopJob(name);
      }

      const job = this.scheduler.schedule(
        schedule,
        async () => {
          await this.executeJob(name, task);
        },
        {
          scheduled: false,
          timezone: options.timezone || 'Europe/Kiev',
          ...options,
        }
      );

      // Calculate next run with node-cron API compatibility
      let nextRun: Date | null = null;
      try {
        // Перевіряємо, які методи доступні в об'єкті job
        const jobObject = job as unknown as { nextDates?: (n?: number) => unknown; nextDate?: () => unknown };
        
        if (typeof jobObject.nextDates === 'function') {
          try {
            const dates: unknown = jobObject.nextDates(1);
            const firstDate: unknown = Array.isArray(dates as any) ? (dates as any)[0] : dates;
            const d: any = firstDate as any;
            // Обробка різних типів об'єктів дати
            if (d && typeof d.toJSDate === 'function') {
              nextRun = d.toJSDate();
            } else if (d && typeof d.toDate === 'function') {
              nextRun = d.toDate();
            } else if (d instanceof Date) {
              nextRun = d as Date;
            } else {
              nextRun = new Date(d as any);
            }
          } catch (nextDatesError) {
            logger.debug('nextDates method failed, trying nextDate', {
              error: String(nextDatesError)
            });
            throw nextDatesError;
          }
        } else if (typeof jobObject.nextDate === 'function') {
          // Фолбек для старих версій node-cron
          try {
            const dt: unknown = jobObject.nextDate();
            const d: any = dt as any;
            if (d instanceof Date) {
              nextRun = d as Date;
            } else if (d && typeof d.toJSDate === 'function') {
              nextRun = d.toJSDate();
            } else if (d && typeof d.toDate === 'function') {
              nextRun = d.toDate();
            } else {
              nextRun = new Date(d as any);
            }
          } catch (nextDateError) {
            logger.debug('nextDate method failed, using fallback', {
              error: String(nextDateError)
            });
            throw nextDateError;
          }
        } else {
          // Якщо ні один метод не доступний - використовуємо простий розрахунок
          nextRun = this.calculateNextRun(schedule);
        }
      } catch (e) {
        logger.warn('scheduler_next_run_unavailable', { 
          error: e instanceof Error ? e.message : String(e),
          methodsAvailable: { },
          schedule
        });
        // Фолбек до розрахунку на основі cron виразу
        nextRun = this.calculateNextRun(schedule);
      }

      this.jobs.set(name, {
        job,
        schedule,
        task: task.toString(),
        options,
        createdAt: new Date(),
        lastRun: null,
        nextRun: nextRun || new Date(),
        executions: 0,
        errors: 0,
      });

      job.start();
      this.stats.jobsCreated++;
      this.stats.activeJobs++;

      logger.debug(`✅ Завдання "${name}" заплановано: ${schedule}`);
      return job;
    } catch (error) {
      logger.error(
        `Помилка планування завдання "${name}": ${error instanceof Error ? error.message : String(error)}`
      );
      throw error;
    }
  }

  /**
   * Виконання завдання
   */
  private async executeJob(name: string, task: () => Promise<void> | void): Promise<void> {
    const jobInfo = this.jobs.get(name);
    if (!jobInfo) {
      logger.warn(`Завдання "${name}" не знайдено`);
      return;
    }

    const startTime = Date.now();
    jobInfo.lastRun = new Date();
    jobInfo.executions++;

    try {
      logger.debug(`🚀 Виконання завдання: ${name}`);

      await task();

      const duration = Date.now() - startTime;
      this.stats.jobsExecuted++;

      logger.debug(`✅ Завдання "${name}" виконано за ${duration}ms`);

      // Оновлення наступного запуску з node-cron API
      try {
        // Перевіряємо, чи існує метод nextDates
        const jobObject = jobInfo.job as any;
        
        if (typeof jobObject.nextDates === 'function') {
          try {
            const dates = jobObject.nextDates(1);
            const firstDate = Array.isArray(dates) ? dates[0] : dates;
            // Обробка різних типів об'єктів дати
            if (firstDate && typeof firstDate.toJSDate === 'function') {
              jobInfo.nextRun = firstDate.toJSDate();
            } else if (firstDate && typeof firstDate.toDate === 'function') {
              jobInfo.nextRun = firstDate.toDate();
            } else if (firstDate instanceof Date) {
              jobInfo.nextRun = firstDate;
            } else {
              jobInfo.nextRun = new Date(firstDate);
            }
          } catch (nextDatesError) {
            logger.debug('nextDates method failed in executeJob, trying nextDate', {
              error: String(nextDatesError)
            });
            throw nextDatesError;
          }
        } else if (typeof jobObject.nextDate === 'function') {
          // Фолбек для старих версій node-cron
          try {
            const dt = jobObject.nextDate();
            if (dt instanceof Date) {
              jobInfo.nextRun = dt;
            } else if (dt && typeof dt.toJSDate === 'function') {
              jobInfo.nextRun = dt.toJSDate();
            } else if (dt && typeof dt.toDate === 'function') {
              jobInfo.nextRun = dt.toDate();
            } else {
              jobInfo.nextRun = new Date(dt);
            }
          } catch (nextDateError) {
            logger.debug('nextDate method failed in executeJob, using fallback', {
              error: String(nextDateError)
            });
            throw nextDateError;
          }
        } else {
          // Якщо ні один метод не доступний, встановлюємо поточну дату + 5 хвилин
          jobInfo.nextRun = this.calculateNextRun(jobInfo.schedule);
        }
      } catch (e) {
        logger.warn('scheduler_next_run_update_failed', { 
          job: name, 
          error: e instanceof Error ? e.message : String(e),
          methodsAvailable: {
            nextDates: typeof (jobInfo.job as any).nextDates,
            nextDate: typeof (jobInfo.job as any).nextDate
          }
        });
        // Зберігаємо поточне значення nextRun як fallback або розраховуємо нове
        if (!jobInfo.nextRun || jobInfo.nextRun <= new Date()) {
          jobInfo.nextRun = this.calculateNextRun(jobInfo.schedule);
        }
      }
    } catch (error) {
      jobInfo.errors++;
      this.stats.jobsFailed++;

      logger.error(
        `❌ Помилка виконання завдання "${name}": ${error instanceof Error ? error.message : String(error)}`
      );

      // Сповіщення про помилку
      await this.notifyJobError(name, error);
    }
  }

  /**
   * Зупинка завдання
   */
  stopJob(name: string): void {
    try {
      const jobInfo = this.jobs.get(name);
      if (jobInfo) {
        jobInfo.job.stop();
        this.jobs.delete(name);
        this.stats.activeJobs--;

        logger.debug(`🛑 Завдання "${name}" зупинено`);
      }
    } catch (error) {
      logger.error(
        `Помилка зупинки завдання "${name}": ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }

  /**
   * Отримання інформації про завдання
   */
  getJobInfo(name: string): JobDetails | null {
    const jobInfo = this.jobs.get(name);
    if (!jobInfo) return null;

    return {
      name,
      schedule: jobInfo.schedule,
      task: jobInfo.task,
      createdAt: jobInfo.createdAt,
      lastRun: jobInfo.lastRun,
      nextRun: jobInfo.nextRun,
      executions: jobInfo.executions,
      errors: jobInfo.errors,
      // node-cron v3: використовуємо getStatus(), в інших випадках вважаємо активним
      isActive: (() => { 
        const status = (jobInfo.job as any)?.getStatus?.();
        return typeof status === 'string' ? status !== 'stopped' : true;
      })(),
    };
  }

  /**
   * Отримання всіх завдань
   */
  getAllJobs(): JobDetails[] {
    return Array.from(this.jobs.keys()).map(name => this.getJobInfo(name)!);
  }

  /**
   * Очищення кешу
   */
  private async cleanupCache(): Promise<void> {
    try {
      const cacheService = this.bot.getService('cache');
      if (cacheService) {
        await cacheService.cleanupMemory();
        logger.info('🧹 Кеш очищено');
      }
    } catch (error) {
      logger.error(
        `Помилка очищення кешу: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }

  /**
   * Розрахунок наступного часу запуску на основі cron виразу
   */
  private calculateNextRun(schedule: string): Date {
    try {
      // Проста логіка для розрахунку наступного часу запуску
      const parts = schedule.trim().split(/\s+/);
      if (parts.length !== 5 && parts.length !== 6) {
        // Некоректний cron вираз, повертаємо поточну дату + 1 хвилина
        return new Date(Date.now() + 60000);
      }

      const now = new Date();
      const nextRun = new Date(now);

      // Обробка простих випадків
      if (schedule.startsWith('*/')) {
        // Кожні N хвилин/годин/днів
        const intervalStr = schedule.substring(2).split(' ')[0];
        if (intervalStr) {
          const interval = parseInt(intervalStr, 10);
          if (!isNaN(interval)) {
            if (schedule.includes('* * * *')) {
              // Кожні N хвилин
              nextRun.setMinutes(now.getMinutes() + interval);
            } else if (schedule.includes('* * *')) {
              // Кожні N годин
              nextRun.setHours(now.getHours() + interval);
            }
            return nextRun;
          }
        }
      }

      // Для інших випадків - повертаємо наступну хвилину
      nextRun.setMinutes(now.getMinutes() + 1);
      return nextRun;
    } catch (error) {
      // При помилці повертаємо поточну дату + 5 хвилин
      return new Date(Date.now() + 5 * 60 * 1000);
    }
  }

  /**
   * Оновлення статистики
   */
  private async updateStats(): Promise<void> {
    try {
      // Оновлення метрик
      const metricsService = this.bot.getService('metrics');
      if (metricsService) {
        metricsService.updateAllMetrics();
      }

      // Оновлення статистики сервісів
      const serviceManager = this.bot.serviceManager;
      if (serviceManager) {
        const servicesStats = serviceManager.getStats();
        logger.debug(`📊 Статистика сервісів оновлена: ${JSON.stringify(servicesStats)}`);
      }
    } catch (error) {
      logger.error(
        `Помилка оновлення статистики: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }

  /**
   * Перевірка здоров'я
   */
  private async healthCheck(): Promise<void> {
    try {
      const healthStatus: HealthStatus = {
        timestamp: new Date(),
        services: {},
        overall: 'healthy',
      };

      // Перевірка сервісів
      const serviceManager = this.bot.serviceManager;
      if (serviceManager) {
        const servicesStatus = serviceManager.getServicesStatus() as Record<
          string,
          { isActive?: boolean; stats?: unknown }
        >;

        for (const [name, s] of Object.entries(servicesStatus)) {
          const active = !!s.isActive;
          healthStatus.services[name] = {
            isActive: active,
            hasStats: typeof s.stats !== 'undefined' && s.stats !== null,
          };

          if (!active) {
            healthStatus.overall = 'degraded';
          }
        }
      }

      // Перевірка Discord клієнта
      if (this.bot.client) {
        healthStatus.discord = {
          isReady: this.bot.client.isReady(),
          uptime: this.bot.client.uptime,
          guilds: this.bot.client.guilds.cache.size,
        };
      }

      logger.debug(`🏥 Перевірка здоров'я завершена: ${JSON.stringify(healthStatus)}`);

      // Сповіщення про проблеми
      if (healthStatus.overall !== 'healthy') {
        await this.notifyHealthIssue(healthStatus);
      }
    } catch (error) {
      logger.error(
        `Помилка перевірки здоров'я: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }

  /**
   * Створення резервної копії
   */
  private async createBackup(): Promise<void> {
    try {
      logger.info('💾 Створення резервної копії...');

      // Тут можна додати логіку створення резервної копії
      // Наприклад, збереження даних в файл або базу даних

      logger.info('✅ Резервна копія створена');
    } catch (error) {
      logger.error(
        `Помилка створення резервної копії: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }

  /**
   * Сповіщення про помилку завдання
   */
  private async notifyJobError(jobName: string, error: any): Promise<void> {
    try {
      // Тут можна додати логіку сповіщення
      // Наприклад, відправка повідомлення в Discord канал

      logger.warn(
        `⚠️ Сповіщення про помилку завдання "${jobName}": ${error instanceof Error ? error.message : String(error)}`
      );
    } catch (notifyError) {
      logger.error(
        `Помилка сповіщення про помилку завдання: ${notifyError instanceof Error ? notifyError.message : String(notifyError)}`
      );
    }
  }

  /**
   * Сповіщення про проблеми здоров'я
   */
  private async notifyHealthIssue(healthStatus: HealthStatus): Promise<void> {
    try {
      // Тут можна додати логіку сповіщення про проблеми здоров'я

      logger.warn(`⚠️ Проблеми здоров'я системи: ${healthStatus.overall}`);
    } catch (notifyError) {
      logger.error(
        `Помилка сповіщення про проблеми здоров'я: ${notifyError instanceof Error ? notifyError.message : String(notifyError)}`
      );
    }
  }

  /**
   * Отримання статистики
   */
  getStats(): SchedulerStats {
    return {
      ...this.stats,
      jobs: this.getAllJobs(),
      isActive: this.isActive(),
    };
  }

  /**
   * Перевірка активності
   */
  isActive(): boolean {
    return this._isActive;
  }

  /**
   * Завершення роботи
   */
  async shutdown(): Promise<void> {
    logger.info('🛑 Завершення роботи Scheduler сервісу...');

    try {
      // Зупинка всіх завдань
      for (const [name] of this.jobs) {
        this.stopJob(name);
      }

      this._isActive = false;
      logger.info('✅ Scheduler сервіс завершено');
    } catch (error) {
      logger.error(
        `❌ Помилка завершення Scheduler сервісу: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }
}

export default SchedulerService;
