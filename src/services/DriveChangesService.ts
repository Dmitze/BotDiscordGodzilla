import { BaseServiceClass } from '@/core/BaseService';
import type { BotConfig } from '@/types';
import type { GoogleService } from '@/services/GoogleService';
import type { DriveFile } from '@/types/drive';
import type { SchedulerService } from '@/services/SchedulerService';
import logger from '@/utils/logger';

export interface ChangeNotification {
  fileId: string;
  fileName: string;
  changeType: 'created' | 'modified' | 'deleted' | 'shared';
  timestamp: Date;
  userId?: string;
  details?: any;
}

export interface WatchedFolder {
  folderId: string;
  folderName: string;
  channelId: string;
  lastChecked: Date;
  usersToNotify: string[]; // Discord user IDs
}

export class DriveChangesService extends BaseServiceClass {
  private google: GoogleService | null = null;
  private scheduler: SchedulerService | null = null;
  private watchedFolders: WatchedFolder[] = [];
  private changeHistory: Map<string, ChangeNotification[]> = new Map();
  private readonly CHANGE_HISTORY_LIMIT = 100;

  constructor(config: BotConfig) {
    super('DriveChangesService', config);
  }

  /**
   * Ініціалізує сервіс з необхідними залежностями
   */
  initializeServices(google: GoogleService, scheduler: SchedulerService): void {
    this.google = google;
    this.scheduler = scheduler;
    
    // Налаштовуємо регулярну перевірку змін
    if (this.scheduler) {
      this.scheduler.scheduleJob('drive-changes-check', '*/5 * * * *', async () => {
        try {
          await this.checkForChanges();
        } catch (error) {
          logger.error('Помилка перевірки змін у Drive', {
            component: 'DriveChangesService',
            error: error instanceof Error ? error.message : String(error)
          });
        }
      });
    }
  }

  /**
   * Додає папку до списку відстежуваних
   */
  async watchFolder(folderId: string, channelId: string, usersToNotify: string[]): Promise<void> {
    try {
      if (!this.google) {
        throw new Error('GoogleService не ініціалізовано');
      }

      // Отримуємо назву папки
      const folderMeta = await this.google.getDriveFile(folderId);
      const folderName = folderMeta.name || 'Без назви';

      // Додаємо до списку відстежуваних
      this.watchedFolders.push({
        folderId,
        folderName,
        channelId,
        lastChecked: new Date(),
        usersToNotify
      });

      logger.info('Папка додана до відстеження', {
        component: 'DriveChangesService',
        folderId,
        folderName,
        channelId
      });
    } catch (error) {
      logger.error('Помилка додавання папки до відстеження', {
        component: 'DriveChangesService',
        folderId,
        error: error instanceof Error ? error.message : String(error)
      });
      throw error;
    }
  }

  /**
   * Видаляє папку зі списку відстежуваних
   */
  unwatchFolder(folderId: string): void {
    this.watchedFolders = this.watchedFolders.filter(f => f.folderId !== folderId);
    
    logger.info('Папка видалена з відстеження', {
      component: 'DriveChangesService',
      folderId
    });
  }

  /**
   * Перевіряє наявність змін у відстежуваних папках
   */
  private async checkForChanges(): Promise<void> {
    if (!this.google) return;

    logger.debug('Перевірка змін у Drive', {
      component: 'DriveChangesService',
      watchedFolders: this.watchedFolders.length
    });

    for (const watchedFolder of this.watchedFolders) {
      try {
        // Отримуємо список файлів у папці
        const result = await this.google.listDriveFiles({
          folderId: watchedFolder.folderId,
          pageSize: 50 // Отримуємо останні 50 файлів
        });

        // Перевіряємо кожен файл на наявність змін
        for (const file of result.files) {
          await this.checkFileForChanges(file, watchedFolder);
        }

        // Оновлюємо час останньої перевірки
        watchedFolder.lastChecked = new Date();
      } catch (error) {
        logger.error('Помилка перевірки змін у папці', {
          component: 'DriveChangesService',
          folderId: watchedFolder.folderId,
          error: error instanceof Error ? error.message : String(error)
        });
      }
    }
  }

  /**
   * Перевіряє файл на наявність змін
   */
  private async checkFileForChanges(file: DriveFile, watchedFolder: WatchedFolder): Promise<void> {
    try {
      // Отримуємо попередню інформацію про файл якщо вона є
      const previousInfo = await this.getFilePreviousInfo(file.id);
      
      // Якщо файл новий
      if (!previousInfo) {
        await this.handleFileCreated(file, watchedFolder);
        return;
      }
      
      // Перевіряємо чи змінився файл
      const wasModified = this.isFileModified(file, previousInfo);
      if (wasModified) {
        await this.handleFileModified(file, watchedFolder);
        return;
      }
      
      // Зберігаємо поточну інформацію про файл
      await this.saveFileCurrentInfo(file);
    } catch (error) {
      logger.error('Помилка перевірки файлу на зміни', {
        component: 'DriveChangesService',
        fileId: file.id,
        error: error instanceof Error ? error.message : String(error)
      });
    }
  }

  /**
   * Обробляє створення нового файлу
   */
  private async handleFileCreated(file: DriveFile, watchedFolder: WatchedFolder): Promise<void> {
    // Зберігаємо інформацію про файл
    await this.saveFileCurrentInfo(file);
    
    // Створюємо сповіщення
    const notification: ChangeNotification = {
      fileId: file.id,
      fileName: file.name || 'Без назви',
      changeType: 'created',
      timestamp: new Date(),
      details: {
        mimeType: file.mimeType,
        size: file.size
      }
    };
    
    // Зберігаємо сповіщення
    this.saveNotification(file.id, notification);
    
    // Відправляємо сповіщення (в реалізації потрібно інтегрувати з Discord)
    await this.sendNotification(notification, watchedFolder);
    
    logger.info('Новий файл створено', {
      component: 'DriveChangesService',
      fileId: file.id,
      fileName: file.name
    });
  }

  /**
   * Обробляє зміну файлу
   */
  private async handleFileModified(file: DriveFile, watchedFolder: WatchedFolder): Promise<void> {
    // Оновлюємо інформацію про файл
    await this.saveFileCurrentInfo(file);
    
    // Створюємо сповіщення
    const notification: ChangeNotification = {
      fileId: file.id,
      fileName: file.name || 'Без назви',
      changeType: 'modified',
      timestamp: new Date(),
      details: {
        mimeType: file.mimeType,
        size: file.size
      }
    };
    
    // Зберігаємо сповіщення
    this.saveNotification(file.id, notification);
    
    // Відправляємо сповіщення
    await this.sendNotification(notification, watchedFolder);
    
    logger.info('Файл змінено', {
      component: 'DriveChangesService',
      fileId: file.id,
      fileName: file.name
    });
  }

  /**
   * Відправляє сповіщення про зміни
   */
  private async sendNotification(notification: ChangeNotification, watchedFolder: WatchedFolder): Promise<void> {
    // В реальній реалізації тут потрібно відправити повідомлення в Discord канал
    // Це потребує інтеграції з Discord клієнтом
    
    logger.debug('Сповіщення про зміни у Drive', {
      component: 'DriveChangesService',
      notification,
      channelId: watchedFolder.channelId,
      usersToNotify: watchedFolder.usersToNotify
    });
  }

  /**
   * Перевіряє чи файл було змінено
   */
  private isFileModified(current: DriveFile, previous: any): boolean {
    // Перевіряємо дату зміни
    if (current.modifiedTime && previous.modifiedTime) {
      const currentModified = new Date(current.modifiedTime);
      const previousModified = new Date(previous.modifiedTime);
      return currentModified.getTime() > previousModified.getTime();
    }
    
    // Перевіряємо розмір
    if (typeof current.size === 'number' && typeof previous.size === 'number') {
      return current.size !== previous.size;
    }
    
    return false;
  }

  /**
   * Отримує попередню інформацію про файл
   */
  private async getFilePreviousInfo(fileId: string): Promise<any> {
    // В реальній реалізації тут потрібно отримати інформацію з бази даних або кешу
    // Для спрощення повертаємо null
    return null;
  }

  /**
   * Зберігає поточну інформацію про файл
   */
  private async saveFileCurrentInfo(file: DriveFile): Promise<void> {
    // В реальній реалізації тут потрібно зберегти інформацію в базу даних або кеш
    // Для спрощення нічого не робимо
  }

  /**
   * Зберігає сповіщення
   */
  private saveNotification(fileId: string, notification: ChangeNotification): void {
    // Отримуємо існуючу історію для файлу
    let history = this.changeHistory.get(fileId) || [];
    
    // Додаємо нове сповіщення
    history.push(notification);
    
    // Обмежуємо розмір історії
    if (history.length > this.CHANGE_HISTORY_LIMIT) {
      history = history.slice(-this.CHANGE_HISTORY_LIMIT);
    }
    
    // Зберігаємо оновлену історію
    this.changeHistory.set(fileId, history);
  }

  /**
   * Отримує історію змін для файлу
   */
  getChangeHistory(fileId: string): ChangeNotification[] {
    return this.changeHistory.get(fileId) || [];
  }

  /**
   * Отримує всі відстежувані папки
   */
  getWatchedFolders(): WatchedFolder[] {
    return [...this.watchedFolders];
  }

  // === BaseServiceClass required methods ===
  
  protected async onInitialize(): Promise<void> {
    logger.info('DriveChangesService ініціалізовано', {
      component: 'DriveChangesService'
    });
  }

  protected async onShutdown(): Promise<void> {
    logger.info('DriveChangesService зупинено', {
      component: 'DriveChangesService'
    });
  }

  protected async onHealthCheck(): Promise<any> {
    return {
      healthy: true,
      service: this.name,
      watchedFolders: this.watchedFolders.length,
      changeHistorySize: this.changeHistory.size
    };
  }

  protected onGetStats(): any {
    return {
      watchedFolders: this.watchedFolders.length,
      changeHistorySize: this.changeHistory.size
    };
  }
}