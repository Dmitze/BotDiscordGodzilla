import { BaseService } from '@/core/BaseService';
import type { BotConfig } from '@/types';
import type { GoogleService } from '@/services/GoogleService';
import type { DriveFile } from '@/types/drive';
import type SchedulerService from '@/services/SchedulerService';
import logger from '@/utils/logger';

// Define the interface for the Drive Changes Provider
export interface IDriveChangesProvider {
  getStartPageToken(): Promise<string>;
  listChanges(pageToken: string): Promise<{
    changes: Array<{
      removed?: boolean;
      fileId?: string;
      file?: DriveFile;
      time?: string;
    }>;
    nextPageToken?: string;
    newStartPageToken?: string;
  }>;
}

// Define the cache interface
interface ICache {
  get<T>(key: string): Promise<T | undefined>;
  set(key: string, val: any, ttlSec?: number): Promise<void>;
}

export interface ChangeNotification {
  fileId: string;
  fileName: string;
  changeType: 'created' | 'modified' | 'deleted' | 'shared' | 'version_added' | 'access_changed';
  timestamp: Date;
  userId?: string;
  details?: any;
}

export interface DriveChangeEvent {
  fileId: string;
  type: 'created' | 'modified' | 'removed';
  fileName: string;
  mimeType?: string | undefined;
  webViewLink?: string | undefined;
  modifiedTime?: string | undefined;
  owners?: Array<{ emailAddress: string }> | undefined;
}

export interface WatchedFolder {
  folderId: string;
  folderName: string;
  channelId: string;
  lastChecked: Date;
  usersToNotify: string[]; // Discord user IDs
}

// New interface for version history
export interface FileVersion {
  id: string;
  fileId: string;
  versionId: string;
  modifiedTime: Date;
  lastModifyingUser?: string;
  size?: number;
  md5Checksum?: string;
}

// New interface for access monitoring
export interface FileAccessInfo {
  fileId: string;
  userId: string;
  accessType: 'owner' | 'writer' | 'reader' | 'commenter';
  timestamp: Date;
  grantedBy?: string;
}

export class DriveChangesService extends BaseService {
  private google: GoogleService | null = null;
  private scheduler: SchedulerService | null = null;
  private watchedFolders: WatchedFolder[] = [];
  private changeHistory: Map<string, ChangeNotification[]> = new Map();
  private provider: IDriveChangesProvider | null = null;
  private cache: ICache | null = null;
  // New properties for enhanced functionality
  private versionHistory: Map<string, FileVersion[]> = new Map();
  private accessHistory: Map<string, FileAccessInfo[]> = new Map();
  private readonly CHANGE_HISTORY_LIMIT = 100;
  private readonly VERSION_HISTORY_LIMIT = 50;
  private readonly ACCESS_HISTORY_LIMIT = 100;

  constructor(config: BotConfig, provider?: IDriveChangesProvider, cache?: ICache) {
    super('DriveChangesService', config);
    this.provider = provider || null;
    this.cache = cache || null;
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
      
      // Schedule version history check every 30 minutes
      this.scheduler.scheduleJob('drive-version-check', '*/30 * * * *', async () => {
        try {
          await this.checkFileVersionHistory();
        } catch (error) {
          logger.error('Помилка перевірки історії версій у Drive', {
            component: 'DriveChangesService',
            error: error instanceof Error ? error.message : String(error)
          });
        }
      });
      
      // Schedule access monitoring check every hour
      this.scheduler.scheduleJob('drive-access-check', '0 * * * *', async () => {
        try {
          await this.checkFileAccessChanges();
        } catch (error) {
          logger.error('Помилка перевірки доступу до файлів у Drive', {
            component: 'DriveChangesService',
            error: error instanceof Error ? error.message : String(error)
          });
        }
      });
    }
  }

  /**
   * Ініціалізує сервіс для відстеження змін через Google Drive API
   */
  override async initialize(): Promise<void> {
    if (!this.provider || !this.cache) {
      logger.warn('DriveChangesService: provider or cache not available for initialization', {
        component: 'DriveChangesService'
      });
      return;
    }

    try {
      // Check if we have a stored token
      const storedToken = await this.cache.get<string>('drive:changes:startPageToken');
      
      if (!storedToken) {
        // Get start page token from provider
        // Add retry logic in case Google Service is not yet fully initialized
        let startToken: string | null = null;
        let retries = 0;
        const maxRetries = 5;
        const retryDelay = 2000; // 2 seconds
        
        while (retries < maxRetries && !startToken) {
          try {
            startToken = await this.provider.getStartPageToken();
          } catch (error) {
            retries++;
            if (retries >= maxRetries) {
              throw error;
            }
            logger.warn(`Failed to get start page token, retry ${retries}/${maxRetries} in ${retryDelay}ms`, {
              component: 'DriveChangesService',
              error: error instanceof Error ? error.message : String(error)
            });
            await new Promise(resolve => setTimeout(resolve, retryDelay));
          }
        }
        
        if (startToken) {
          await this.cache.set('drive:changes:startPageToken', startToken);
          logger.info('DriveChangesService: initialized with start page token', {
            component: 'DriveChangesService',
            startToken
          });
        }
      } else {
        logger.info('DriveChangesService: using existing start page token', {
          component: 'DriveChangesService',
          storedToken
        });
      }
    } catch (error) {
      logger.error('DriveChangesService: error during initialization', {
        component: 'DriveChangesService',
        error: error instanceof Error ? error.message : String(error)
      });
      throw error;
    }
  }

  /**
   * Опитує Google Drive API для отримання змін
   */
  async pollOnce(): Promise<{ events: DriveChangeEvent[]; newToken: string }> {
    if (!this.provider || !this.cache) {
      throw new Error('DriveChangesService not properly initialized with provider and cache');
    }

    // Get current token from cache
    const currentToken = await this.cache.get<string>('drive:changes:startPageToken');
    
    if (!currentToken) {
      throw new Error('No start page token found in cache');
    }

    // Process all pages of changes
    let pageToken = currentToken;
    let allChanges: Array<{
      removed?: boolean;
      fileId?: string;
      file?: DriveFile;
      time?: string;
    }> = [];
    let finalNewStartPageToken: string | undefined = undefined;

    do {
      try {
        // Get changes from provider
        const { changes, nextPageToken, newStartPageToken } = await this.provider.listChanges(pageToken);
        
        // Add changes to our collection
        allChanges = allChanges.concat(changes);
        
        // Keep the newStartPageToken from this page
        if (newStartPageToken) {
          finalNewStartPageToken = newStartPageToken;
        }
        
        // Move to next page if available
        pageToken = nextPageToken || '';
      } catch (error) {
        // If we get an error about Google Drive client not being initialized,
        // we should retry after a short delay
        if (error instanceof Error && error.message.includes('Google Drive client not initialized')) {
          logger.warn('Google Drive client not initialized during poll, retrying in 2 seconds...', {
            component: 'DriveChangesService'
          });
          await new Promise(resolve => setTimeout(resolve, 2000));
          // Retry the same page token
          continue;
        }
        // For other errors, re-throw
        throw error;
      }
    } while (pageToken);

    // Filter and map changes to events
    const events: DriveChangeEvent[] = allChanges
      .filter(change => {
        // Filter out changes not in our watched folder
        if (change.removed && change.fileId) {
          return true; // Removed files don't have file info, but we still want to report them
        }
        if (change.file && change.file.parents) {
          const folderId = this.config.drive?.folderId;
          return folderId ? change.file.parents.includes(folderId) : true;
        }
        return false;
      })
      .map(change => {
        if (change.removed && change.fileId) {
          return {
            fileId: change.fileId,
            type: 'removed' as const,
            fileName: `file-${change.fileId}`,
            modifiedTime: change.time
          };
        }
        
        if (change.file) {
          // For modified files, we'll determine if it's created based on whether we have previous info about it
          // Since we don't have createdTime in DriveFile, we'll use a heuristic based on modifiedTime
          // In a real implementation, we would check our database/cache for previous file info
          const isCreated = !change.file.modifiedTime; // If no modifiedTime, assume it's new
          
          return {
            fileId: change.file.id,
            type: isCreated ? 'created' : 'modified',
            fileName: change.file.name || 'Unnamed file',
            mimeType: change.file.mimeType,
            webViewLink: this.config.drive?.hideWebLink ? undefined : change.file.webViewLink,
            modifiedTime: change.file.modifiedTime,
            owners: change.file.owners ? change.file.owners.map(owner => ({ emailAddress: typeof owner === 'string' ? owner : (owner as any).emailAddress || '' })) : undefined
          };
        }
        
        // Fallback
        return {
          fileId: change.fileId || 'unknown',
          type: 'modified' as const,
          fileName: 'Unknown file'
        };
      });

    // Update token in cache if we have a new one
    let finalToken = currentToken;
    if (finalNewStartPageToken) {
      await this.cache.set('drive:changes:startPageToken', finalNewStartPageToken);
      finalToken = finalNewStartPageToken;
    }

    return { events, newToken: finalToken };
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
   * Перевіряє історію версій файлів
   */
  private async checkFileVersionHistory(): Promise<void> {
    if (!this.google) return;

    logger.debug('Перевірка історії версій файлів', {
      component: 'DriveChangesService'
    });

    try {
      // Отримуємо всі відстежувані файли
      for (const [fileId, _history] of this.changeHistory) {
        // Отримуємо інформацію про версії файлу
        const versions = await this.getFileVersions(fileId);
        
        // Перевіряємо чи є нові версії
        const previousVersions = this.versionHistory.get(fileId) || [];
        const newVersions = versions.filter(version => 
          !previousVersions.some(prev => prev.versionId === version.versionId)
        );
        
        // Якщо є нові версії, створюємо сповіщення
        for (const version of newVersions) {
          const notification: ChangeNotification = {
            fileId,
            fileName: version.fileId, // This would be looked up in a real implementation
            changeType: 'version_added',
            timestamp: version.modifiedTime,
            details: {
              versionId: version.versionId,
              size: version.size,
              lastModifyingUser: version.lastModifyingUser
            }
          };
          
          // Зберігаємо сповіщення
          this.saveNotification(fileId, notification);
          
          // Оновлюємо історію версій
          this.saveFileVersion(fileId, version);
        }
      }
    } catch (error) {
      logger.error('Помилка перевірки історії версій файлів', {
        component: 'DriveChangesService',
        error: error instanceof Error ? error.message : String(error)
      });
    }
  }

  /**
   * Перевіряє зміни в доступі до файлів
   */
  private async checkFileAccessChanges(): Promise<void> {
    if (!this.google) return;

    logger.debug('Перевірка змін в доступі до файлів', {
      component: 'DriveChangesService'
    });

    try {
      // Отримуємо всі відстежувані файли
      for (const [fileId, _history] of this.changeHistory) {
        // Отримуємо інформацію про доступ до файлу
        const accessInfo = await this.getFileAccessInfo(fileId);
        
        // Перевіряємо чи є зміни в доступі
        const previousAccess = this.accessHistory.get(fileId) || [];
        const newAccess = accessInfo.filter(info => 
          !previousAccess.some(prev => prev.userId === info.userId && prev.accessType === info.accessType)
        );
        
        // Якщо є зміни в доступі, створюємо сповіщення
        for (const access of newAccess) {
          const notification: ChangeNotification = {
            fileId,
            fileName: access.fileId, // This would be looked up in a real implementation
            changeType: 'access_changed',
            timestamp: access.timestamp,
            details: {
              userId: access.userId,
              accessType: access.accessType,
              grantedBy: access.grantedBy
            }
          };
          
          // Зберігаємо сповіщення
          this.saveNotification(fileId, notification);
          
          // Оновлюємо історію доступу
          this.saveFileAccess(fileId, access);
        }
      }
    } catch (error) {
      logger.error('Помилка перевірки змін в доступі до файлів', {
        component: 'DriveChangesService',
        error: error instanceof Error ? error.message : String(error)
      });
    }
  }

  /**
   * Отримує версії файлу
   */
  private async getFileVersions(_fileId: string): Promise<FileVersion[]> {
    // В реальній реалізації тут потрібно отримати інформацію про версії файлу з Google Drive API
    // Для спрощення повертаємо порожній масив
    return [];
  }

  /**
   * Отримує інформацію про доступ до файлу
   */
  private async getFileAccessInfo(_fileId: string): Promise<FileAccessInfo[]> {
    // В реальній реалізації тут потрібно отримати інформацію про доступ до файлу з Google Drive API
    // Для спрощення повертаємо порожній масив
    return [];
  }

  /**
   * Зберігає інформацію про версію файлу
   */
  private saveFileVersion(fileId: string, version: FileVersion): void {
    let versions = this.versionHistory.get(fileId) || [];
    versions.push(version);
    
    // Обмежуємо розмір історії версій
    if (versions.length > this.VERSION_HISTORY_LIMIT) {
      versions = versions.slice(-this.VERSION_HISTORY_LIMIT);
    }
    
    this.versionHistory.set(fileId, versions);
  }

  /**
   * Зберігає інформацію про доступ до файлу
   */
  private saveFileAccess(fileId: string, access: FileAccessInfo): void {
    let accessInfo = this.accessHistory.get(fileId) || [];
    accessInfo.push(access);
    
    // Обмежуємо розмір історії доступу
    if (accessInfo.length > this.ACCESS_HISTORY_LIMIT) {
      accessInfo = accessInfo.slice(-this.ACCESS_HISTORY_LIMIT);
    }
    
    this.accessHistory.set(fileId, accessInfo);
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
  private async getFilePreviousInfo(_fileId: string): Promise<any> {
    // В реальній реалізації тут потрібно отримати інформацію з бази даних або кешу
    // Для спрощення повертаємо null
    return null;
  }

  /**
   * Зберігає поточну інформацію про файл
   */
  private async saveFileCurrentInfo(_file: DriveFile): Promise<void> {
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
   * Отримує історію версій для файлу
   */
  getVersionHistory(fileId: string): FileVersion[] {
    return this.versionHistory.get(fileId) || [];
  }

  /**
   * Отримує історію доступу для файлу
   */
  getAccessHistory(fileId: string): FileAccessInfo[] {
    return this.accessHistory.get(fileId) || [];
  }

  /**
   * Отримує всі відстежувані папки
   */
  getWatchedFolders(): WatchedFolder[] {
    return [...this.watchedFolders];
  }

  // === BaseService required methods ===
  
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
      changeHistorySize: this.changeHistory.size,
      versionHistorySize: this.versionHistory.size,
      accessHistorySize: this.accessHistory.size
    };
  }

  protected onGetStats(): any {
    return {
      watchedFolders: this.watchedFolders.length,
      changeHistorySize: this.changeHistory.size,
      versionHistorySize: this.versionHistory.size,
      accessHistorySize: this.accessHistory.size
    };
  }
}