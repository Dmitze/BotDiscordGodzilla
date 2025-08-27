import { BaseService } from '@/core/BaseService';
import type { BotConfig } from '@/types';
import type { GoogleService } from '@/services/GoogleService';
import type { DriveFile } from '@/types/drive';
import type { SchedulerService } from '@/services/SchedulerService';
import type { SmartDocumentClassifier } from '@/services/SmartDocumentClassifier';
import type { DocumentAnalyticsService } from '@/services/DocumentAnalyticsService';
import type { ClassifiedDocument } from '@/services/SmartDocumentClassifier';
import logger from '@/utils/logger';
import queueManager from '@/utils/queueManager';

export interface DocumentTrigger {
  id: string;
  folderId: string;
  folderName: string;
  channelId: string;
  enabled: boolean;
  conditions: DocumentCondition[];
  actions: DocumentAction[];
  usersToNotify: string[]; // Discord user IDs
  createdAt: Date;
  lastRun?: Date;
  // New properties for enhanced functionality
  autoTaggingConfig?: AutoTaggingConfig;
  notificationTemplate?: NotificationTemplate;
}

// New interfaces for enhanced functionality
export interface AutoTaggingConfig {
  enabled: boolean;
  useAI: boolean;
  keywordThreshold: number; // 0-1 confidence threshold
  maxTags: number; // Maximum number of tags to generate
  customTags?: string[]; // Additional tags to always include
}

export interface NotificationTemplate {
  title: string;
  message: string;
  includeFileInfo: boolean;
  includeTags: boolean;
  includePreview: boolean;
  previewLength: number;
}

export interface DocumentCondition {
  type: 'fileType' | 'fileNamePattern' | 'fileSize' | 'createdDate' | 'modifiedDate';
  operator: 'equals' | 'contains' | 'startsWith' | 'endsWith' | 'greaterThan' | 'lessThan';
  value: string | number;
}

export interface DocumentAction {
  type: 'analyze' | 'classify' | 'tag' | 'notify' | 'export' | 'move' | 'delete' | 'summarize' | 'compare_versions';
  parameters?: Record<string, any>;
  // New property for background processing
  runInBackground?: boolean;
}

// New interface for document version comparison
export interface DocumentVersion {
  versionId: string;
  modifiedTime: string;
  lastModifyingUser?: string;
  size?: number;
  md5Checksum?: string;
}

// New interface for version comparison results
export interface VersionComparison {
  fileId: string;
  fileName: string;
  versions: DocumentVersion[];
  differences: {
    added: string[];
    removed: string[];
    modified: string[];
  };
  summary: string;
}

export interface ProcessedDocument {
  fileId: string;
  fileName: string;
  actionsTaken: string[];
  timestamp: Date;
  results?: any;
  // New properties for enhanced functionality
  autoTags?: string[];
  classification?: ClassifiedDocument;
  summary?: string; // Add summary property
  versionComparison?: VersionComparison; // Add version comparison property
}

export class AutomatedDocumentProcessor extends BaseService {
  private google: GoogleService | null = null;
  private scheduler: SchedulerService | null = null;
  private classifier: SmartDocumentClassifier | null = null;
  private analytics: DocumentAnalyticsService | null = null;
  private cache: CacheService | null = null; // Add cache service
  private triggers: DocumentTrigger[] = [];
  private processedDocuments: ProcessedDocument[] = [];
  private readonly MAX_PROCESSED_HISTORY = 1000;
  // New properties for enhanced functionality
  private readonly DEFAULT_AUTO_TAGGING_CONFIG: AutoTaggingConfig = {
    enabled: true,
    useAI: true,
    keywordThreshold: 0.3,
    maxTags: 10
  };
  private readonly DEFAULT_NOTIFICATION_TEMPLATE: NotificationTemplate = {
    title: 'Новий документ виявлено',
    message: 'Було знайдено новий документ, який відповідає вашим критеріям.',
    includeFileInfo: true,
    includeTags: true,
    includePreview: true,
    previewLength: 200
  };

  constructor(config: BotConfig) {
    super('AutomatedDocumentProcessor', config);
  }

  /**
   * Ініціалізує сервіс з необхідними залежностями
   */
  initializeServices(
    google: GoogleService,
    scheduler: SchedulerService,
    classifier: SmartDocumentClassifier,
    analytics: DocumentAnalyticsService,
    cache?: CacheService // Add cache parameter
  ): void {
    this.google = google;
    this.scheduler = scheduler;
    this.classifier = classifier;
    this.analytics = analytics;
    this.cache = cache || null; // Set cache service
    
    // Налаштовуємо регулярну перевірку нових документів
    if (this.scheduler) {
      this.scheduler.scheduleJob('auto-doc-processing', '*/10 * * * *', async () => {
        try {
          await this.processTriggers();
        } catch (error) {
          logger.error('Помилка автоматичної обробки документів', {
            component: 'AutomatedDocumentProcessor',
            error: error instanceof Error ? error.message : String(error)
          });
        }
      });
    }
  }

  /**
   * Додає новий тригер
   */
  addTrigger(trigger: Omit<DocumentTrigger, 'id' | 'createdAt'>): DocumentTrigger {
    const newTrigger: DocumentTrigger = {
      id: this.generateId(),
      ...trigger,
      createdAt: new Date()
    };
    
    this.triggers.push(newTrigger);
    
    logger.info('Додано новий тригер автоматичної обробки', {
      component: 'AutomatedDocumentProcessor',
      triggerId: newTrigger.id,
      folderId: newTrigger.folderId,
      conditions: newTrigger.conditions.length,
      actions: newTrigger.actions.length
    });
    
    return newTrigger;
  }

  /**
   * Оновлює існуючий тригер
   */
  updateTrigger(triggerId: string, updates: Partial<DocumentTrigger>): boolean {
    const index = this.triggers.findIndex(t => t.id === triggerId);
    
    if (index === -1) {
      return false;
    }
    
    this.triggers[index] = {
      ...this.triggers[index],
      ...updates
    };
    
    logger.info('Оновлено тригер автоматичної обробки', {
      component: 'AutomatedDocumentProcessor',
      triggerId
    });
    
    return true;
  }

  /**
   * Видаляє тригер
   */
  removeTrigger(triggerId: string): boolean {
    const initialLength = this.triggers.length;
    this.triggers = this.triggers.filter(t => t.id !== triggerId);
    
    const removed = this.triggers.length < initialLength;
    
    if (removed) {
      logger.info('Видалено тригер автоматичної обробки', {
        component: 'AutomatedDocumentProcessor',
        triggerId
      });
    }
    
    return removed;
  }

  /**
   * Отримує всі тригери
   */
  getTriggers(): DocumentTrigger[] {
    return [...this.triggers];
  }

  /**
   * Обробляє всі активні тригери
   */
  private async processTriggers(): Promise<void> {
    logger.debug('Початок автоматичної обробки документів', {
      component: 'AutomatedDocumentProcessor',
      activeTriggers: this.triggers.filter(t => t.enabled).length
    });

    for (const trigger of this.triggers.filter(t => t.enabled)) {
      try {
        await this.processTrigger(trigger);
        trigger.lastRun = new Date();
      } catch (error) {
        logger.error('Помилка обробки тригера', {
          component: 'AutomatedDocumentProcessor',
          triggerId: trigger.id,
          error: error instanceof Error ? error.message : String(error)
        });
      }
    }
  }

  /**
   * Обробляє окремий тригер
   */
  private async processTrigger(trigger: DocumentTrigger): Promise<void> {
    if (!this.google) {
      throw new Error('GoogleService не ініціалізовано');
    }

    logger.debug('Обробка тригера', {
      component: 'AutomatedDocumentProcessor',
      triggerId: trigger.id,
      folderId: trigger.folderId
    });

    // Отримуємо список файлів у папці
    const result = await this.google.listDriveFiles({
      folderId: trigger.folderId,
      pageSize: 100 // Отримуємо останні 100 файлів
    });

    // Фільтруємо файли за умовами тригера
    const filteredFiles = result.files.filter(file => 
      this.matchesConditions(file, trigger.conditions)
    );

    logger.debug('Відфільтровані файли', {
      component: 'AutomatedDocumentProcessor',
      triggerId: trigger.id,
      totalFiles: result.files.length,
      matchingFiles: filteredFiles.length
    });

    // Обробляємо кожен відповідний файл
    for (const file of filteredFiles) {
      // Check if any action should run in background
      const hasBackgroundActions = trigger.actions.some(action => action.runInBackground);
      
      if (hasBackgroundActions) {
        // Add to background queue
        queueManager.addJob('normal', {
          type: 'file_operation',
          data: {
            operation: 'process_file_background',
            fileId: file.id,
            triggerId: trigger.id
          },
          handler: async () => {
            await this.processFileInBackground(file, trigger);
          }
        });
      } else {
        // Process synchronously
        await this.processFile(file, trigger);
      }
    }
  }

  /**
   * Обробляє файл у фоновому режимі
   */
  private async processFileInBackground(file: DriveFile, trigger: DocumentTrigger): Promise<void> {
    try {
      logger.info('Початок фонової обробки файлу', {
        component: 'AutomatedDocumentProcessor',
        fileId: file.id,
        fileName: file.name,
        triggerId: trigger.id
      });

      await this.processFile(file, trigger);

      logger.info('Фонова обробка файлу завершена', {
        component: 'AutomatedDocumentProcessor',
        fileId: file.id,
        fileName: file.name,
        triggerId: trigger.id
      });
    } catch (error) {
      logger.error('Помилка фонової обробки файлу', {
        component: 'AutomatedDocumentProcessor',
        fileId: file.id,
        fileName: file.name,
        triggerId: trigger.id,
        error: error instanceof Error ? error.message : String(error)
      });
    }
  }

  /**
   * Перевіряє чи файл відповідає умовам
   */
  private matchesConditions(file: DriveFile, conditions: DocumentCondition[]): boolean {
    for (const condition of conditions) {
      let matches = false;
      
      switch (condition.type) {
        case 'fileType':
          if (file.mimeType) {
            matches = this.evaluateCondition(file.mimeType, condition.operator, condition.value);
          }
          break;
          
        case 'fileNamePattern':
          if (file.name) {
            matches = this.evaluateCondition(file.name, condition.operator, condition.value);
          }
          break;
          
        case 'fileSize':
          if (typeof file.size === 'number') {
            matches = this.evaluateCondition(file.size, condition.operator, condition.value);
          }
          break;
          
        case 'createdDate':
        case 'modifiedDate':
          const dateValue = condition.type === 'createdDate' ? file.createdTime : file.modifiedTime;
          if (dateValue) {
            const date = new Date(dateValue);
            matches = this.evaluateCondition(date, condition.operator, condition.value);
          }
          break;
      }
      
      // Якщо хоч одна умова не виконується, файл не підходить
      if (!matches) {
        return false;
      }
    }
    
    return true;
  }

  /**
   * Оцінює умову
   */
  private evaluateCondition(
    actualValue: string | number | Date,
    operator: string,
    expectedValue: string | number
  ): boolean {
    try {
      switch (operator) {
        case 'equals':
          return String(actualValue) === String(expectedValue);
          
        case 'contains':
          return String(actualValue).includes(String(expectedValue));
          
        case 'startsWith':
          return String(actualValue).startsWith(String(expectedValue));
          
        case 'endsWith':
          return String(actualValue).endsWith(String(expectedValue));
          
        case 'greaterThan':
          if (actualValue instanceof Date) {
            const expectedDate = new Date(expectedValue as string);
            return actualValue.getTime() > expectedDate.getTime();
          }
          return Number(actualValue) > Number(expectedValue);
          
        case 'lessThan':
          if (actualValue instanceof Date) {
            const expectedDate = new Date(expectedValue as string);
            return actualValue.getTime() < expectedDate.getTime();
          }
          return Number(actualValue) < Number(expectedValue);
          
        default:
          return false;
      }
    } catch (error) {
      logger.warn('Помилка оцінки умови', {
        component: 'AutomatedDocumentProcessor',
        operator,
        error: error instanceof Error ? error.message : String(error)
      });
      return false;
    }
  }

  /**
   * Аналізує документ
   */
  private async analyzeDocument(file: DriveFile): Promise<any> {
    // У реальній реалізації тут буде аналіз вмісту документа
    logger.debug('Аналіз документа', {
      component: 'AutomatedDocumentProcessor',
      fileId: file.id,
      fileName: file.name
    });
    
    // Get document content for analysis
    let content = '';
    if (this.google) {
      try {
        const result = await this.google.extractTextForChat(file.id);
        content = result.text;
      } catch (error) {
        logger.warn('Не вдалося отримати вміст документа для аналізу', {
          component: 'AutomatedDocumentProcessor',
          fileId: file.id,
          error: error instanceof Error ? error.message : String(error)
        });
      }
    }
    
    return {
      summary: 'Документ проаналізовано',
      wordCount: content.split(/\s+/).length,
      language: this.detectLanguage(content),
      keyPhrases: this.extractKeyPhrases(content)
    };
  }

  /**
   * Класифікує документ
   */
  private async classifyDocument(file: DriveFile): Promise<ClassifiedDocument | null> {
    if (!this.classifier) {
      logger.warn('SmartDocumentClassifier не ініціалізовано', {
        component: 'AutomatedDocumentProcessor',
        fileId: file.id
      });
      return null;
    }

    logger.debug('Класифікація документа', {
      component: 'AutomatedDocumentProcessor',
      fileId: file.id,
      fileName: file.name
    });
    
    try {
      // Get document content for classification
      let content = '';
      if (this.google) {
        try {
          const result = await this.google.extractTextForChat(file.id);
          content = result.text;
        } catch (error) {
          logger.warn('Не вдалося отримати вміст документа для класифікації', {
            component: 'AutomatedDocumentProcessor',
            fileId: file.id,
            error: error instanceof Error ? error.message : String(error)
          });
        }
      }
      
      // Classify the document
      const classified = await this.classifier.classifyDocument(file, content);
      return classified;
    } catch (error) {
      logger.error('Помилка класифікації документа', {
        component: 'AutomatedDocumentProcessor',
        fileId: file.id,
        error: error instanceof Error ? error.message : String(error)
      });
      return null;
    }
  }

  /**
   * Автоматично додає теги до документа на основі вмісту з кешуванням
   */
  private async autoTagDocument(file: DriveFile, trigger: DocumentTrigger): Promise<string[]> {
    const config = trigger.autoTaggingConfig || this.DEFAULT_AUTO_TAGGING_CONFIG;
    
    if (!config.enabled) {
      return [];
    }
    
    logger.debug('Автоматичне тегування документа', {
      component: 'AutomatedDocumentProcessor',
      fileId: file.id,
      fileName: file.name
    });
    
    try {
      // Спробуємо отримати теги з кешу
      const cacheKey = `doc:tags:${file.id}`;
      if (this.cache) {
        const cachedTags = await this.cache.get<string[]>(cacheKey);
        if (cachedTags) {
          logger.debug('Теги отримано з кешу', {
            component: 'AutomatedDocumentProcessor',
            fileId: file.id
          });
          return cachedTags;
        }
      }
      
      const tags: string[] = [];
      
      // Add custom tags if specified
      if (config.customTags) {
        tags.push(...config.customTags);
      }
      
      // Use AI-based tagging if enabled
      if (config.useAI && this.classifier) {
        try {
          // Get document content
          let content = '';
          if (this.google) {
            try {
              const result = await this.google.extractTextForChat(file.id);
              content = result.text;
            } catch (error) {
              logger.warn('Не вдалося отримати вміст документа для тегування', {
                component: 'AutomatedDocumentProcessor',
                fileId: file.id,
                error: error instanceof Error ? error.message : String(error)
              });
            }
          }
          
          // Classify document to get tags
          const classified = await this.classifier.classifyDocument(file, content);
          if (classified && classified.tags) {
            // Filter tags by confidence threshold
            const filteredTags = classified.tags.slice(0, config.maxTags);
            tags.push(...filteredTags);
          }
        } catch (error) {
          logger.error('Помилка автоматичного тегування документа', {
            component: 'AutomatedDocumentProcessor',
            fileId: file.id,
            error: error instanceof Error ? error.message : String(error)
          });
        }
      } else {
        // Use simple keyword-based tagging
        const fileNameKeywords = this.extractKeywords(file.name || '');
        tags.push(...fileNameKeywords.slice(0, config.maxTags));
      }
      
      // Remove duplicates and limit tags
      const uniqueTags = [...new Set(tags)].slice(0, config.maxTags);
      
      // Зберігаємо теги в кеш
      if (this.cache) {
        await this.cache.set(cacheKey, uniqueTags, 3600); // Кешуємо на 1 годину
      }
      
      logger.info('Додано автоматичні теги до документа', {
        component: 'AutomatedDocumentProcessor',
        fileId: file.id,
        tagCount: uniqueTags.length,
        tags: uniqueTags
      });
      
      return uniqueTags;
    } catch (error) {
      logger.error('Помилка автоматичного тегування документа', {
        component: 'AutomatedDocumentProcessor',
        fileId: file.id,
        error: error instanceof Error ? error.message : String(error)
      });
      return [];
    }
  }

  /**
   * Порівнює версії документа з кешуванням
   */
  private async compareDocumentVersions(file: DriveFile): Promise<VersionComparison> {
    logger.debug('Порівняння версій документа', {
      component: 'AutomatedDocumentProcessor',
      fileId: file.id,
      fileName: file.name
    });
    
    try {
      // Спробуємо отримати результат порівняння з кешу
      const cacheKey = `doc:version-comparison:${file.id}`;
      if (this.cache) {
        const cachedComparison = await this.cache.get<VersionComparison>(cacheKey);
        if (cachedComparison) {
          logger.debug('Порівняння версій отримано з кешу', {
            component: 'AutomatedDocumentProcessor',
            fileId: file.id
          });
          return cachedComparison;
        }
      }
      
      // Отримуємо історію версій документа
      // В реальній реалізації тут потрібно отримати інформацію про версії з Google Drive API
      // Для спрощення створюємо фіктивні версії
      const versions: DocumentVersion[] = [
        {
          versionId: '1',
          modifiedTime: new Date(Date.now() - 86400000).toISOString(), // 1 day ago
          lastModifyingUser: 'user1@example.com',
          size: 1024,
          md5Checksum: 'abc123'
        },
        {
          versionId: '2',
          modifiedTime: new Date().toISOString(),
          lastModifyingUser: 'user2@example.com',
          size: 1536,
          md5Checksum: 'def456'
        }
      ];
      
      // Порівнюємо версії (в реальній реалізації тут потрібно отримати вміст кожної версії та порівняти їх)
      // Для спрощення створюємо фіктивні результати порівняння
      const differences = {
        added: ['Новий розділ про безпеку', 'Додаткові рекомендації'],
        removed: ['Старі рекомендації'],
        modified: ['Оновлені вимоги до обладнання']
      };
      
      // Створюємо підсумок порівняння
      const summary = `Документ "${file.name}" має ${versions.length} версій. Знайдено ${differences.added.length} нових елементів, ${differences.removed.length} видалених елементів та ${differences.modified.length} змінених елементів.`;
      
      const comparison: VersionComparison = {
        fileId: file.id,
        fileName: file.name || 'Без назви',
        versions,
        differences,
        summary
      };
      
      // Зберігаємо результат порівняння в кеш
      if (this.cache) {
        await this.cache.set(cacheKey, comparison, 1800); // Кешуємо на 30 хвилин
      }
      
      logger.info('Порівняння версій документа завершено', {
        component: 'AutomatedDocumentProcessor',
        fileId: file.id,
        versionCount: versions.length
      });
      
      return comparison;
    } catch (error) {
      logger.error('Помилка порівняння версій документа', {
        component: 'AutomatedDocumentProcessor',
        fileId: file.id,
        error: error instanceof Error ? error.message : String(error)
      });
      
      // Повертаємо порожнє порівняння у разі помилки
      return {
        fileId: file.id,
        fileName: file.name || 'Без назви',
        versions: [],
        differences: {
          added: [],
          removed: [],
          modified: []
        },
        summary: 'Помилка при порівнянні версій документа'
      };
    }
  }

  /**
   * Автоматично створює резюме документа з кешуванням
   */
  private async summarizeDocument(file: DriveFile): Promise<string> {
    logger.debug('Автоматичне створення резюме документа', {
      component: 'AutomatedDocumentProcessor',
      fileId: file.id,
      fileName: file.name
    });
    
    try {
      // Спробуємо отримати резюме з кешу
      const cacheKey = `doc:summary:${file.id}`;
      if (this.cache) {
        const cachedSummary = await this.cache.get<string>(cacheKey);
        if (cachedSummary) {
          logger.debug('Резюме отримано з кешу', {
            component: 'AutomatedDocumentProcessor',
            fileId: file.id
          });
          return cachedSummary;
        }
      }
      
      // Get document content
      let content = '';
      if (this.google) {
        try {
          const result = await this.google.extractTextForChat(file.id);
          content = result.text;
        } catch (error) {
          logger.warn('Не вдалося отримати вміст документа для резюме', {
            component: 'AutomatedDocumentProcessor',
            fileId: file.id,
            error: error instanceof Error ? error.message : String(error)
          });
          return 'Не вдалося отримати вміст документа для резюме';
        }
      }
      
      // If we have an AI service, use it for summarization
      // For now, we'll use the existing summarizeTlDr function
      const { summarizeTlDr } = await import('@/utils/fileProcessor');
      const summary = summarizeTlDr(content, { budget: 500 });
      
      // Зберігаємо резюме в кеш
      if (this.cache) {
        await this.cache.set(cacheKey, summary, 3600); // Кешуємо на 1 годину
      }
      
      logger.info('Створено резюме документа', {
        component: 'AutomatedDocumentProcessor',
        fileId: file.id,
        summaryLength: summary.length
      });
      
      return summary;
    } catch (error) {
      logger.error('Помилка створення резюме документа', {
        component: 'AutomatedDocumentProcessor',
        fileId: file.id,
        error: error instanceof Error ? error.message : String(error)
      });
      return 'Помилка при створенні резюме документа';
    }
  }

  /**
   * Додає теги до документа
   */
  private async tagDocument(file: DriveFile, parameters?: Record<string, any>): Promise<any> {
    const tags = parameters?.tags || ['auto-processed'];
    
    // У реальній реалізації тут буде додавання тегів до документа
    logger.debug('Додавання тегів до документа', {
      component: 'AutomatedDocumentProcessor',
      fileId: file.id,
      fileName: file.name,
      tags
    });
    
    return {
      addedTags: tags
    };
  }

  /**
   * Надсилає сповіщення
   */
  private async sendNotification(
    file: DriveFile,
    trigger: DocumentTrigger,
    actionsTaken: string[],
    results: any,
    autoTags?: string[],
    classification?: ClassifiedDocument
  ): Promise<void> {
    // У реальній реалізації тут буде надсилання сповіщення в Discord
    logger.debug('Надсилання сповіщення про обробку документа', {
      component: 'AutomatedDocumentProcessor',
      fileId: file.id,
      fileName: file.name,
      channelId: trigger.channelId,
      usersToNotify: trigger.usersToNotify,
      actionsTaken
    });
    
    // Build notification message using template
    const template = trigger.notificationTemplate || this.DEFAULT_NOTIFICATION_TEMPLATE;
    
    let message = `**${template.title}**\n${template.message}`;
    
    if (template.includeFileInfo) {
      message += `\n\n📄 **${file.name || 'Без назви'}**`;
      if (file.mimeType) {
        message += `\n📎 Тип: ${this.getMimeTypeLabel(file.mimeType)}`;
      }
      if (file.size) {
        message += `\n⚖️ Розмір: ${this.formatFileSize(file.size)}`;
      }
    }
    
    if (template.includeTags && autoTags && autoTags.length > 0) {
      message += `\n\n🏷️ Теги: ${autoTags.map(tag => `\`${tag}\``).join(' ')}`;
    }
    
    if (template.includePreview && this.google) {
      try {
        const result = await this.google.extractTextForChat(file.id);
        const preview = result.text.substring(0, template.previewLength);
        message += `\n\n🔍 Попередній перегляд:\n\`\`\`${preview}${result.text.length > template.previewLength ? '...' : ''}\`\`\``;
      } catch (error) {
        logger.warn('Не вдалося отримати попередній перегляд документа', {
          component: 'AutomatedDocumentProcessor',
          fileId: file.id,
          error: error instanceof Error ? error.message : String(error)
        });
      }
    }
    
    message += `\n\n⚙️ Виконані дії: ${actionsTaken.map(action => `\`${action}\``).join(', ')}`;
    
    // In a real implementation, this would send a message to Discord
    // For now, we just log the message that would be sent
    logger.info('Підготовлене сповіщення', {
      component: 'AutomatedDocumentProcessor',
      channelId: trigger.channelId,
      message: message.substring(0, 100) + '...'
    });
  }

  /**
   * Обробляє окремий файл
   */
  private async processFile(file: DriveFile, trigger: DocumentTrigger): Promise<void> {
    try {
      // Перевіряємо чи файл вже був оброблений
      if (this.isFileProcessed(file.id)) {
        return;
      }

      logger.info('Обробка файлу тригером', {
        component: 'AutomatedDocumentProcessor',
        fileId: file.id,
        fileName: file.name,
        triggerId: trigger.id
      });

      const actionsTaken: string[] = [];
      let results: any = {};
      let autoTags: string[] = [];
      let classification: ClassifiedDocument | null = null;
      let summary: string = '';
      let versionComparison: VersionComparison | null = null;

      // Автоматичне тегування якщо увімкнено
      if (trigger.autoTaggingConfig?.enabled) {
        autoTags = await this.autoTagDocument(file, trigger);
        actionsTaken.push('auto-tag');
      }

      // Класифікація документа якщо потрібно
      if (trigger.actions.some(a => a.type === 'classify')) {
        classification = await this.classifyDocument(file);
        if (classification) {
          actionsTaken.push('classify');
          results.classification = classification;
        }
      }

      // Створення резюме документа якщо потрібно
      if (trigger.actions.some(a => a.type === 'summarize')) {
        summary = await this.summarizeDocument(file);
        actionsTaken.push('summarize');
        results.summary = summary;
      }

      // Порівняння версій документа якщо потрібно
      if (trigger.actions.some(a => a.type === 'compare_versions')) {
        versionComparison = await this.compareDocumentVersions(file);
        actionsTaken.push('compare_versions');
        results.versionComparison = versionComparison;
      }

      // Виконуємо всі дії тригера
      for (const action of trigger.actions) {
        try {
          const actionResult = await this.executeAction(file, action);
          actionsTaken.push(action.type);
          
          if (actionResult) {
            results[action.type] = actionResult;
          }
        } catch (error) {
          logger.error('Помилка виконання дії', {
            component: 'AutomatedDocumentProcessor',
            fileId: file.id,
            action: action.type,
            error: error instanceof Error ? error.message : String(error)
          });
        }
      }

      // Зберігаємо інформацію про оброблений файл
      this.recordProcessedFile({
        fileId: file.id,
        fileName: file.name || 'Без назви',
        actionsTaken,
        timestamp: new Date(),
        results,
        autoTags,
        classification: classification || undefined,
        summary: summary || undefined,
        versionComparison: versionComparison || undefined
      });

      // Надсилаємо сповіщення якщо потрібно
      if (trigger.actions.some(a => a.type === 'notify')) {
        await this.sendNotification(file, trigger, actionsTaken, results, autoTags, classification || undefined);
      }

      logger.info('Файл оброблено успішно', {
        component: 'AutomatedDocumentProcessor',
        fileId: file.id,
        fileName: file.name,
        actionsTaken: actionsTaken.length,
        tagCount: autoTags.length
      });
    } catch (error) {
      logger.error('Помилка обробки файлу', {
        component: 'AutomatedDocumentProcessor',
        fileId: file.id,
        fileName: file.name,
        error: error instanceof Error ? error.message : String(error)
      });
    }
  }

  /**
   * Виконує дію над файлом
   */
  private async executeAction(file: DriveFile, action: DocumentAction): Promise<any> {
    switch (action.type) {
      case 'analyze':
        return await this.analyzeDocument(file);
        
      case 'classify':
        return await this.classifyDocument(file);
        
      case 'tag':
        return await this.tagDocument(file, action.parameters);
        
      case 'export':
        return await this.exportDocument(file, action.parameters);
        
      case 'move':
        return await this.moveDocument(file, action.parameters);
        
      case 'delete':
        return await this.deleteDocument(file);
        
      case 'summarize':
        return await this.summarizeDocument(file);
        
      case 'compare_versions':
        return await this.compareDocumentVersions(file);
        
      default:
        logger.warn('Невідома дія', {
          component: 'AutomatedDocumentProcessor',
          action: action.type
        });
        return null;
    }
  }

  /**
   * Експортує документ
   */
  private async exportDocument(file: DriveFile, parameters?: Record<string, any>): Promise<any> {
    const format = parameters?.format || 'pdf';
    
    // У реальній реалізації тут буде експорт документа
    logger.debug('Експорт документа', {
      component: 'AutomatedDocumentProcessor',
      fileId: file.id,
      fileName: file.name,
      format
    });
    
    return {
      exportedFormat: format,
      exportPath: `/exports/${file.id}.${format}` // Потрібно реалізувати реальний експорт
    };
  }

  /**
   * Переміщує документ
   */
  private async moveDocument(file: DriveFile, parameters?: Record<string, any>): Promise<any> {
    const targetFolderId = parameters?.targetFolderId;
    
    if (!targetFolderId) {
      logger.warn('Не вказано цільову папку для переміщення', {
        component: 'AutomatedDocumentProcessor',
        fileId: file.id
      });
      return null;
    }

    // У реальній реалізації тут буде переміщення документа
    logger.debug('Переміщення документа', {
      component: 'AutomatedDocumentProcessor',
      fileId: file.id,
      fileName: file.name,
      targetFolderId
    });
    
    return {
      movedTo: targetFolderId
    };
  }

  /**
   * Видаляє документ
   */
  private async deleteDocument(file: DriveFile): Promise<any> {
    // У реальній реалізації тут буде видалення документа
    logger.debug('Видалення документа', {
      component: 'AutomatedDocumentProcessor',
      fileId: file.id,
      fileName: file.name
    });
    
    return {
      deleted: true
    };
  }

  /**
   * Перевіряє чи файл вже був оброблений
   */
  private isFileProcessed(fileId: string): boolean {
    return this.processedDocuments.some(pd => pd.fileId === fileId);
  }

  /**
   * Зберігає інформацію про оброблений файл
   */
  private recordProcessedFile(processed: ProcessedDocument): void {
    this.processedDocuments.push(processed);
    
    // Обмежуємо історію оброблених файлів
    if (this.processedDocuments.length > this.MAX_PROCESSED_HISTORY) {
      this.processedDocuments = this.processedDocuments.slice(-this.MAX_PROCESSED_HISTORY);
    }
  }

  /**
   * Генерує унікальний ID
   */
  private generateId(): string {
    return Date.now().toString(36) + Math.random().toString(36).substr(2, 5);
  }

  /**
   * Отримує історію оброблених документів
   */
  getProcessedDocuments(limit: number = 50): ProcessedDocument[] {
    return this.processedDocuments
      .sort((a, b) => b.timestamp.getTime() - a.timestamp.getTime())
      .slice(0, limit);
  }

  /**
   * Витягує ключові слова з тексту
   */
  private extractKeywords(text: string): string[] {
    // Simple keyword extraction - in a real implementation, this would be more sophisticated
    const stopWords = new Set(['і', 'в', 'на', 'з', 'до', 'та', 'що', 'як', 'по', 'за', 'не', 'то', 'це', 'він', 'вона', 'воно', 'ми', 'ви', 'вони']);
    const words = text.toLowerCase()
      .replace(/[^\p{L}\p{N}\s]/gu, ' ')
      .split(/\s+/)
      .filter(word => word.length > 2 && !stopWords.has(word));
    
    // Count word frequencies
    const wordCounts = new Map<string, number>();
    for (const word of words) {
      wordCounts.set(word, (wordCounts.get(word) || 0) + 1);
    }
    
    // Return top 10 words by frequency
    return Array.from(wordCounts.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10)
      .map(entry => entry[0]);
  }

  /**
   * Витягує ключові фрази з тексту
   */
  private extractKeyPhrases(text: string): string[] {
    // Simple key phrase extraction - in a real implementation, this would use NLP
    const sentences = text.split(/[.!?]+/).filter(s => s.trim().length > 0);
    const keyPhrases: string[] = [];
    
    for (const sentence of sentences) {
      const words = sentence.trim().split(/\s+/);
      if (words.length >= 3 && words.length <= 10) {
        keyPhrases.push(words.join(' '));
      }
    }
    
    return keyPhrases.slice(0, 5);
  }

  /**
   * Визначає мову тексту
   */
  private detectLanguage(text: string): string {
    // Simple language detection - in a real implementation, this would be more accurate
    const ukrainianChars = /[іїєґ]/gi;
    const russianChars = /[ыэъ]/gi;
    const englishChars = /[a-z]/gi;
    
    const ukrainianMatches = (text.match(ukrainianChars) || []).length;
    const russianMatches = (text.match(russianChars) || []).length;
    const englishMatches = (text.match(englishChars) || []).length;
    
    if (ukrainianMatches > russianMatches && ukrainianMatches > englishMatches) {
      return 'uk';
    }
    
    if (russianMatches > ukrainianMatches && russianMatches > englishMatches) {
      return 'ru';
    }
    
    if (englishMatches > ukrainianMatches && englishMatches > russianMatches) {
      return 'en';
    }
    
    return 'unknown';
  }

  /**
   * Отримує назву типу файлу
   */
  private getMimeTypeLabel(mimeType: string): string {
    const labelMap: Record<string, string> = {
      'application/pdf': 'PDF документ',
      'application/vnd.google-apps.document': 'Google Docs',
      'application/vnd.google-apps.spreadsheet': 'Google Sheets',
      'application/vnd.google-apps.presentation': 'Google Slides',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'Word документ',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'Excel таблиця',
      'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'PowerPoint презентація',
      'text/plain': 'Текстовий файл',
      'image/': 'Зображення'
    };
    
    for (const [key, label] of Object.entries(labelMap)) {
      if (mimeType.startsWith(key) || mimeType.includes(key)) {
        return label;
      }
    }
    
    return 'Файл';
  }

  /**
   * Форматує розмір файлу
   */
  private formatFileSize(bytes: number): string {
    if (bytes === 0) return '0 Bytes';
    
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  }

  // === BaseServiceClass required methods ===
  
  protected async onInitialize(): Promise<void> {
    logger.info('AutomatedDocumentProcessor ініціалізовано', {
      component: 'AutomatedDocumentProcessor'
    });
  }

  protected async onShutdown(): Promise<void> {
    logger.info('AutomatedDocumentProcessor зупинено', {
      component: 'AutomatedDocumentProcessor'
    });
  }

  protected async onHealthCheck(): Promise<any> {
    return {
      healthy: true,
      service: this.name,
      triggers: this.triggers.length,
      processedDocuments: this.processedDocuments.length
    };
  }

  protected onGetStats(): any {
    return {
      triggers: this.triggers.length,
      processedDocuments: this.processedDocuments.length
    };
  }
}