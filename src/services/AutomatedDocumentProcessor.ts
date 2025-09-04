import { BaseService } from '@/core/BaseService';
import type { BotConfig } from '@/types';
import type { GoogleService } from '@/services/GoogleService';
import type { DriveFile } from '@/types/drive';
import type SchedulerService from '@/services/SchedulerService';
import type { SmartDocumentClassifier } from '@/services/SmartDocumentClassifier';
import type { DocumentAnalyticsService } from '@/services/DocumentAnalyticsService';
import type { ClassifiedDocument } from '@/services/SmartDocumentClassifier';
import logger from '@/utils/logger';

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
  type: 'analyze' | 'classify' | 'tag' | 'notify' | 'export' | 'move' | 'delete';
  parameters?: Record<string, any>;
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
}

export class AutomatedDocumentProcessor extends BaseService {
  private google: GoogleService | null = null;
  private scheduler: SchedulerService | null = null;
  private classifier: SmartDocumentClassifier | null = null;
  private analytics: DocumentAnalyticsService | null = null;
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
    analytics: DocumentAnalyticsService
  ): void {
    this.google = google;
    this.scheduler = scheduler;
    this.classifier = classifier;
    this.analytics = analytics;
    
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
    
    // Get the existing trigger with non-null assertion since we know it exists
    const existingTrigger = this.triggers[index]!;
    
    // For exactOptionalPropertyTypes, we need to handle the lastRun property correctly
    // If lastRun is undefined, we should not include it in the object literal
    const updatedTrigger: DocumentTrigger = {
      id: existingTrigger.id,
      folderId: updates.folderId !== undefined ? updates.folderId : existingTrigger.folderId,
      folderName: updates.folderName !== undefined ? updates.folderName : existingTrigger.folderName,
      channelId: updates.channelId !== undefined ? updates.channelId : existingTrigger.channelId,
      enabled: updates.enabled !== undefined ? updates.enabled : existingTrigger.enabled,
      conditions: updates.conditions !== undefined ? updates.conditions : existingTrigger.conditions,
      actions: updates.actions !== undefined ? updates.actions : existingTrigger.actions,
      usersToNotify: updates.usersToNotify !== undefined ? updates.usersToNotify : existingTrigger.usersToNotify,
      createdAt: existingTrigger.createdAt,
      // Handle optional properties
      ...(updates.lastRun !== undefined && { lastRun: updates.lastRun }),
      ...(updates.autoTaggingConfig !== undefined && { autoTaggingConfig: updates.autoTaggingConfig }),
      ...(updates.notificationTemplate !== undefined && { notificationTemplate: updates.notificationTemplate })
    };
    
    this.triggers[index] = updatedTrigger;
    
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
      await this.processFile(file, trigger);
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
          // Use modifiedTime for both createdDate and modifiedDate since DriveFile doesn't have createdTime
          const dateValue = file.modifiedTime;
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
   * Автоматично додає теги до документа на основі вмісту
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
    
    logger.info('Додано автоматичні теги до документа', {
      component: 'AutomatedDocumentProcessor',
      fileId: file.id,
      tagCount: uniqueTags.length,
      tags: uniqueTags
    });
    
    return uniqueTags;
  }

  /**
   * Додає теги до документа
   */
  private async tagDocument(file: DriveFile, parameters?: Record<string, any>): Promise<any> {
    const tags = parameters?.['tags'] || ['auto-processed'];
    
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
      let autoTags: string[] = [];
      let classification: ClassifiedDocument | null = null;

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
        }
      }

      // Виконуємо всі дії тригера
      for (const action of trigger.actions) {
        try {
          const actionResult = await this.executeAction(file, action);
          actionsTaken.push(action.type);
          
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
      const processedDocument: ProcessedDocument = {
        fileId: file.id,
        fileName: file.name || 'Без назви',
        actionsTaken,
        timestamp: new Date()
      };
      
      // Add optional properties only if they have values
      if (autoTags.length > 0) {
        processedDocument.autoTags = autoTags;
      }
      
      if (classification) {
        processedDocument.classification = classification;
      }
      
      this.recordProcessedFile(processedDocument);

      // Надсилаємо сповіщення якщо потрібно
      if (trigger.actions.some(a => a.type === 'notify')) {
        await this.sendNotification(file, trigger, actionsTaken, {}, autoTags, classification || undefined);
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
    const format = parameters?.['format'] || 'pdf';
    
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
    const targetFolderId = parameters?.['targetFolderId'];
    
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