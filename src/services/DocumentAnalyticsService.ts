import { BaseServiceClass } from '@/core/BaseService';
import type { BotConfig } from '@/types';
import type { DriveFile } from '@/types/drive';
import logger from '@/utils/logger';

export interface DocumentAccessRecord {
  fileId: string;
  userId: string;
  timestamp: Date;
  action: 'view' | 'download' | 'search' | 'analyze' | 'export';
  sessionId: string;
}

export interface UserDocumentInteraction {
  userId: string;
  fileId: string;
  interactionCount: number;
  lastInteraction: Date;
  preferredActions: string[];
}

export interface DocumentRecommendation {
  file: DriveFile;
  score: number;
  reason: string;
  category?: string;
}

export interface SearchPattern {
  query: string;
  userId: string;
  timestamp: Date;
  resultsCount: number;
  selectedResult?: string;
}

export class DocumentAnalyticsService extends BaseServiceClass {
  private accessRecords: DocumentAccessRecord[] = [];
  private userInteractions: Map<string, UserDocumentInteraction[]> = new Map();
  private searchPatterns: SearchPattern[] = [];
  private readonly MAX_RECORDS = 10000;
  private readonly MAX_SEARCH_PATTERNS = 5000;

  constructor(config: BotConfig) {
    super('DocumentAnalyticsService', config);
  }

  /**
   * Записує доступ до документа
   */
  recordDocumentAccess(
    fileId: string,
    userId: string,
    action: 'view' | 'download' | 'search' | 'analyze' | 'export',
    sessionId: string
  ): void {
    try {
      const record: DocumentAccessRecord = {
        fileId,
        userId,
        timestamp: new Date(),
        action,
        sessionId
      };

      // Додаємо запис
      this.accessRecords.push(record);

      // Обмежуємо кількість записів
      if (this.accessRecords.length > this.MAX_RECORDS) {
        this.accessRecords = this.accessRecords.slice(-this.MAX_RECORDS);
      }

      // Оновлюємо інтеракції користувача
      this.updateUserInteraction(userId, fileId, action);

      logger.debug('Записано доступ до документа', {
        component: 'DocumentAnalyticsService',
        fileId,
        userId,
        action
      });
    } catch (error) {
      logger.error('Помилка запису доступу до документа', {
        component: 'DocumentAnalyticsService',
        fileId,
        userId,
        error: error instanceof Error ? error.message : String(error)
      });
    }
  }

  /**
   * Оновлює інтеракції користувача з документом
   */
  private updateUserInteraction(userId: string, fileId: string, action: string): void {
    try {
      let userInteractions = this.userInteractions.get(userId) || [];
      
      // Шукаємо існуючу інтеракцію
      let interaction = userInteractions.find(i => i.fileId === fileId);
      
      if (interaction) {
        // Оновлюємо існуючу інтеракцію
        interaction.interactionCount++;
        interaction.lastInteraction = new Date();
        
        // Додаємо дію до списку улюблених дій якщо її ще немає
        if (!interaction.preferredActions.includes(action)) {
          interaction.preferredActions.push(action);
        }
      } else {
        // Створюємо нову інтеракцію
        interaction = {
          userId,
          fileId,
          interactionCount: 1,
          lastInteraction: new Date(),
          preferredActions: [action]
        };
        userInteractions.push(interaction);
      }
      
      // Оновлюємо інтеракції користувача
      this.userInteractions.set(userId, userInteractions);
    } catch (error) {
      logger.error('Помилка оновлення інтеракцій користувача', {
        component: 'DocumentAnalyticsService',
        userId,
        fileId,
        error: error instanceof Error ? error.message : String(error)
      });
    }
  }

  /**
   * Записує пошуковий запит
   */
  recordSearchPattern(
    query: string,
    userId: string,
    resultsCount: number,
    selectedResult?: string
  ): void {
    try {
      const pattern: SearchPattern = {
        query,
        userId,
        timestamp: new Date(),
        resultsCount,
        selectedResult
      };

      // Додаємо патерн
      this.searchPatterns.push(pattern);

      // Обмежуємо кількість патернів
      if (this.searchPatterns.length > this.MAX_SEARCH_PATTERNS) {
        this.searchPatterns = this.searchPatterns.slice(-this.MAX_SEARCH_PATTERNS);
      }

      logger.debug('Записано пошуковий патерн', {
        component: 'DocumentAnalyticsService',
        query,
        userId,
        resultsCount
      });
    } catch (error) {
      logger.error('Помилка запису пошукового патерна', {
        component: 'DocumentAnalyticsService',
        query,
        userId,
        error: error instanceof Error ? error.message : String(error)
      });
    }
  }

  /**
   * Генерує рекомендації для користувача
   */
  generateRecommendations(userId: string, files: DriveFile[], limit: number = 10): DocumentRecommendation[] {
    try {
      const recommendations: DocumentRecommendation[] = [];
      
      // Отримуємо інтеракції користувача
      const userInteractions = this.userInteractions.get(userId) || [];
      
      // Якщо немає інтеракцій, повертаємо випадкові документи
      if (userInteractions.length === 0) {
        return this.getRandomRecommendations(files, limit);
      }
      
      // Аналізуємо пошукові патерни користувача
      const userSearchPatterns = this.searchPatterns
        .filter(p => p.userId === userId)
        .slice(-50); // Останні 50 пошуків
      
      // Знаходимо популярні документи серед інтеракцій користувача
      const popularFiles = userInteractions
        .sort((a, b) => b.interactionCount - a.interactionCount)
        .slice(0, 20);
      
      // Генеруємо рекомендації на основі популярних документів
      for (const interaction of popularFiles) {
        const file = files.find(f => f.id === interaction.fileId);
        if (file) {
          recommendations.push({
            file,
            score: interaction.interactionCount * 0.5 + 
                   (Date.now() - interaction.lastInteraction.getTime()) / (1000 * 60 * 60 * 24) * 0.3,
            reason: `Ви часто переглядали цей документ (${interaction.interactionCount} разів)`,
            category: this.getFileCategory(file)
          });
        }
      }
      
      // Генеруємо рекомендації на основі пошукових патернів
      const searchTerms = this.extractSearchTerms(userSearchPatterns);
      for (const term of searchTerms) {
        const matchingFiles = files.filter(f => 
          (f.name && f.name.toLowerCase().includes(term.toLowerCase())) ||
          (f.mimeType && f.mimeType.toLowerCase().includes(term.toLowerCase()))
        );
        
        for (const file of matchingFiles) {
          // Перевіряємо чи файл вже в рекомендаціях
          const existing = recommendations.find(r => r.file.id === file.id);
          if (!existing) {
            recommendations.push({
              file,
              score: 0.7,
              reason: `Відповідає вашим пошуковим запитам (${term})`,
              category: this.getFileCategory(file)
            });
          }
        }
      }
      
      // Сортуємо за спаданням оцінки
      recommendations.sort((a, b) => b.score - a.score);
      
      // Повертаємо топ-N рекомендацій
      return recommendations.slice(0, limit);
    } catch (error) {
      logger.error('Помилка генерації рекомендацій', {
        component: 'DocumentAnalyticsService',
        userId,
        error: error instanceof Error ? error.message : String(error)
      });
      
      // Повертаємо випадкові рекомендації у разі помилки
      return this.getRandomRecommendations(files, limit);
    }
  }

  /**
   * Отримує випадкові рекомендації
   */
  private getRandomRecommendations(files: DriveFile[], limit: number): DocumentRecommendation[] {
    // Перемішуємо масив файлів
    const shuffled = [...files].sort(() => 0.5 - Math.random());
    
    // Повертаємо перші N файлів
    return shuffled.slice(0, limit).map(file => ({
      file,
      score: 0.5,
      reason: 'Можливо, це буде цікаво',
      category: this.getFileCategory(file)
    }));
  }

  /**
   * Витягує пошукові терміни з патернів
   */
  private extractSearchTerms(patterns: SearchPattern[]): string[] {
    const terms = new Set<string>();
    
    for (const pattern of patterns) {
      // Розбиваємо запит на слова
      const words = pattern.query
        .toLowerCase()
        .replace(/[^\p{L}\p{N}\s]/gu, ' ')
        .split(/\s+/)
        .filter(word => word.length > 2);
      
      for (const word of words) {
        terms.add(word);
      }
    }
    
    return [...terms].slice(0, 20); // Максимум 20 термінів
  }

  /**
   * Визначає категорію файлу
   */
  private getFileCategory(file: DriveFile): string {
    const mimeType = file.mimeType || '';
    
    if (mimeType.includes('document')) return 'Документи';
    if (mimeType.includes('spreadsheet')) return 'Таблиці';
    if (mimeType.includes('presentation')) return 'Презентації';
    if (mimeType.includes('pdf')) return 'PDF';
    if (mimeType.startsWith('image/')) return 'Зображення';
    if (mimeType.startsWith('video/')) return 'Відео';
    if (mimeType.startsWith('audio/')) return 'Аудіо';
    
    return 'Інше';
  }

  /**
   * Отримує статистику використання документа
   */
  getDocumentUsageStats(fileId: string): {
    totalViews: number;
    uniqueUsers: number;
    lastAccessed: Date | null;
    popularActions: string[];
  } {
    try {
      // Фільтруємо записи для конкретного файлу
      const fileRecords = this.accessRecords.filter(r => r.fileId === fileId);
      
      // Загальна кількість переглядів
      const totalViews = fileRecords.length;
      
      // Унікальні користувачі
      const uniqueUsers = new Set(fileRecords.map(r => r.userId)).size;
      
      // Останній доступ
      const lastAccessed = fileRecords.length > 0 
        ? new Date(Math.max(...fileRecords.map(r => r.timestamp.getTime())))
        : null;
      
      // Популярні дії
      const actionCounts = new Map<string, number>();
      for (const record of fileRecords) {
        actionCounts.set(record.action, (actionCounts.get(record.action) || 0) + 1);
      }
      
      const popularActions = [...actionCounts.entries()]
        .sort((a, b) => b[1] - a[1])
        .slice(0, 5)
        .map(entry => entry[0]);
      
      return {
        totalViews,
        uniqueUsers,
        lastAccessed,
        popularActions
      };
    } catch (error) {
      logger.error('Помилка отримання статистики документа', {
        component: 'DocumentAnalyticsService',
        fileId,
        error: error instanceof Error ? error.message : String(error)
      });
      
      return {
        totalViews: 0,
        uniqueUsers: 0,
        lastAccessed: null,
        popularActions: []
      };
    }
  }

  /**
   * Отримує найпопулярніші документи
   */
  getPopularDocuments(files: DriveFile[], limit: number = 10): DocumentRecommendation[] {
    try {
      // Підраховуємо кількість доступів для кожного файлу
      const accessCounts = new Map<string, number>();
      
      for (const record of this.accessRecords) {
        accessCounts.set(record.fileId, (accessCounts.get(record.fileId) || 0) + 1);
      }
      
      // Створюємо рекомендації
      const recommendations: DocumentRecommendation[] = [];
      
      for (const [fileId, count] of accessCounts) {
        const file = files.find(f => f.id === fileId);
        if (file) {
          recommendations.push({
            file,
            score: count,
            reason: `Популярний документ (${count} переглядів)`,
            category: this.getFileCategory(file)
          });
        }
      }
      
      // Сортуємо за спаданням оцінки
      recommendations.sort((a, b) => b.score - a.score);
      
      // Повертаємо топ-N
      return recommendations.slice(0, limit);
    } catch (error) {
      logger.error('Помилка отримання популярних документів', {
        component: 'DocumentAnalyticsService',
        error: error instanceof Error ? error.message : String(error)
      });
      
      return [];
    }
  }

  /**
   * Отримує історію доступу користувача
   */
  getUserAccessHistory(userId: string, limit: number = 50): DocumentAccessRecord[] {
    return this.accessRecords
      .filter(r => r.userId === userId)
      .sort((a, b) => b.timestamp.getTime() - a.timestamp.getTime())
      .slice(0, limit);
  }

  /**
   * Отримує найпопулярніші пошукові запити
   */
  getPopularSearchQueries(limit: number = 20): { query: string; count: number }[] {
    try {
      // Підраховуємо кількість запитів
      const queryCounts = new Map<string, number>();
      
      for (const pattern of this.searchPatterns) {
        queryCounts.set(pattern.query, (queryCounts.get(pattern.query) || 0) + 1);
      }
      
      // Сортуємо за спаданням
      return [...queryCounts.entries()]
        .sort((a, b) => b[1] - a[1])
        .slice(0, limit)
        .map(entry => ({ query: entry[0], count: entry[1] }));
    } catch (error) {
      logger.error('Помилка отримання популярних пошукових запитів', {
        component: 'DocumentAnalyticsService',
        error: error instanceof Error ? error.message : String(error)
      });
      
      return [];
    }
  }

  // === BaseServiceClass required methods ===
  
  protected async onInitialize(): Promise<void> {
    logger.info('DocumentAnalyticsService ініціалізовано', {
      component: 'DocumentAnalyticsService'
    });
  }

  protected async onShutdown(): Promise<void> {
    logger.info('DocumentAnalyticsService зупинено', {
      component: 'DocumentAnalyticsService'
    });
  }

  protected async onHealthCheck(): Promise<any> {
    return {
      healthy: true,
      service: this.name,
      accessRecords: this.accessRecords.length,
      userInteractions: this.userInteractions.size,
      searchPatterns: this.searchPatterns.length
    };
  }

  protected onGetStats(): any {
    return {
      accessRecords: this.accessRecords.length,
      userInteractions: this.userInteractions.size,
      searchPatterns: this.searchPatterns.length
    };
  }
}