import { BaseService } from '@/core/BaseService';
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

// New interfaces for enhanced functionality
export interface PersonalizedCollection {
  id: string;
  userId: string;
  name: string;
  description: string;
  fileIds: string[];
  createdAt: Date;
  updatedAt: Date;
  isPublic: boolean;
}

export interface SearchPatternAnalysis {
  userId: string;
  frequentTerms: { term: string; frequency: number }[];
  preferredCategories: { category: string; count: number }[];
  searchTimePatterns: { hour: number; count: number }[];
  averageResultsPerSearch: number;
  clickThroughRate: number;
}

export interface UserDocumentPreference {
  userId: string;
  fileId: string;
  preferenceScore: number;
  lastUpdated: Date;
  tags: string[];
}

export class DocumentAnalyticsService extends BaseService {
  private accessRecords: DocumentAccessRecord[] = [];
  private userInteractions: Map<string, UserDocumentInteraction[]> = new Map();
  private searchPatterns: SearchPattern[] = [];
  // New properties for enhanced functionality
  private personalizedCollections: PersonalizedCollection[] = [];
  private userPreferences: Map<string, UserDocumentPreference[]> = new Map();
  private readonly MAX_RECORDS = 10000;
  private readonly MAX_SEARCH_PATTERNS = 5000;
  private readonly MAX_COLLECTIONS_PER_USER = 50;
  private readonly MAX_PREFERENCES_PER_USER = 500;

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

      // Update user preferences based on access
      this.updateUserPreferences(userId, fileId, action);

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
      
      // Обмежуємо кількість інтеракцій
      if (userInteractions.length > 1000) {
        userInteractions = userInteractions.slice(-1000);
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
   * Оновлює переваги користувача щодо документів
   */
  private updateUserPreferences(userId: string, fileId: string, action: string): void {
    try {
      let preferences = this.userPreferences.get(userId) || [];
      
      // Шукаємо існуючу перевагу
      let preference = preferences.find(p => p.fileId === fileId);
      
      // Calculate preference score based on action type
      const actionScores: Record<string, number> = {
        'view': 1,
        'download': 3,
        'analyze': 2,
        'export': 2,
        'search': 0.5
      };
      
      const scoreIncrement = actionScores[action] || 1;
      
      if (preference) {
        // Оновлюємо існуючу перевагу
        preference.preferenceScore += scoreIncrement;
        preference.lastUpdated = new Date();
      } else {
        // Створюємо нову перевагу
        preference = {
          userId,
          fileId,
          preferenceScore: scoreIncrement,
          lastUpdated: new Date(),
          tags: []
        };
        preferences.push(preference);
      }
      
      // Обмежуємо кількість переваг
      if (preferences.length > this.MAX_PREFERENCES_PER_USER) {
        preferences = preferences
          .sort((a, b) => b.preferenceScore - a.preferenceScore)
          .slice(0, this.MAX_PREFERENCES_PER_USER);
      }
      
      // Оновлюємо переваги користувача
      this.userPreferences.set(userId, preferences);
    } catch (error) {
      logger.error('Помилка оновлення переваг користувача', {
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
   * Аналізує пошукові патерни користувача
   */
  analyzeUserSearchPatterns(userId: string): SearchPatternAnalysis {
    try {
      // Фільтруємо патерни для конкретного користувача
      const userPatterns = this.searchPatterns.filter(p => p.userId === userId);
      
      if (userPatterns.length === 0) {
        return {
          userId,
          frequentTerms: [],
          preferredCategories: [],
          searchTimePatterns: [],
          averageResultsPerSearch: 0,
          clickThroughRate: 0
        };
      }
      
      // Аналіз частотних термінів
      const termFrequency = new Map<string, number>();
      for (const pattern of userPatterns) {
        const words = pattern.query
          .toLowerCase()
          .replace(/[^\p{L}\p{N}\s]/gu, ' ')
          .split(/\s+/)
          .filter(word => word.length > 2);
        
        for (const word of words) {
          termFrequency.set(word, (termFrequency.get(word) || 0) + 1);
        }
      }
      
      const frequentTerms = Array.from(termFrequency.entries())
        .sort((a, b) => b[1] - a[1])
        .slice(0, 20)
        .map(([term, frequency]) => ({ term, frequency }));
      
      // Аналіз часових патернів пошуку
      const timePatterns = new Map<number, number>();
      for (const pattern of userPatterns) {
        const hour = pattern.timestamp.getHours();
        timePatterns.set(hour, (timePatterns.get(hour) || 0) + 1);
      }
      
      const searchTimePatterns = Array.from(timePatterns.entries())
        .map(([hour, count]) => ({ hour, count }));
      
      // Середня кількість результатів на пошук
      const totalResults = userPatterns.reduce((sum, p) => sum + p.resultsCount, 0);
      const averageResultsPerSearch = totalResults / userPatterns.length;
      
      // Click-through rate (відсоток пошуків, де було обрано результат)
      const searchesWithSelection = userPatterns.filter(p => p.selectedResult).length;
      const clickThroughRate = userPatterns.length > 0 
        ? searchesWithSelection / userPatterns.length 
        : 0;
      
      return {
        userId,
        frequentTerms,
        preferredCategories: [], // This would be implemented with document categorization
        searchTimePatterns,
        averageResultsPerSearch,
        clickThroughRate
      };
    } catch (error) {
      logger.error('Помилка аналізу пошукових патернів користувача', {
        component: 'DocumentAnalyticsService',
        userId,
        error: error instanceof Error ? error.message : String(error)
      });
      
      return {
        userId,
        frequentTerms: [],
        preferredCategories: [],
        searchTimePatterns: [],
        averageResultsPerSearch: 0,
        clickThroughRate: 0
      };
    }
  }

  /**
   * Створює персоналізовану колекцію
   */
  createPersonalizedCollection(
    userId: string,
    name: string,
    description: string,
    fileIds: string[] = [],
    isPublic: boolean = false
  ): PersonalizedCollection {
    try {
      // Перевіряємо ліміт колекцій для користувача
      const userCollections = this.personalizedCollections.filter(c => c.userId === userId);
      if (userCollections.length >= this.MAX_COLLECTIONS_PER_USER) {
        throw new Error(`Користувач досягнув максимального ліміту колекцій (${this.MAX_COLLECTIONS_PER_USER})`);
      }
      
      const collection: PersonalizedCollection = {
        id: this.generateId(),
        userId,
        name,
        description,
        fileIds,
        createdAt: new Date(),
        updatedAt: new Date(),
        isPublic
      };
      
      this.personalizedCollections.push(collection);
      
      logger.info('Створено персоналізовану колекцію', {
        component: 'DocumentAnalyticsService',
        userId,
        collectionId: collection.id,
        name
      });
      
      return collection;
    } catch (error) {
      logger.error('Помилка створення персоналізованої колекції', {
        component: 'DocumentAnalyticsService',
        userId,
        name,
        error: error instanceof Error ? error.message : String(error)
      });
      
      throw error;
    }
  }

  /**
   * Оновлює персоналізовану колекцію
   */
  updatePersonalizedCollection(
    collectionId: string,
    userId: string,
    updates: Partial<Omit<PersonalizedCollection, 'id' | 'userId' | 'createdAt'>>
  ): boolean {
    try {
      const index = this.personalizedCollections.findIndex(c => c.id === collectionId && c.userId === userId);
      
      if (index === -1) {
        return false;
      }
      
      this.personalizedCollections[index] = {
        ...this.personalizedCollections[index],
        ...updates,
        updatedAt: new Date()
      };
      
      logger.info('Оновлено персоналізовану колекцію', {
        component: 'DocumentAnalyticsService',
        userId,
        collectionId
      });
      
      return true;
    } catch (error) {
      logger.error('Помилка оновлення персоналізованої колекції', {
        component: 'DocumentAnalyticsService',
        userId,
        collectionId,
        error: error instanceof Error ? error.message : String(error)
      });
      
      return false;
    }
  }

  /**
   * Видаляє персоналізовану колекцію
   */
  deletePersonalizedCollection(collectionId: string, userId: string): boolean {
    try {
      const initialLength = this.personalizedCollections.length;
      this.personalizedCollections = this.personalizedCollections.filter(c => 
        !(c.id === collectionId && c.userId === userId)
      );
      
      const deleted = this.personalizedCollections.length < initialLength;
      
      if (deleted) {
        logger.info('Видалено персоналізовану колекцію', {
          component: 'DocumentAnalyticsService',
          userId,
          collectionId
        });
      }
      
      return deleted;
    } catch (error) {
      logger.error('Помилка видалення персоналізованої колекції', {
        component: 'DocumentAnalyticsService',
        userId,
        collectionId,
        error: error instanceof Error ? error.message : String(error)
      });
      
      return false;
    }
  }

  /**
   * Отримує персоналізовані колекції користувача
   */
  getUserCollections(userId: string): PersonalizedCollection[] {
    return this.personalizedCollections.filter(c => c.userId === userId);
  }

  /**
   * Отримує публічні колекції
   */
  getPublicCollections(): PersonalizedCollection[] {
    return this.personalizedCollections.filter(c => c.isPublic);
  }

  /**
   * Додає документ до колекції
   */
  addDocumentToCollection(collectionId: string, userId: string, fileId: string): boolean {
    try {
      const collection = this.personalizedCollections.find(c => c.id === collectionId && c.userId === userId);
      
      if (!collection) {
        return false;
      }
      
      // Перевіряємо чи документ вже в колекції
      if (!collection.fileIds.includes(fileId)) {
        collection.fileIds.push(fileId);
        collection.updatedAt = new Date();
        
        logger.info('Додано документ до колекції', {
          component: 'DocumentAnalyticsService',
          userId,
          collectionId,
          fileId
        });
      }
      
      return true;
    } catch (error) {
      logger.error('Помилка додавання документа до колекції', {
        component: 'DocumentAnalyticsService',
        userId,
        collectionId,
        fileId,
        error: error instanceof Error ? error.message : String(error)
      });
      
      return false;
    }
  }

  /**
   * Видаляє документ з колекції
   */
  removeDocumentFromCollection(collectionId: string, userId: string, fileId: string): boolean {
    try {
      const collection = this.personalizedCollections.find(c => c.id === collectionId && c.userId === userId);
      
      if (!collection) {
        return false;
      }
      
      // Видаляємо документ з колекції
      const initialLength = collection.fileIds.length;
      collection.fileIds = collection.fileIds.filter(id => id !== fileId);
      
      if (collection.fileIds.length < initialLength) {
        collection.updatedAt = new Date();
        
        logger.info('Видалено документ з колекції', {
          component: 'DocumentAnalyticsService',
          userId,
          collectionId,
          fileId
        });
        
        return true;
      }
      
      return false;
    } catch (error) {
      logger.error('Помилка видалення документа з колекції', {
        component: 'DocumentAnalyticsService',
        userId,
        collectionId,
        fileId,
        error: error instanceof Error ? error.message : String(error)
      });
      
      return false;
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
      
      // Отримуємо переваги користувача
      const userPreferences = this.userPreferences.get(userId) || [];
      
      // Якщо немає інтеракцій, повертаємо випадкові документи
      if (userInteractions.length === 0 && userPreferences.length === 0) {
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
      
      // Генеруємо рекомендації на основі переваг користувача
      const preferredFiles = userPreferences
        .sort((a, b) => b.preferenceScore - a.preferenceScore)
        .slice(0, 20);
      
      for (const preference of preferredFiles) {
        const file = files.find(f => f.id === preference.fileId);
        if (file) {
          // Перевіряємо чи файл вже в рекомендаціях
          const existing = recommendations.find(r => r.file.id === file.id);
          if (!existing) {
            recommendations.push({
              file,
              score: preference.preferenceScore * 0.7,
              reason: `Відповідає вашим інтересам`,
              category: this.getFileCategory(file)
            });
          }
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
    
    return Array.from(terms).slice(0, 20); // Максимум 20 термінів
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
   * Генерує ID для нових об'єктів
   */
  private generateId(): string {
    return Date.now().toString(36) + Math.random().toString(36).substr(2);
  }

  // === BaseService required methods ===
  
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
      searchPatterns: this.searchPatterns.length,
      personalizedCollections: this.personalizedCollections.length,
      userPreferences: Array.from(this.userPreferences.values()).reduce((sum, prefs) => sum + prefs.length, 0)
    };
  }

  protected onGetStats(): any {
    return {
      accessRecords: this.accessRecords.length,
      userInteractions: this.userInteractions.size,
      searchPatterns: this.searchPatterns.length,
      personalizedCollections: this.personalizedCollections.length,
      userPreferences: Array.from(this.userPreferences.values()).reduce((sum, prefs) => sum + prefs.length, 0)
    };
  }
}