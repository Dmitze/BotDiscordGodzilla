/**
 * Розумна класифікація та автоматичне тегування документів
 * Smart Document Classification & Auto-Tagging Service
 */

import { BaseServiceClass } from '@/core/BaseService';
import type { BotConfig } from '@/types';
import type { DriveFile } from '@/types/drive';
import logger from '@/utils/logger';

export interface DocumentCategory {
  id: string;
  name: string;
  description: string;
  keywords: string[];
  priority: number;
}

export interface ClassifiedDocument {
  file: DriveFile;
  categories: DocumentCategory[];
  confidence: number;
  tags: string[];
}

export class SmartDocumentClassifier extends BaseServiceClass {
  private categories: DocumentCategory[] = [
    {
      id: 'orders',
      name: 'Накази',
      description: 'Офіційні накази та розпорядження',
      keywords: ['наказ', 'розпорядження', 'постанова', 'директива', 'інструкція', 'наказую', 'розпоряджаю'],
      priority: 1
    },
    {
      id: 'reports',
      name: 'Звіти',
      description: 'Статистичні та аналітичні звіти',
      keywords: ['звіт', 'статистика', 'аналіз', 'дослідження', 'відомості', 'підсумки'],
      priority: 2
    },
    {
      id: 'personnel',
      name: 'Особовий склад',
      description: 'Документи щодо особового складу',
      keywords: ['особовий склад', 'військовослужбовець', 'призов', 'звільнення', 'перевірка', 'атестація'],
      priority: 3
    },
    {
      id: 'logistics',
      name: 'Матеріально-технічне забезпечення',
      description: 'Документи з постачання та логістики',
      keywords: ['постачання', 'логістика', 'матеріали', 'обладнання', 'запаси', 'доставка'],
      priority: 4
    },
    {
      id: 'finance',
      name: 'Фінансові документи',
      description: 'Бюджетні та фінансові документи',
      keywords: ['бюджет', 'фінансування', 'кошторис', 'видатки', 'фінанси', 'рахунок'],
      priority: 5
    },
    {
      id: 'operations',
      name: 'Операційні документи',
      description: 'Документи щодо операційної діяльності',
      keywords: ['операція', 'бойове завдання', 'маневри', 'тренування', 'взаємодія'],
      priority: 6
    },
    {
      id: 'training',
      name: 'Навчальні матеріали',
      description: 'Навчальні посібники та методичні матеріали',
      keywords: ['навчання', 'посібник', 'методика', 'інструкція', 'тренінг', 'курс'],
      priority: 7
    },
    {
      id: 'communications',
      name: 'Комунікації',
      description: 'Службові листи та комунікації',
      keywords: ['лист', 'повідомлення', 'спілкування', 'зв\'язок', 'кореспонденція'],
      priority: 8
    }
  ];

  constructor(config: BotConfig) {
    super('SmartDocumentClassifier', config);
  }

  /**
   * Класифікує документ на основі його вмісту та метаданих
   */
  async classifyDocument(file: DriveFile, content: string): Promise<ClassifiedDocument> {
    try {
      // Отримуємо ключові слова з назви файлу
      const nameKeywords = this.extractKeywords(file.name || '');
      
      // Отримуємо ключові слова з вмісту файлу
      const contentKeywords = this.extractKeywords(content);
      
      // Об'єднуємо всі ключові слова
      const allKeywords = [...nameKeywords, ...contentKeywords];
      
      // Оцінюємо відповідність до кожної категорії
      const categoryScores = this.categories.map(category => {
        const score = this.calculateCategoryScore(category, allKeywords);
        return { category, score };
      });
      
      // Сортуємо за спаданням оцінки
      categoryScores.sort((a, b) => b.score - a.score);
      
      // Відбираємо категорії з високою оцінкою
      const matchedCategories = categoryScores
        .filter(cs => cs.score > 0.1)
        .map(cs => cs.category);
      
      // Генеруємо теги на основі ключових слів
      const tags = this.generateTags(allKeywords, matchedCategories);
      
      // Обчислюємо загальну впевненість
      const confidence = matchedCategories.length > 0 ? 
        categoryScores[0].score : 0;
      
      return {
        file,
        categories: matchedCategories,
        confidence,
        tags
      };
    } catch (error) {
      logger.error('Помилка класифікації документу', {
        component: 'SmartDocumentClassifier',
        fileId: file.id,
        error: error instanceof Error ? error.message : String(error)
      });
      
      // Повертаємо документ без категорій у разі помилки
      return {
        file,
        categories: [],
        confidence: 0,
        tags: []
      };
    }
  }

  /**
   * Групує документи за категоріями
   */
  groupDocumentsByCategory(documents: ClassifiedDocument[]): Map<string, ClassifiedDocument[]> {
    const groups = new Map<string, ClassifiedDocument[]>();
    
    // Додаємо всі категорії як порожні групи
    for (const category of this.categories) {
      groups.set(category.id, []);
    }
    
    // Додаємо групу для некатегоризованих документів
    groups.set('uncategorized', []);
    
    // Розподіляємо документи по групах
    for (const doc of documents) {
      if (doc.categories.length > 0) {
        // Використовуємо категорію з найвищим пріоритетом
        const primaryCategory = doc.categories
          .sort((a, b) => a.priority - b.priority)[0];
        groups.get(primaryCategory.id)?.push(doc);
      } else {
        groups.get('uncategorized')?.push(doc);
      }
    }
    
    return groups;
  }

  /**
   * Витягує ключові слова з тексту
   */
  private extractKeywords(text: string): string[] {
    if (!text) return [];
    
    // Очищуємо текст та розбиваємо на слова
    const cleanText = text
      .toLowerCase()
      .replace(/[^\p{L}\p{N}\s]/gu, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    
    // Розбиваємо на слова та фільтруємо короткі слова
    const words = cleanText
      .split(' ')
      .filter(word => word.length > 2);
    
    // Видаляємо дублікати
    return [...new Set(words)];
  }

  /**
   * Обчислює оцінку відповідності категорії
   */
  private calculateCategoryScore(category: DocumentCategory, keywords: string[]): number {
    if (keywords.length === 0) return 0;
    
    // Підраховуємо кількість ключових слів, що відповідають категорії
    const matches = keywords.filter(keyword => 
      category.keywords.some(catKeyword => 
        keyword.includes(catKeyword) || catKeyword.includes(keyword)
      )
    );
    
    // Обчислюємо оцінку (від 0 до 1)
    return matches.length / keywords.length;
  }

  /**
   * Генерує теги на основі ключових слів та категорій
   */
  private generateTags(keywords: string[], categories: DocumentCategory[]): string[] {
    const tags = new Set<string>();
    
    // Додаємо теги з категорій
    for (const category of categories) {
      tags.add(category.name.toLowerCase());
    }
    
    // Додаємо найбільш частотні ключові слова як теги
    const keywordCounts = new Map<string, number>();
    for (const keyword of keywords) {
      keywordCounts.set(keyword, (keywordCounts.get(keyword) || 0) + 1);
    }
    
    // Сортуємо за частотою та обираємо топ-5
    const sortedKeywords = [...keywordCounts.entries()]
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5)
      .map(entry => entry[0]);
    
    for (const keyword of sortedKeywords) {
      tags.add(keyword);
    }
    
    return [...tags].slice(0, 10); // Максимум 10 тегів
  }

  /**
   * Отримує всі доступні категорії
   */
  getCategories(): DocumentCategory[] {
    return [...this.categories];
  }

  /**
   * Додає нову категорію
   */
  addCategory(category: DocumentCategory): void {
    this.categories.push(category);
    // Сортуємо за пріоритетом
    this.categories.sort((a, b) => a.priority - b.priority);
  }

  /**
   * Оновлює існуючу категорію
   */
  updateCategory(categoryId: string, updatedCategory: DocumentCategory): void {
    const index = this.categories.findIndex(c => c.id === categoryId);
    if (index !== -1) {
      this.categories[index] = updatedCategory;
      // Сортуємо за пріоритетом
      this.categories.sort((a, b) => a.priority - b.priority);
    }
  }

  /**
   * Видаляє категорію
   */
  removeCategory(categoryId: string): void {
    this.categories = this.categories.filter(c => c.id !== categoryId);
  }

  // === BaseServiceClass required methods ===
  
  protected async onInitialize(): Promise<void> {
    logger.info('SmartDocumentClassifier ініціалізовано', {
      component: 'SmartDocumentClassifier'
    });
  }

  protected async onShutdown(): Promise<void> {
    logger.info('SmartDocumentClassifier зупинено', {
      component: 'SmartDocumentClassifier'
    });
  }

  protected async onHealthCheck(): Promise<any> {
    return {
      healthy: true,
      service: this.name,
      categoriesCount: this.categories.length
    };
  }

  protected onGetStats(): any {
    return {
      categoriesCount: this.categories.length
    };
  }
}