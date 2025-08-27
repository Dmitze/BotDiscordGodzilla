import { BaseService } from '@/core/BaseService';
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

export interface ProjectTheme {
  id: string;
  name: string;
  description: string;
  keywords: string[];
  color: string;
}

export interface DocumentRelationship {
  sourceId: string;
  targetId: string;
  relationshipType: 'reference' | 'attachment' | 'version' | 'related';
  confidence: number;
}

export interface ClassifiedDocument {
  file: DriveFile;
  categories: DocumentCategory[];
  confidence: number;
  tags: string[];
  // New properties for enhanced functionality
  projectThemes: ProjectTheme[];
  relationships: DocumentRelationship[];
}

export class SmartDocumentClassifier extends BaseService {
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

  // New project themes for grouping documents by projects/themes
  private projectThemes: ProjectTheme[] = [
    {
      id: 'project-a',
      name: 'Проект А',
      description: 'Основний проект розробки',
      keywords: ['проект а', 'розробка', 'основний'],
      color: '#FF6B6B'
    },
    {
      id: 'project-b',
      name: 'Проект Б',
      description: 'Дослідницький проект',
      keywords: ['проект б', 'дослідження', 'експеримент'],
      color: '#4ECDC4'
    },
    {
      id: 'maintenance',
      name: 'Технічне обслуговування',
      description: 'Документи з технічного обслуговування',
      keywords: ['обслуговування', 'техніка', 'ремонт'],
      color: '#45B7D1'
    },
    {
      id: 'training-program',
      name: 'Навчальна програма',
      description: 'Документи навчальної програми',
      keywords: ['навчання', 'програма', 'курс'],
      color: '#96CEB4'
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
      
      // Identify project themes for the document
      const projectThemes = this.identifyProjectThemes(file, allKeywords);
      
      // Identify relationships with other documents
      const relationships = this.identifyDocumentRelationships(file, allKeywords);
      
      return {
        file,
        categories: matchedCategories,
        confidence,
        tags,
        projectThemes,
        relationships
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
        tags: [],
        projectThemes: [],
        relationships: []
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
        const group = groups.get(primaryCategory.id);
        if (group) {
          group.push(doc);
        }
      } else {
        const group = groups.get('uncategorized');
        if (group) {
          group.push(doc);
        }
      }
    }
    
    return groups;
  }

  /**
   * Групує документи за проектами/темами
   */
  groupDocumentsByProjectTheme(documents: ClassifiedDocument[]): Map<string, ClassifiedDocument[]> {
    const groups = new Map<string, ClassifiedDocument[]>();
    
    // Додаємо всі теми як порожні групи
    for (const theme of this.projectThemes) {
      groups.set(theme.id, []);
    }
    
    // Додаємо групу для документів без теми
    groups.set('no-theme', []);
    
    // Розподіляємо документи по групах
    for (const doc of documents) {
      if (doc.projectThemes.length > 0) {
        // Використовуємо першу тему (найбільш вірогідну)
        const primaryTheme = doc.projectThemes[0];
        const group = groups.get(primaryTheme.id);
        if (group) {
          group.push(doc);
        }
      } else {
        const group = groups.get('no-theme');
        if (group) {
          group.push(doc);
        }
      }
    }
    
    return groups;
  }

  /**
   * Візуалізує зв'язки між документами
   */
  visualizeDocumentRelationships(documents: ClassifiedDocument[]): DocumentRelationship[] {
    const allRelationships: DocumentRelationship[] = [];
    
    // Збираємо всі зв'язки з документів
    for (const doc of documents) {
      allRelationships.push(...doc.relationships);
    }
    
    // Видаляємо дублікати зв'язків
    const uniqueRelationships = this.deduplicateRelationships(allRelationships);
    
    return uniqueRelationships;
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
    return Array.from(new Set(words));
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
    const sortedKeywords = Array.from(keywordCounts.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5)
      .map(entry => entry[0]);
    
    for (const keyword of sortedKeywords) {
      tags.add(keyword);
    }
    
    return Array.from(tags).slice(0, 10); // Максимум 10 тегів
  }

  /**
   * Ідентифікує проектні теми для документа
   */
  private identifyProjectThemes(file: DriveFile, keywords: string[]): ProjectTheme[] {
    const themeScores = this.projectThemes.map(theme => {
      const score = this.calculateThemeScore(theme, keywords, file);
      return { theme, score };
    });
    
    // Сортуємо за спаданням оцінки
    themeScores.sort((a, b) => b.score - a.score);
    
    // Відбираємо теми з високою оцінкою
    return themeScores
      .filter(ts => ts.score > 0.1)
      .map(ts => ts.theme);
  }

  /**
   * Обчислює оцінку відповідності теми
   */
  private calculateThemeScore(theme: ProjectTheme, keywords: string[], file: DriveFile): number {
    if (keywords.length === 0) return 0;
    
    // Підраховуємо кількість ключових слів, що відповідають темі
    const matches = keywords.filter(keyword => 
      theme.keywords.some(themeKeyword => 
        keyword.includes(themeKeyword) || themeKeyword.includes(keyword)
      )
    );
    
    // Також перевіряємо назву файлу
    const nameMatches = theme.keywords.filter(keyword => 
      (file.name?.toLowerCase() || '').includes(keyword)
    ).length;
    
    // Обчислюємо оцінку (від 0 до 1)
    const keywordScore = matches.length / keywords.length;
    const nameScore = nameMatches / theme.keywords.length;
    
    return (keywordScore * 0.7) + (nameScore * 0.3);
  }

  /**
   * Ідентифікує зв'язки документа з іншими документами
   */
  private identifyDocumentRelationships(file: DriveFile, keywords: string[]): DocumentRelationship[] {
    const relationships: DocumentRelationship[] = [];
    
    // Шукаємо згадки інших файлів у вмісті
    const fileReferences = this.extractFileReferences(keywords);
    
    // Створюємо зв'язки для знайдених згадок
    for (const reference of fileReferences) {
      relationships.push({
        sourceId: file.id,
        targetId: reference.fileId,
        relationshipType: reference.type,
        confidence: reference.confidence
      });
    }
    
    return relationships;
  }

  /**
   * Витягує згадки інших файлів з ключових слів
   */
  private extractFileReferences(keywords: string[]): Array<{fileId: string, type: DocumentRelationship['relationshipType'], confidence: number}> {
    // Це спрощена реалізація
    // У реальному застосунку тут би був аналіз посилань між документами
    return [];
  }

  /**
   * Видаляє дублікати зв'язків
   */
  private deduplicateRelationships(relationships: DocumentRelationship[]): DocumentRelationship[] {
    const uniqueMap = new Map<string, DocumentRelationship>();
    
    for (const rel of relationships) {
      // Створюємо унікальний ключ для зв'язку
      const key = `${rel.sourceId}-${rel.targetId}-${rel.relationshipType}`;
      
      // Якщо зв'язок вже існує, зберігаємо той з вищою впевненістю
      if (uniqueMap.has(key)) {
        const existing = uniqueMap.get(key)!;
        if (rel.confidence > existing.confidence) {
          uniqueMap.set(key, rel);
        }
      } else {
        uniqueMap.set(key, rel);
      }
    }
    
    return Array.from(uniqueMap.values());
  }

  /**
   * Отримує всі доступні категорії
   */
  getCategories(): DocumentCategory[] {
    return [...this.categories];
  }

  /**
   * Отримує всі доступні теми проектів
   */
  getProjectThemes(): ProjectTheme[] {
    return [...this.projectThemes];
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

  /**
   * Додає нову тему проекту
   */
  addProjectTheme(theme: ProjectTheme): void {
    this.projectThemes.push(theme);
  }

  /**
   * Оновлює існуючу тему проекту
   */
  updateProjectTheme(themeId: string, updatedTheme: ProjectTheme): void {
    const index = this.projectThemes.findIndex(t => t.id === themeId);
    if (index !== -1) {
      this.projectThemes[index] = updatedTheme;
    }
  }

  /**
   * Видаляє тему проекту
   */
  removeProjectTheme(themeId: string): void {
    this.projectThemes = this.projectThemes.filter(t => t.id !== themeId);
  }

  // === BaseService required methods ===
  
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
      categoriesCount: this.categories.length,
      themesCount: this.projectThemes.length
    };
  }

  protected onGetStats(): any {
    return {
      categoriesCount: this.categories.length,
      themesCount: this.projectThemes.length
    };
  }
}