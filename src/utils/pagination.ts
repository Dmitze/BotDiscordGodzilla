/**
 * Утиліта для пагінації великих даних
 * Оптимізована для роботи з Discord embeds та великими наборами даних
 * TypeScript версія
 */

import { EmbedBuilder, ActionRowBuilder, ButtonBuilder, ButtonStyle } from 'discord.js';
import logger from './logger';

interface PaginationOptions {
  itemsPerPage?: number;
  maxPages?: number;
  embedColor?: number;
  title?: string;
  description?: string;
  fields?: string[];
  footer?: string;
  timestamp?: Date;
}

interface PaginationStats {
  totalItems: number;
  totalPages: number;
  currentPage: number;
  itemsPerPage: number;
  hasNext: boolean;
  hasPrevious: boolean;
}

class Pagination {
  private data: any[];
  private currentPage: number;
  private itemsPerPage: number;
  private maxPages: number;
  private embedColor: number;
  private title: string;
  private description: string;
  private fields: string[];
  private footer: string;
  private timestamp: Date;
  private totalItems: number;
  private totalPages: number;

  constructor(data: any[], options: PaginationOptions = {}) {
    this.data = Array.isArray(data) ? data : [];
    this.currentPage = 0;
    this.itemsPerPage = options.itemsPerPage || 10;
    this.maxPages = options.maxPages || 50;
    this.embedColor = options.embedColor || 0x0099ff;
    this.title = options.title || 'Результати';
    this.description = options.description || '';
    this.fields = options.fields || [];
    this.footer = options.footer || '';
    this.timestamp = options.timestamp || new Date();
    
    this.totalItems = this.data.length;
    this.totalPages = Math.min(
      Math.ceil(this.totalItems / this.itemsPerPage),
      this.maxPages
    );
  }

  /**
   * Отримання поточної сторінки
   */
  getCurrentPage(): number {
    return this.currentPage;
  }

  /**
   * Отримання загальної кількості сторінок
   */
  getTotalPages(): number {
    return this.totalPages;
  }

  /**
   * Отримання загальної кількості елементів
   */
  getTotalItems(): number {
    return this.totalItems;
  }

  /**
   * Перевірка чи можна перейти на попередню сторінку
   */
  hasPreviousPage(): boolean {
    return this.currentPage > 0;
  }

  /**
   * Перевірка чи можна перейти на наступну сторінку
   */
  hasNextPage(): boolean {
    return this.currentPage < this.totalPages - 1;
  }

  /**
   * Перехід на попередню сторінку
   */
  previousPage(): boolean {
    if (this.hasPreviousPage()) {
      this.currentPage--;
      return true;
    }
    return false;
  }

  /**
   * Перехід на наступну сторінку
   */
  nextPage(): boolean {
    if (this.hasNextPage()) {
      this.currentPage++;
      return true;
    }
    return false;
  }

  /**
   * Перехід на конкретну сторінку
   */
  goToPage(page: number): boolean {
    if (page >= 0 && page < this.totalPages) {
      this.currentPage = page;
      return true;
    }
    return false;
  }

  /**
   * Отримання даних поточної сторінки
   */
  getCurrentPageData(): any[] {
    const startIndex = this.currentPage * this.itemsPerPage;
    const endIndex = Math.min(startIndex + this.itemsPerPage, this.totalItems);
    return this.data.slice(startIndex, endIndex);
  }

  /**
   * Створення Discord embed для поточної сторінки
   */
  createEmbed(): EmbedBuilder {
    const embed = new EmbedBuilder()
      .setTitle(this.title)
      .setColor(this.embedColor)
      .setTimestamp(this.timestamp);

    if (this.description) {
      embed.setDescription(this.description);
    }

    // Додавання полів
    const pageData = this.getCurrentPageData();
    if (pageData.length > 0) {
      pageData.forEach((item, index) => {
        const fieldName = this.formatFieldName(item, index);
        const fieldValue = this.formatFieldValue(item, index);
        
        if (fieldName && fieldValue) {
          embed.addFields({
            name: fieldName,
            value: fieldValue,
            inline: true
          });
        }
      });
    } else {
      embed.addFields({
        name: 'Немає даних',
        value: 'На цій сторінці немає даних для відображення',
        inline: false
      });
    }

    // Додавання footer
    const footerText = this.createFooterText();
    if (footerText) {
      embed.setFooter({ text: footerText });
    }

    return embed;
  }

  /**
   * Створення кнопок навігації
   */
  createNavigationButtons(): ActionRowBuilder<ButtonBuilder> {
    const row = new ActionRowBuilder<ButtonBuilder>();

    // Кнопка "Перша сторінка"
    row.addComponents(
      new ButtonBuilder()
        .setCustomId('pagination_first')
        .setLabel('⏮️')
        .setStyle(ButtonStyle.Secondary)
        .setDisabled(this.currentPage === 0)
    );

    // Кнопка "Попередня сторінка"
    row.addComponents(
      new ButtonBuilder()
        .setCustomId('pagination_prev')
        .setLabel('◀️')
        .setStyle(ButtonStyle.Primary)
        .setDisabled(!this.hasPreviousPage())
    );

    // Кнопка "Наступна сторінка"
    row.addComponents(
      new ButtonBuilder()
        .setCustomId('pagination_next')
        .setLabel('▶️')
        .setStyle(ButtonStyle.Primary)
        .setDisabled(!this.hasNextPage())
    );

    // Кнопка "Остання сторінка"
    row.addComponents(
      new ButtonBuilder()
        .setCustomId('pagination_last')
        .setLabel('⏭️')
        .setStyle(ButtonStyle.Secondary)
        .setDisabled(this.currentPage === this.totalPages - 1)
    );

    return row;
  }

  /**
   * Обробка взаємодії з кнопками
   */
  handleButtonInteraction(customId: string): boolean {
    switch (customId) {
      case 'pagination_first':
        return this.goToPage(0);
      case 'pagination_prev':
        return this.previousPage();
      case 'pagination_next':
        return this.nextPage();
      case 'pagination_last':
        return this.goToPage(this.totalPages - 1);
      default:
        return false;
    }
  }

  /**
   * Форматування назви поля
   */
  private formatFieldName(item: any, index: number): string {
    if (this.fields.length > 0) {
      const fieldIndex = index % this.fields.length;
      return this.fields[fieldIndex] || `Елемент ${index + 1}`;
    }
    return `Елемент ${index + 1}`;
  }

  /**
   * Форматування значення поля
   */
  private formatFieldValue(item: any, index: number): string {
    if (typeof item === 'string') {
      return this.truncateText(item, 100);
    } else if (typeof item === 'object' && item !== null) {
      return this.formatObjectValue(item);
    } else {
      return this.formatValue(item);
    }
  }

  /**
   * Форматування об'єкта
   */
  private formatObjectValue(obj: any): string {
    try {
      const entries = Object.entries(obj).slice(0, 3); // Беремо перші 3 поля
      const formatted = entries.map(([key, value]) => 
        `${this.capitalizeFirst(key)}: ${this.formatValue(value)}`
      ).join('\n');
      
      return this.truncateText(formatted, 100);
    } catch (error) {
      return 'Помилка форматування';
    }
  }

  /**
   * Форматування значення
   */
  private formatValue(value: any): string {
    if (value === null || value === undefined) {
      return '—';
    } else if (typeof value === 'string') {
      return value;
    } else if (typeof value === 'number') {
      return value.toString();
    } else if (typeof value === 'boolean') {
      return value ? 'Так' : 'Ні';
    } else if (Array.isArray(value)) {
      return value.slice(0, 3).join(', ') + (value.length > 3 ? '...' : '');
    } else {
      return String(value);
    }
  }

  /**
   * Створення тексту footer
   */
  private createFooterText(): string {
    const parts: string[] = [];
    
    if (this.footer) {
      parts.push(this.footer);
    }
    
    parts.push(`Сторінка ${this.currentPage + 1} з ${this.totalPages}`);
    parts.push(`Всього елементів: ${this.totalItems}`);
    
    return parts.join(' • ');
  }

  /**
   * Обрізання тексту
   */
  private truncateText(text: string, maxLength: number): string {
    if (text.length <= maxLength) return text;
    return text.substring(0, maxLength - 3) + '...';
  }

  /**
   * Капіталізація першої літери
   */
  private capitalizeFirst(str: string): string {
    if (!str) return str;
    return str.charAt(0).toUpperCase() + str.slice(1);
  }

  /**
   * Отримання статистики пагінації
   */
  getStats(): PaginationStats {
    return {
      totalItems: this.totalItems,
      totalPages: this.totalPages,
      currentPage: this.currentPage,
      itemsPerPage: this.itemsPerPage,
      hasNext: this.hasNextPage(),
      hasPrevious: this.hasPreviousPage(),
    };
  }

  /**
   * Створення пагінації з фільтром
   */
  static createWithFilter(data: any[], filterFn: (item: any) => boolean, options: PaginationOptions = {}): Pagination {
    const filteredData = data.filter(filterFn);
    return new Pagination(filteredData, options);
  }

  /**
   * Створення пагінації з сортуванням
   */
  static createWithSort(data: any[], sortFn: (a: any, b: any) => number, options: PaginationOptions = {}): Pagination {
    const sortedData = [...data].sort(sortFn);
    return new Pagination(sortedData, options);
  }

  /**
   * Створення пагінації з лімітом
   */
  static createWithLimit(data: any[], limit: number, options: PaginationOptions = {}): Pagination {
    const limitedData = data.slice(0, limit);
    return new Pagination(limitedData, options);
  }

  /**
   * Створення пагінації для пошуку
   */
  static createForSearch(data: any[], searchTerm: string, searchFields: string[] = [], options: PaginationOptions = {}): Pagination {
    if (!searchTerm) {
      return new Pagination(data, options);
    }

    const searchLower = searchTerm.toLowerCase();
    const filteredData = data.filter(item => {
      if (searchFields.length > 0) {
        return searchFields.some(field => {
          const value = item[field];
          return value && String(value).toLowerCase().includes(searchLower);
        });
      } else {
        return Object.values(item).some(value => 
          value && String(value).toLowerCase().includes(searchLower)
        );
      }
    });

    return new Pagination(filteredData, options);
  }
}

export default Pagination;
export { Pagination }; 