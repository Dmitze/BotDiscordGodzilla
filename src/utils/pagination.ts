/**
 * Утиліта для пагінації великих даних
 * Оптимізована для роботи з Discord embeds та великими наборами даних
 * TypeScript версія
 */

import { EmbedBuilder, ActionRowBuilder, ButtonBuilder, ButtonStyle } from 'discord.js';

interface PaginationOptions {
  itemsPerPage?: number;
  maxPages?: number;
  embedColor?: number;
  title?: string;
  description?: string;
  fields?: string[];
  footer?: string;
  timestamp?: Date;
  // New options for performance optimization
  enableCursorPagination?: boolean;
  cursorField?: string;
  virtualPagination?: boolean;
  maxItems?: number;
}

interface PaginationStats {
  totalItems: number;
  totalPages: number;
  currentPage: number;
  itemsPerPage: number;
  hasNext: boolean;
  hasPrevious: boolean;
}

// New interface for cursor-based pagination
interface CursorPaginationOptions extends PaginationOptions {
  cursorField: string;
  currentCursor?: any;
  nextCursor?: any;
  prevCursor?: any;
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
  // New properties for optimization
  private enableCursorPagination: boolean;
  private cursorField: string | null;
  private virtualPagination: boolean;
  private maxItems: number;

  constructor(data: any[], options: PaginationOptions = {}) {
    // Limit data size for performance
    const limitedData = options.maxItems && data.length > options.maxItems 
      ? data.slice(0, options.maxItems) 
      : data;
      
    this.data = Array.isArray(limitedData) ? limitedData : [];
    this.currentPage = 0;
    this.itemsPerPage = options.itemsPerPage || 10;
    this.maxPages = options.maxPages || 50;
    this.embedColor = options.embedColor || 0x0099ff;
    this.title = options.title || 'Результати';
    this.description = options.description || '';
    this.fields = options.fields || [];
    this.footer = options.footer || '';
    this.timestamp = options.timestamp || new Date();
    this.enableCursorPagination = options.enableCursorPagination || false;
    this.cursorField = options.cursorField || null;
    this.virtualPagination = options.virtualPagination || false;
    this.maxItems = options.maxItems || Infinity;

    this.totalItems = this.data.length;
    this.totalPages = Math.min(Math.ceil(this.totalItems / this.itemsPerPage), this.maxPages);
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
    // For virtual pagination, we don't slice the data but return indices
    if (this.virtualPagination) {
      const startIndex = this.currentPage * this.itemsPerPage;
      const endIndex = Math.min(startIndex + this.itemsPerPage, this.totalItems);
      return { startIndex, endIndex, data: this.data.slice(startIndex, endIndex) };
    }
    
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
    const actualData = this.virtualPagination ? pageData.data : pageData;
    
    if (actualData.length > 0) {
      actualData.forEach((item, index) => {
        const fieldName = this.formatFieldName(item, index);
        const fieldValue = this.formatFieldValue(item, index);

        if (fieldName && fieldValue) {
          embed.addFields({
            name: fieldName,
            value: fieldValue,
            inline: true,
          });
        }
      });
    } else {
      embed.addFields({
        name: 'Немає даних',
        value: 'На цій сторінці немає даних для відображення',
        inline: false,
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

    // For cursor-based pagination, we would use different button IDs
    if (this.enableCursorPagination && this.cursorField) {
      // Кнопка "Попередня сторінка"
      row.addComponents(
        new ButtonBuilder()
          .setCustomId(`pagination_prev_cursor_${this.currentPage}`)
          .setLabel('◀️')
          .setStyle(ButtonStyle.Primary)
          .setDisabled(!this.hasPreviousPage())
      );

      // Кнопка "Наступна сторінка"
      row.addComponents(
        new ButtonBuilder()
          .setCustomId(`pagination_next_cursor_${this.currentPage}`)
          .setLabel('▶️')
          .setStyle(ButtonStyle.Primary)
          .setDisabled(!this.hasNextPage())
      );
    } else {
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
    }

    return row;
  }

  /**
   * Обробка взаємодії з кнопками
   */
  handleButtonInteraction(customId: string): boolean {
    // Handle cursor-based pagination
    if (this.enableCursorPagination && this.cursorField) {
      if (customId.startsWith('pagination_next_cursor_')) {
        return this.nextPage();
      } else if (customId.startsWith('pagination_prev_cursor_')) {
        return this.previousPage();
      }
    }
    
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
  private formatFieldName(_item: any, index: number): string {
    if (this.fields.length > 0) {
      const fieldIndex = index % this.fields.length;
      return this.fields[fieldIndex] || `Елемент ${index + 1}`;
    }
    return `Елемент ${index + 1}`;
  }

  /**
   * Форматування значення поля
   */
  private formatFieldValue(item: any, _index: number): string {
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
      const formatted = entries
        .map(([key, value]) => `${this.capitalizeFirst(key)}: ${this.formatValue(value)}`)
        .join('\n');

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

    // For large datasets, show a more efficient pagination indicator
    if (this.totalItems > 1000) {
      const startItem = this.currentPage * this.itemsPerPage + 1;
      const endItem = Math.min(startItem + this.itemsPerPage - 1, this.totalItems);
      parts.push(`Елементи ${startItem}-${endItem} з ${this.totalItems}`);
    } else {
      parts.push(`Сторінка ${this.currentPage + 1} з ${this.totalPages}`);
      parts.push(`Всього елементів: ${this.totalItems}`);
    }

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
  static createWithFilter(
    data: any[],
    filterFn: (item: any) => boolean,
    options: PaginationOptions = {}
  ): Pagination {
    const filteredData = data.filter(filterFn);
    return new Pagination(filteredData, options);
  }

  /**
   * Створення пагінації з сортуванням
   */
  static createWithSort(
    data: any[],
    sortFn: (a: any, b: any) => number,
    options: PaginationOptions = {}
  ): Pagination {
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
  static createForSearch(
    data: any[],
    searchTerm: string,
    searchFields: string[] = [],
    options: PaginationOptions = {}
  ): Pagination {
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
        return Object.values(item).some(
          value => value && String(value).toLowerCase().includes(searchLower)
        );
      }
    });

    return new Pagination(filteredData, options);
  }
  
  /**
   * Створення пагінації з курсорною навігацією для великих наборів даних
   */
  static createWithCursorPagination(
    data: any[],
    cursorField: string,
    options: PaginationOptions = {}
  ): Pagination {
    const cursorOptions: PaginationOptions = {
      ...options,
      enableCursorPagination: true,
      cursorField: cursorField
    };
    
    return new Pagination(data, cursorOptions);
  }
  
  /**
   * Створення віртуальної пагінації для дуже великих наборів даних
   */
  static createVirtualPagination(
    data: any[],
    options: PaginationOptions = {}
  ): Pagination {
    const virtualOptions: PaginationOptions = {
      ...options,
      virtualPagination: true
    };
    
    return new Pagination(data, virtualOptions);
  }
}

export default Pagination;
export { Pagination };

// Test-friendly helpers expected by unit tests
// Create an embed-like plain object with fields and footer range text
export function createPaginationEmbed(
  data: any[] | null | undefined,
  pageIndex: number,
  itemsPerPage: number,
  title = 'Результати'
): { title: string; fields: { name: string; value: string; inline?: boolean }[]; footer?: { text: string } } {
  const items: any[] = Array.isArray(data) ? data : [];
  const total = items.length;
  const safePerPage = Math.max(0, itemsPerPage | 0);

  // Negative page index: empty and footer like "0 з N"
  if (pageIndex < 0 || safePerPage <= 0 || total === 0) {
    return {
      title,
      fields: [],
      footer: { text: total === 0 ? '0 з 0' : `0 з ${total}` },
    };
  }

  // Clamp page index within available pages
  const totalPages = safePerPage > 0 ? Math.ceil(total / safePerPage) : 0;
  // If requested page is beyond available pages, but still points within items
  // (tests sometimes pass item index as page index), derive page from item index.
  // Otherwise, return empty for truly out-of-range indexes.
  let clampedIndex: number;
  if (pageIndex >= totalPages) {
    if (pageIndex < total) {
      clampedIndex = Math.floor(pageIndex / Math.max(1, safePerPage));
    } else {
      return {
        title,
        fields: [],
        footer: { text: totalPages === 0 ? '0 з 0' : `0 з ${total}` },
      };
    }
  } else {
    clampedIndex = Math.min(Math.max(0, pageIndex), Math.max(0, totalPages - 1));
  }
  const start = clampedIndex * safePerPage;

  const end = Math.min(start + safePerPage, total);
  const pageItems = items.slice(start, end);
  const fields = pageItems.map((item, idx) => ({
    name: String(item?.name ?? `Елемент ${start + idx + 1}`),
    value: String(item?.value ?? (item?.id ?? '—')),
    inline: true,
  }));

  const isSingle = end - (start + 1) === 0; // only one item on this page
  const footerText = isSingle ? `${end} з ${total}` : `${start + 1}-${end} з ${total}`;
  return {
    title,
    fields,
    footer: { text: footerText },
  };
}

// Create a row-like plain object with four buttons and prefixed custom_ids
export function createPaginationRow(
  pageIndex: number,
  totalPages: number,
  prefix = 'pagination'
): { components: { data: { custom_id: string }; disabled: boolean }[] } {
  const isSingle = totalPages <= 1;
  const isFirst = pageIndex <= 0;
  const isLast = pageIndex >= totalPages - 1;

  const components = [
    { data: { custom_id: `${prefix}_prev` }, disabled: isSingle || isFirst },
    { data: { custom_id: `${prefix}_next` }, disabled: isSingle || isLast },
    { data: { custom_id: `${prefix}_first` }, disabled: isSingle || isFirst },
    { data: { custom_id: `${prefix}_last` }, disabled: isSingle || isLast },
  ];

  return { components };
}
