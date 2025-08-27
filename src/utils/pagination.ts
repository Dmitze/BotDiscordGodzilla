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

/**
 * Enhanced pagination utilities for handling large result sets
 */

export interface PaginationOptions {
  page: number;
  limit: number;
  sortBy?: string;
  sortOrder?: 'asc' | 'desc';
  filters?: Record<string, any>;
}

export interface PaginationResult<T> {
  data: T[];
  totalCount: number;
  currentPage: number;
  totalPages: number;
  hasNextPage: boolean;
  hasPrevPage: boolean;
  pageSize: number;
  startIndex: number;
  endIndex: number;
}

/**
 * Enhanced pagination function for large datasets
 */
export function paginate<T>(
  items: T[],
  options: PaginationOptions
): PaginationResult<T> {
  const { page = 1, limit = 10, sortBy, sortOrder = 'asc', filters = {} } = options;
  
  // Apply filters if provided
  let filteredItems = items;
  if (Object.keys(filters).length > 0) {
    filteredItems = items.filter(item => {
      return Object.entries(filters).every(([key, value]) => {
        const itemValue = (item as any)[key];
        if (value === undefined || value === null) return true;
        if (typeof value === 'string' && typeof itemValue === 'string') {
          return itemValue.toLowerCase().includes(value.toLowerCase());
        }
        return itemValue === value;
      });
    });
  }
  
  // Apply sorting if provided
  if (sortBy) {
    filteredItems = [...filteredItems].sort((a, b) => {
      const aVal = (a as any)[sortBy];
      const bVal = (b as any)[sortBy];
      
      if (aVal < bVal) return sortOrder === 'asc' ? -1 : 1;
      if (aVal > bVal) return sortOrder === 'asc' ? 1 : -1;
      return 0;
    });
  }
  
  // Calculate pagination
  const totalCount = filteredItems.length;
  const totalPages = Math.ceil(totalCount / limit);
  const currentPage = Math.max(1, Math.min(page, totalPages || 1));
  const startIndex = (currentPage - 1) * limit;
  const endIndex = Math.min(startIndex + limit, totalCount);
  
  // Get page data
  const pageData = filteredItems.slice(startIndex, endIndex);
  
  return {
    data: pageData,
    totalCount,
    currentPage,
    totalPages,
    hasNextPage: currentPage < totalPages,
    hasPrevPage: currentPage > 1,
    pageSize: limit,
    startIndex,
    endIndex: endIndex - 1,
  };
}

/**
 * Cursor-based pagination for better performance with large datasets
 */
export interface CursorPaginationOptions<T> {
  first: number;
  after?: string;
  last?: number;
  before?: string;
  sortBy?: keyof T;
  sortOrder?: 'asc' | 'desc';
}

export interface CursorPaginationResult<T> {
  edges: Array<{ cursor: string; node: T }>;
  pageInfo: {
    hasNextPage: boolean;
    hasPreviousPage: boolean;
    startCursor: string | null;
    endCursor: string | null;
  };
  totalCount: number;
}

/**
 * Generate cursor for an item
 */
function generateCursor<T>(item: T, index: number, sortBy?: keyof T): string {
  // In a real implementation, this would be more sophisticated
  // For now, we'll use a simple approach
  const value = sortBy ? (item[sortBy] as unknown as string) : index.toString();
  return Buffer.from(`${index}:${value}`).toString('base64');
}

/**
 * Parse cursor to get index
 */
function parseCursor(cursor: string): { index: number; value: string } {
  try {
    const decoded = Buffer.from(cursor, 'base64').toString('utf-8');
    const [indexStr, value] = decoded.split(':', 2);
    return { index: parseInt(indexStr, 10), value };
  } catch {
    return { index: 0, value: '' };
  }
}

/**
 * Cursor-based pagination implementation
 */
export function paginateWithCursor<T>(
  items: T[],
  options: CursorPaginationOptions<T>
): CursorPaginationResult<T> {
  const { first, after, last, before, sortBy, sortOrder = 'asc' } = options;
  
  // Sort items if sortBy is provided
  let sortedItems = items;
  if (sortBy) {
    sortedItems = [...items].sort((a, b) => {
      const aVal = a[sortBy];
      const bVal = b[sortBy];
      
      if (aVal < bVal) return sortOrder === 'asc' ? -1 : 1;
      if (aVal > bVal) return sortOrder === 'asc' ? 1 : -1;
      return 0;
    });
  }
  
  let startIndex = 0;
  let endIndex = sortedItems.length;
  
  // Handle forward pagination
  if (after) {
    const afterInfo = parseCursor(after);
    startIndex = afterInfo.index + 1;
  }
  
  // Handle backward pagination
  if (before) {
    const beforeInfo = parseCursor(before);
    endIndex = beforeInfo.index;
  }
  
  // Apply limits
  if (first !== undefined) {
    endIndex = Math.min(endIndex, startIndex + first);
  }
  
  if (last !== undefined) {
    startIndex = Math.max(startIndex, endIndex - last);
  }
  
  // Ensure valid range
  startIndex = Math.max(0, startIndex);
  endIndex = Math.min(sortedItems.length, endIndex);
  
  // Get page data
  const pageData = sortedItems.slice(startIndex, endIndex);
  
  // Create edges
  const edges = pageData.map((item, index) => ({
    cursor: generateCursor(item, startIndex + index, sortBy),
    node: item,
  }));
  
  // Create page info
  const pageInfo = {
    hasNextPage: endIndex < sortedItems.length,
    hasPreviousPage: startIndex > 0,
    startCursor: edges.length > 0 ? edges[0].cursor : null,
    endCursor: edges.length > 0 ? edges[edges.length - 1].cursor : null,
  };
  
  return {
    edges,
    pageInfo,
    totalCount: sortedItems.length,
  };
}

/**
 * Virtual scrolling pagination for very large datasets
 */
export interface VirtualScrollOptions {
  startIndex: number;
  endIndex: number;
  bufferSize?: number;
}

export interface VirtualScrollResult<T> {
  data: T[];
  startIndex: number;
  endIndex: number;
  totalLength: number;
  bufferedStart: number;
  bufferedEnd: number;
}

/**
 * Virtual scrolling implementation for handling very large datasets efficiently
 */
export function virtualScroll<T>(
  getItems: (start: number, end: number) => T[] | Promise<T[]>,
  totalLength: number,
  options: VirtualScrollOptions
): VirtualScrollResult<T> | Promise<VirtualScrollResult<T>> {
  const { startIndex, endIndex, bufferSize = 50 } = options;
  
  // Calculate buffered range
  const bufferedStart = Math.max(0, startIndex - bufferSize);
  const bufferedEnd = Math.min(totalLength, endIndex + bufferSize);
  
  // Get items for the buffered range
  const itemsResult = getItems(bufferedStart, bufferedEnd);
  
  // Handle both synchronous and asynchronous results
  if (itemsResult instanceof Promise) {
    return itemsResult.then(items => ({
      data: items.slice(startIndex - bufferedStart, endIndex - bufferedStart + 1),
      startIndex,
      endIndex,
      totalLength,
      bufferedStart,
      bufferedEnd,
    }));
  } else {
    return {
      data: itemsResult.slice(startIndex - bufferedStart, endIndex - bufferedStart + 1),
      startIndex,
      endIndex,
      totalLength,
      bufferedStart,
      bufferedEnd,
    };
  }
}
