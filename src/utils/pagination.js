/**
 * Утиліта для пагінації великих даних
 * Оптимізована для роботи з Discord embeds та великими наборами даних
 */

const logger = require('./logger');

class Pagination {
  constructor(data, options = {}) {
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
  getCurrentPage() {
    return this.currentPage;
  }

  /**
   * Отримання загальної кількості сторінок
   */
  getTotalPages() {
    return this.totalPages;
  }

  /**
   * Отримання загальної кількості елементів
   */
  getTotalItems() {
    return this.totalItems;
  }

  /**
   * Перевірка чи можна перейти на попередню сторінку
   */
  hasPreviousPage() {
    return this.currentPage > 0;
  }

  /**
   * Перевірка чи можна перейти на наступну сторінку
   */
  hasNextPage() {
    return this.currentPage < this.totalPages - 1;
  }

  /**
   * Перехід на попередню сторінку
   */
  previousPage() {
    if (this.hasPreviousPage()) {
      this.currentPage--;
      return true;
    }
    return false;
  }

  /**
   * Перехід на наступну сторінку
   */
  nextPage() {
    if (this.hasNextPage()) {
      this.currentPage++;
      return true;
    }
    return false;
  }

  /**
   * Перехід на конкретну сторінку
   */
  goToPage(page) {
    if (page >= 0 && page < this.totalPages) {
      this.currentPage = page;
      return true;
    }
    return false;
  }

  /**
   * Отримання даних поточної сторінки
   */
  getCurrentPageData() {
    const startIndex = this.currentPage * this.itemsPerPage;
    const endIndex = startIndex + this.itemsPerPage;
    return this.data.slice(startIndex, endIndex);
  }

  /**
   * Створення Discord Embed для поточної сторінки
   */
  createEmbed() {
    const { EmbedBuilder } = require('discord.js');
    const currentData = this.getCurrentPageData();
    
    const embed = new EmbedBuilder()
      .setColor(this.embedColor)
      .setTitle(this.title)
      .setDescription(this.description)
      .setTimestamp(this.timestamp);

    // Додавання полів
    if (this.fields.length > 0) {
      embed.addFields(this.fields);
    }

    // Додавання даних поточної сторінки
    if (currentData.length > 0) {
      currentData.forEach((item, index) => {
        const fieldName = this.formatFieldName(item, index);
        const fieldValue = this.formatFieldValue(item, index);
        
        if (fieldValue && fieldValue.length > 0) {
          embed.addFields({
            name: fieldName,
            value: fieldValue,
            inline: false,
          });
        }
      });
    } else {
      embed.addFields({
        name: '📭 Немає даних',
        value: 'На цій сторінці немає даних для відображення',
        inline: false,
      });
    }

    // Додавання футера з інформацією про сторінки
    const footerText = this.createFooterText();
    embed.setFooter({ text: footerText });

    return embed;
  }

  /**
   * Створення кнопок навігації
   */
  createNavigationButtons() {
    const { ActionRowBuilder, ButtonBuilder, ButtonStyle } = require('discord.js');
    
    const row = new ActionRowBuilder();

    // Кнопка "Перша сторінка"
    row.addComponents(
      new ButtonBuilder()
        .setCustomId('pagination_first')
        .setLabel('⏮️ Перша')
        .setStyle(ButtonStyle.Secondary)
        .setDisabled(this.currentPage === 0)
    );

    // Кнопка "Попередня сторінка"
    row.addComponents(
      new ButtonBuilder()
        .setCustomId('pagination_prev')
        .setLabel('◀️ Попередня')
        .setStyle(ButtonStyle.Primary)
        .setDisabled(!this.hasPreviousPage())
    );

    // Кнопка "Наступна сторінка"
    row.addComponents(
      new ButtonBuilder()
        .setCustomId('pagination_next')
        .setLabel('Наступна ▶️')
        .setStyle(ButtonStyle.Primary)
        .setDisabled(!this.hasNextPage())
    );

    // Кнопка "Остання сторінка"
    row.addComponents(
      new ButtonBuilder()
        .setCustomId('pagination_last')
        .setLabel('⏭️ Остання')
        .setStyle(ButtonStyle.Secondary)
        .setDisabled(this.currentPage === this.totalPages - 1)
    );

    return [row];
  }

  /**
   * Обробка натискання кнопки навігації
   */
  handleButtonInteraction(customId) {
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
  formatFieldName(item, index) {
    const globalIndex = this.currentPage * this.itemsPerPage + index + 1;
    
    if (typeof item === 'object' && item.name) {
      return `📋 ${item.name}`;
    }
    
    if (typeof item === 'object' && item.title) {
      return `📋 ${item.title}`;
    }
    
    return `📋 Запис ${globalIndex}`;
  }

  /**
   * Форматування значення поля
   */
  formatFieldValue(item, index) {
    if (typeof item === 'string') {
      return this.truncateText(item, 1024);
    }
    
    if (typeof item === 'object') {
      return this.formatObjectValue(item);
    }
    
    return String(item);
  }

  /**
   * Форматування об'єкта
   */
  formatObjectValue(obj) {
    const lines = [];
    
    for (const [key, value] of Object.entries(obj)) {
      if (value !== null && value !== undefined && key !== 'name' && key !== 'title') {
        const formattedValue = this.formatValue(value);
        if (formattedValue) {
          lines.push(`**${this.capitalizeFirst(key)}:** ${formattedValue}`);
        }
      }
    }
    
    const result = lines.join('\n');
    return this.truncateText(result, 1024);
  }

  /**
   * Форматування значення
   */
  formatValue(value) {
    if (typeof value === 'string') {
      return value;
    }
    
    if (typeof value === 'number') {
      return value.toLocaleString();
    }
    
    if (value instanceof Date) {
      return value.toLocaleDateString('uk-UA');
    }
    
    if (typeof value === 'boolean') {
      return value ? '✅ Так' : '❌ Ні';
    }
    
    if (Array.isArray(value)) {
      return value.slice(0, 5).join(', ') + (value.length > 5 ? '...' : '');
    }
    
    return String(value);
  }

  /**
   * Створення тексту футера
   */
  createFooterText() {
    const baseText = `Сторінка ${this.currentPage + 1} з ${this.totalPages} • Всього записів: ${this.totalItems}`;
    
    if (this.footer) {
      return `${baseText} • ${this.footer}`;
    }
    
    return baseText;
  }

  /**
   * Обрізання тексту до максимальної довжини
   */
  truncateText(text, maxLength) {
    if (text.length <= maxLength) {
      return text;
    }
    
    return text.substring(0, maxLength - 3) + '...';
  }

  /**
   * Капіталізація першої літери
   */
  capitalizeFirst(str) {
    return str.charAt(0).toUpperCase() + str.slice(1);
  }

  /**
   * Отримання статистики пагінації
   */
  getStats() {
    return {
      currentPage: this.currentPage,
      totalPages: this.totalPages,
      totalItems: this.totalItems,
      itemsPerPage: this.itemsPerPage,
      hasPreviousPage: this.hasPreviousPage(),
      hasNextPage: this.hasNextPage(),
      currentPageItems: this.getCurrentPageData().length,
    };
  }

  /**
   * Створення пагінації з фільтрацією
   */
  static createWithFilter(data, filterFn, options = {}) {
    const filteredData = data.filter(filterFn);
    return new Pagination(filteredData, options);
  }

  /**
   * Створення пагінації з сортуванням
   */
  static createWithSort(data, sortFn, options = {}) {
    const sortedData = [...data].sort(sortFn);
    return new Pagination(sortedData, options);
  }

  /**
   * Створення пагінації з обмеженням
   */
  static createWithLimit(data, limit, options = {}) {
    const limitedData = data.slice(0, limit);
    return new Pagination(limitedData, options);
  }

  /**
   * Створення пагінації для пошуку
   */
  static createForSearch(data, searchTerm, searchFields = [], options = {}) {
    const searchResults = data.filter(item => {
      if (typeof item === 'string') {
        return item.toLowerCase().includes(searchTerm.toLowerCase());
      }
      
      if (typeof item === 'object') {
        return searchFields.some(field => {
          const value = item[field];
          return value && String(value).toLowerCase().includes(searchTerm.toLowerCase());
        });
      }
      
      return false;
    });
    
    return new Pagination(searchResults, {
      ...options,
      title: options.title || `Результати пошуку: "${searchTerm}"`,
    });
  }
}

module.exports = Pagination; 