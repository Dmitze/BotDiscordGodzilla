/**
 * Оптимізована команда пошуку
 * Використовує Redis кешування, Connection Pool та пагінацію
 */

const { SlashCommandBuilder, EmbedBuilder } = require('discord.js');
const logger = require('../utils/logger');
const Pagination = require('../utils/pagination');

module.exports = {
  data: new SlashCommandBuilder()
    .setName('пошук')
    .setDescription('🔍 Гнучкий пошук по документах ЗСУ')
    .addStringOption(option =>
      option
        .setName('запит')
        .setDescription('Що шукати? (наприклад: "особовий склад", "техніка", "зброя")')
        .setRequired(true)
    )
    .addStringOption(option =>
      option
        .setName('тип_документа')
        .setDescription('Тип документа для пошуку')
        .addChoices(
          { name: 'Всі документи', value: 'all' },
          { name: 'Накази', value: 'orders' },
          { name: 'Доповіді', value: 'reports' },
          { name: 'Звіти', value: 'statistics' },
          { name: 'Плани', value: 'plans' },
          { name: 'Інструкції', value: 'instructions' },
          { name: 'Протоколи', value: 'protocols' },
          { name: 'Картки', value: 'cards' },
          { name: 'Журнали', value: 'journals' }
        )
    )
    .addStringOption(option =>
      option
        .setName('дата_від')
        .setDescription('Дата від (формат: ДД.ММ.РРРР)')
    )
    .addStringOption(option =>
      option
        .setName('дата_до')
        .setDescription('Дата до (формат: ДД.ММ.РРРР)')
    )
    .addStringOption(option =>
      option
        .setName('підрозділ')
        .setDescription('Підрозділ для пошуку')
    )
    .addStringOption(option =>
      option
        .setName('пріоритет')
        .setDescription('Пріоритет документа')
        .addChoices(
          { name: 'Всі', value: 'all' },
          { name: 'Критичний', value: 'critical' },
          { name: 'Високий', value: 'high' },
          { name: 'Середній', value: 'medium' },
          { name: 'Низький', value: 'low' }
        )
    )
    .addIntegerOption(option =>
      option
        .setName('ліміт')
        .setDescription('Кількість результатів (макс. 50)')
        .setMinValue(1)
        .setMaxValue(50)
    ),

  async execute(interaction) {
    const startTime = Date.now();
    
    try {
      // Отримання параметрів пошуку
      const query = interaction.options.getString('запит');
      const documentType = interaction.options.getString('тип_документа') || 'all';
      const dateFrom = interaction.options.getString('дата_від');
      const dateTo = interaction.options.getString('дата_до');
      const unit = interaction.options.getString('підрозділ');
      const priority = interaction.options.getString('пріоритет') || 'all';
      const limit = interaction.options.getInteger('ліміт') || 20;

      // Відкладена відповідь
      await interaction.deferReply();

      // Отримання сервісів
      const cacheService = interaction.client.serviceContainer.get('cache');
      const googleService = interaction.client.serviceContainer.get('google');

      // Створення ключа кешу
      const cacheKey = this.generateCacheKey({
        query,
        documentType,
        dateFrom,
        dateTo,
        unit,
        priority,
        limit,
        userId: interaction.user.id,
      });

      // Спроба отримати з кешу
      let searchResults = await cacheService.get(cacheKey);
      
      if (!searchResults) {
        logger.info(`🔍 Пошук: ${query} (не знайдено в кеші)`);
        
        // Виконання пошуку
        searchResults = await this.performSearch(
          googleService,
          {
            query,
            documentType,
            dateFrom,
            dateTo,
            unit,
            priority,
            limit,
          }
        );

        // Збереження в кеш на 5 хвилин
        await cacheService.set(cacheKey, searchResults, 300, { log: true });
      } else {
        logger.info(`⚡ Пошук: ${query} (знайдено в кеші)`);
      }

      // Перевірка результатів
      if (!searchResults || searchResults.length === 0) {
        const noResultsEmbed = new EmbedBuilder()
          .setColor(0xFF6B6B)
          .setTitle('🔍 Результати пошуку')
          .setDescription(`Нічого не знайдено за запитом: **"${query}"**`)
          .addFields(
            { name: 'Запит', value: query, inline: true },
            { name: 'Тип документа', value: this.getDocumentTypeName(documentType), inline: true },
            { name: 'Час виконання', value: `${Date.now() - startTime}ms`, inline: true }
          )
          .setTimestamp();

        await interaction.editReply({ embeds: [noResultsEmbed] });
        return;
      }

      // Створення пагінації
      const pagination = new Pagination(searchResults, {
        itemsPerPage: 5,
        maxPages: 10,
        title: '🔍 Результати пошуку',
        description: `Знайдено **${searchResults.length}** результатів за запитом: **"${query}"**`,
        embedColor: 0x0099FF,
        footer: `Запит: ${query} • Тип: ${this.getDocumentTypeName(documentType)}`,
      });

      // Створення embed та кнопок
      const embed = pagination.createEmbed();
      const buttons = pagination.createNavigationButtons();

      // Додавання інформації про пошук
      embed.addFields(
        { name: '⏱️ Час виконання', value: `${Date.now() - startTime}ms`, inline: true },
        { name: '💾 Кеш', value: searchResults.cached ? '✅ Використано' : '❌ Не використано', inline: true },
        { name: '📊 Загальна кількість', value: searchResults.length.toString(), inline: true }
      );

      // Відправка результату
      const response = await interaction.editReply({
        embeds: [embed],
        components: buttons,
      });

      // Збереження пагінації для подальшого використання
      this.savePaginationState(interaction.user.id, pagination);

      logger.info(`✅ Пошук завершено: ${searchResults.length} результатів за ${Date.now() - startTime}ms`);

    } catch (error) {
      logger.error('❌ Помилка виконання команди пошуку:', error);
      
      const errorEmbed = new EmbedBuilder()
        .setColor(0xFF0000)
        .setTitle('❌ Помилка пошуку')
        .setDescription('Виникла помилка при виконанні пошуку. Спробуйте ще раз.')
        .addFields(
          { name: 'Помилка', value: error.message.substring(0, 1000), inline: false },
          { name: 'Час виконання', value: `${Date.now() - startTime}ms`, inline: true }
        )
        .setTimestamp();

      await interaction.editReply({ embeds: [errorEmbed] });
    }
  },

  /**
   * Виконання пошуку
   */
  async performSearch(googleService, searchParams) {
    const { query, documentType, dateFrom, dateTo, unit, priority, limit } = searchParams;

    try {
      // Отримання даних з Google Sheets
      const sheetData = await googleService.getSheetData(
        null, // використовуємо конфігурацію за замовчуванням
        'A:Z', // всі колонки
        {
          valueRenderOption: 'UNFORMATTED_VALUE',
          dateTimeRenderOption: 'SERIAL_NUMBER',
        }
      );

      if (!sheetData.values || sheetData.values.length < 2) {
        return [];
      }

      const headers = sheetData.values[0];
      const rows = sheetData.values.slice(1);

      // Фільтрація даних
      const filteredResults = this.filterData(rows, headers, searchParams);

      // Обмеження результатів
      const limitedResults = filteredResults.slice(0, limit);

      // Форматування результатів
      return this.formatResults(limitedResults, headers);

    } catch (error) {
      logger.error('❌ Помилка виконання пошуку:', error);
      throw error;
    }
  },

  /**
   * Фільтрація даних
   */
  filterData(rows, headers, searchParams) {
    const { query, documentType, dateFrom, dateTo, unit, priority } = searchParams;

    return rows.filter(row => {
      // Пошук по запиту
      const queryMatch = this.matchesQuery(row, headers, query);
      if (!queryMatch) return false;

      // Фільтр по типу документа
      if (documentType !== 'all') {
        const docTypeMatch = this.matchesDocumentType(row, headers, documentType);
        if (!docTypeMatch) return false;
      }

      // Фільтр по даті
      if (dateFrom || dateTo) {
        const dateMatch = this.matchesDateRange(row, headers, dateFrom, dateTo);
        if (!dateMatch) return false;
      }

      // Фільтр по підрозділу
      if (unit) {
        const unitMatch = this.matchesUnit(row, headers, unit);
        if (!unitMatch) return false;
      }

      // Фільтр по пріоритету
      if (priority !== 'all') {
        const priorityMatch = this.matchesPriority(row, headers, priority);
        if (!priorityMatch) return false;
      }

      return true;
    });
  },

  /**
   * Перевірка відповідності запиту
   */
  matchesQuery(row, headers, query) {
    if (!query) return true;

    const searchTerm = query.toLowerCase();
    
    return row.some((cell, index) => {
      if (!cell) return false;
      
      const header = headers[index];
      const cellValue = String(cell).toLowerCase();
      
      // Пошук в назві документа
      if (header && header.toLowerCase().includes('назва')) {
        return cellValue.includes(searchTerm);
      }
      
      // Пошук в описі
      if (header && header.toLowerCase().includes('опис')) {
        return cellValue.includes(searchTerm);
      }
      
      // Загальний пошук
      return cellValue.includes(searchTerm);
    });
  },

  /**
   * Перевірка типу документа
   */
  matchesDocumentType(row, headers, documentType) {
    const typeColumnIndex = headers.findIndex(h => 
      h.toLowerCase().includes('тип') || h.toLowerCase().includes('вид')
    );

    if (typeColumnIndex === -1) return true;

    const cellValue = String(row[typeColumnIndex]).toLowerCase();
    const typeMap = {
      'orders': ['наказ', 'order'],
      'reports': ['доповідь', 'report'],
      'statistics': ['звіт', 'statistics'],
      'plans': ['план', 'plan'],
      'instructions': ['інструкція', 'instruction'],
      'protocols': ['протокол', 'protocol'],
      'cards': ['картка', 'card'],
      'journals': ['журнал', 'journal'],
    };

    const allowedTypes = typeMap[documentType] || [];
    return allowedTypes.some(type => cellValue.includes(type));
  },

  /**
   * Перевірка діапазону дат
   */
  matchesDateRange(row, headers, dateFrom, dateTo) {
    const dateColumnIndex = headers.findIndex(h => 
      h.toLowerCase().includes('дата') || h.toLowerCase().includes('date')
    );

    if (dateColumnIndex === -1) return true;

    const cellValue = row[dateColumnIndex];
    if (!cellValue) return true;

    const cellDate = this.parseDate(cellValue);
    if (!cellDate) return true;

    if (dateFrom) {
      const fromDate = this.parseDate(dateFrom);
      if (fromDate && cellDate < fromDate) return false;
    }

    if (dateTo) {
      const toDate = this.parseDate(dateTo);
      if (toDate && cellDate > toDate) return false;
    }

    return true;
  },

  /**
   * Перевірка підрозділу
   */
  matchesUnit(row, headers, unit) {
    const unitColumnIndex = headers.findIndex(h => 
      h.toLowerCase().includes('підрозділ') || h.toLowerCase().includes('unit')
    );

    if (unitColumnIndex === -1) return true;

    const cellValue = String(row[unitColumnIndex]).toLowerCase();
    const searchUnit = unit.toLowerCase();

    return cellValue.includes(searchUnit);
  },

  /**
   * Перевірка пріоритету
   */
  matchesPriority(row, headers, priority) {
    const priorityColumnIndex = headers.findIndex(h => 
      h.toLowerCase().includes('пріоритет') || h.toLowerCase().includes('priority')
    );

    if (priorityColumnIndex === -1) return true;

    const cellValue = String(row[priorityColumnIndex]).toLowerCase();
    const priorityMap = {
      'critical': ['критичний', 'critical'],
      'high': ['високий', 'high'],
      'medium': ['середній', 'medium'],
      'low': ['низький', 'low'],
    };

    const allowedPriorities = priorityMap[priority] || [];
    return allowedPriorities.some(p => cellValue.includes(p));
  },

  /**
   * Парсинг дати
   */
  parseDate(dateString) {
    if (!dateString) return null;

    // Спроба парсингу різних форматів
    const formats = [
      /(\d{2})\.(\d{2})\.(\d{4})/, // ДД.ММ.РРРР
      /(\d{4})-(\d{2})-(\d{2})/,   // РРРР-ММ-ДД
      /(\d{1,2})\/(\d{1,2})\/(\d{4})/, // М/Д/РРРР
    ];

    for (const format of formats) {
      const match = dateString.toString().match(format);
      if (match) {
        const [, day, month, year] = match;
        return new Date(year, month - 1, day);
      }
    }

    // Спроба прямого парсингу
    const date = new Date(dateString);
    return isNaN(date.getTime()) ? null : date;
  },

  /**
   * Форматування результатів
   */
  formatResults(rows, headers) {
    return rows.map(row => {
      const result = {};
      
      headers.forEach((header, index) => {
        if (header && row[index] !== undefined) {
          result[header] = row[index];
        }
      });

      return result;
    });
  },

  /**
   * Генерація ключа кешу
   */
  generateCacheKey(params) {
    const { query, documentType, dateFrom, dateTo, unit, priority, limit, userId } = params;
    
    const keyParts = [
      'search',
      query.toLowerCase().replace(/[^a-z0-9]/g, ''),
      documentType,
      dateFrom || 'none',
      dateTo || 'none',
      unit ? unit.toLowerCase().replace(/[^a-z0-9]/g, '') : 'none',
      priority,
      limit,
      userId,
    ];

    return keyParts.join(':');
  },

  /**
   * Отримання назви типу документа
   */
  getDocumentTypeName(type) {
    const typeNames = {
      'all': 'Всі документи',
      'orders': 'Накази',
      'reports': 'Доповіді',
      'statistics': 'Звіти',
      'plans': 'Плани',
      'instructions': 'Інструкції',
      'protocols': 'Протоколи',
      'cards': 'Картки',
      'journals': 'Журнали',
    };

    return typeNames[type] || 'Невідомий тип';
  },

  /**
   * Збереження стану пагінації
   */
  savePaginationState(userId, pagination) {
    if (!this.paginationStates) {
      this.paginationStates = new Map();
    }

    this.paginationStates.set(userId, {
      pagination,
      timestamp: Date.now(),
    });

    // Очищення старих станів через 5 хвилин
    setTimeout(() => {
      this.paginationStates.delete(userId);
    }, 5 * 60 * 1000);
  },

  /**
   * Отримання стану пагінації
   */
  getPaginationState(userId) {
    if (!this.paginationStates) return null;

    const state = this.paginationStates.get(userId);
    if (!state) return null;

    // Перевірка чи не застарів стан
    if (Date.now() - state.timestamp > 5 * 60 * 1000) {
      this.paginationStates.delete(userId);
      return null;
    }

    return state.pagination;
  },
};
