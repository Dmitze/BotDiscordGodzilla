const { 
  EmbedBuilder, 
  ActionRowBuilder, 
  ButtonBuilder, 
  ButtonStyle,
  StringSelectMenuBuilder,
  StringSelectMenuOptionBuilder
} = require('discord.js');
const { GoogleSheetsRetryHandler } = require('../utils/retry');
const DataFormatters = require('../utils/formatters');
const config = require('../config/Config');

class EnhancedSearch {
  constructor() {
    this.retryHandler = new GoogleSheetsRetryHandler();
  }

  /**
   * Покращений пошук з діапазонами та сортуванням
   */
  async execute(interaction) {
    await interaction.deferReply();

    try {
      // Отримуємо параметри пошуку
      const filters = this.extractFilters(interaction);
      
      // Отримуємо дані з Google Sheets
      const sheetData = await this.getSheetData();
      if (!sheetData || sheetData.length === 0) {
        return interaction.editReply('❌ Немає даних для пошуку');
      }

      const headers = sheetData[0];
      const data = sheetData.slice(1);

      // Виконуємо пошук
      const results = this.performSearch(data, headers, filters);
      
      if (results.length === 0) {
        return interaction.editReply('🔍 Нічого не знайдено за вказаними критеріями');
      }

      // Сортуємо результати
      const sortedResults = this.sortResults(results, headers, filters.sortBy, filters.sortOrder);

      // Створюємо embed з результатами
      const embed = this.createResultsEmbed(sortedResults, headers, filters);
      
      // Створюємо кнопки для навігації
      const components = this.createNavigationComponents(sortedResults.length);

      await interaction.editReply({
        embeds: [embed],
        components: components
      });

    } catch (error) {
      console.error('Помилка покращеного пошуку:', error);
      await interaction.editReply('❌ Помилка при виконанні пошуку');
    }
  }

  /**
   * Отримання даних з Google Sheets з повторними спробами
   */
  async getSheetData() {
    return await this.retryHandler.execute(async () => {
      const response = await fetch(config.getGoogleSheetsUrl());
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }
      const data = await response.json();
      return data.values || [];
    });
  }

  /**
   * Витягування фільтрів з interaction
   */
  extractFilters(interaction) {
    const filters = {
      name: interaction.options.getString('номенклатура'),
      client: interaction.options.getString('контрагент'),
      series: interaction.options.getString('серія'),
      priceFrom: interaction.options.getNumber('ціна_від'),
      priceTo: interaction.options.getNumber('ціна_до'),
      quantityFrom: interaction.options.getNumber('кількість_від'),
      quantityTo: interaction.options.getNumber('кількість_до'),
      sortBy: interaction.options.getString('сортування') || 'назва',
      sortOrder: interaction.options.getString('порядок') || 'asc'
    };

    return filters;
  }

  /**
   * Виконання пошуку з фільтрами
   */
  performSearch(data, headers, filters) {
    return data.filter(row => {
      // Фільтр за назвою
      if (filters.name) {
        const nameIndex = this.getColumnIndex(headers, 'назва');
        if (nameIndex !== -1) {
          const name = String(row[nameIndex] || '').toLowerCase();
          if (!name.includes(filters.name.toLowerCase())) {
            return false;
          }
        }
      }

      // Фільтр за контрагентом
      if (filters.client) {
        const clientIndex = this.getColumnIndex(headers, 'контрагент');
        if (clientIndex !== -1) {
          const client = String(row[clientIndex] || '').toLowerCase();
          if (!client.includes(filters.client.toLowerCase())) {
            return false;
          }
        }
      }

      // Фільтр за серійним номером
      if (filters.series) {
        const seriesIndex = this.getColumnIndex(headers, 'серія');
        if (seriesIndex !== -1) {
          const series = String(row[seriesIndex] || '').toLowerCase();
          if (!series.includes(filters.series.toLowerCase())) {
            return false;
          }
        }
      }

      // Фільтр за діапазоном ціни
      if (filters.priceFrom !== null || filters.priceTo !== null) {
        const priceIndex = this.getColumnIndex(headers, 'ціна');
        if (priceIndex !== -1) {
          const price = parseFloat(row[priceIndex]) || 0;
          
          if (filters.priceFrom !== null && price < filters.priceFrom) {
            return false;
          }
          if (filters.priceTo !== null && price > filters.priceTo) {
            return false;
          }
        }
      }

      // Фільтр за діапазоном кількості
      if (filters.quantityFrom !== null || filters.quantityTo !== null) {
        const quantityIndex = this.getColumnIndex(headers, 'кількість');
        if (quantityIndex !== -1) {
          const quantity = parseFloat(row[quantityIndex]) || 0;
          
          if (filters.quantityFrom !== null && quantity < filters.quantityFrom) {
            return false;
          }
          if (filters.quantityTo !== null && quantity > filters.quantityTo) {
            return false;
          }
        }
      }

      return true;
    });
  }

  /**
   * Сортування результатів
   */
  sortResults(results, headers, sortBy, sortOrder) {
    const sortIndex = this.getColumnIndex(headers, sortBy);
    if (sortIndex === -1) return results;

    return results.sort((a, b) => {
      let aVal = a[sortIndex];
      let bVal = b[sortIndex];

      // Спроба числового сортування
      const aNum = parseFloat(aVal);
      const bNum = parseFloat(bVal);
      
      if (!isNaN(aNum) && !isNaN(bNum)) {
        aVal = aNum;
        bVal = bNum;
      } else {
        // Строкове сортування
        aVal = String(aVal || '').toLowerCase();
        bVal = String(bVal || '').toLowerCase();
      }

      if (sortOrder === 'desc') {
        return aVal > bVal ? -1 : aVal < bVal ? 1 : 0;
      } else {
        return aVal < bVal ? -1 : aVal > bVal ? 1 : 0;
      }
    });
  }

  /**
   * Створення embed з результатами
   */
  createResultsEmbed(results, headers, filters) {
    const embed = new EmbedBuilder()
      .setTitle('🔍 Результати покращеного пошуку')
      .setColor(0x00ff00)
      .setTimestamp();

    // Додаємо інформацію про фільтри
    const activeFilters = this.getActiveFilters(filters);
    if (activeFilters.length > 0) {
      embed.addFields({
        name: '📋 Активні фільтри',
        value: activeFilters.join('\n'),
        inline: false
      });
    }

    // Додаємо статистику
    embed.addFields({
      name: '📊 Статистика',
      value: `Знайдено: **${results.length}** записів\nСортування: **${filters.sortBy}** (${filters.sortOrder})`,
      inline: true
    });

    // Додаємо таблицю з результатами
    const tableData = results.slice(0, 10); // Перші 10 результатів
    const table = DataFormatters.formatTable(tableData, headers, 10);
    
    embed.setDescription(`\`\`\`md\n${table}\`\`\``);

    if (results.length > 10) {
      embed.setFooter({ 
        text: `Показано 10 з ${results.length} результатів. Використайте кнопки для навігації.` 
      });
    }

    return embed;
  }

  /**
   * Створення компонентів навігації
   */
  createNavigationComponents(totalResults) {
    const components = [];

    // Кнопки навігації
    const navigationRow = new ActionRowBuilder()
      .addComponents(
        new ButtonBuilder()
          .setCustomId('search_first_page')
          .setLabel('⏮️ Перша')
          .setStyle(ButtonStyle.Secondary),
        new ButtonBuilder()
          .setCustomId('search_prev_page')
          .setLabel('⬅️ Попередня')
          .setStyle(ButtonStyle.Primary),
        new ButtonBuilder()
          .setCustomId('search_next_page')
          .setLabel('➡️ Наступна')
          .setStyle(ButtonStyle.Primary),
        new ButtonBuilder()
          .setCustomId('search_last_page')
          .setLabel('⏭️ Остання')
          .setStyle(ButtonStyle.Secondary)
      );

    components.push(navigationRow);

    // Кнопки дій
    const actionsRow = new ActionRowBuilder()
      .addComponents(
        new ButtonBuilder()
          .setCustomId('search_export_excel')
          .setLabel('📊 Експорт Excel')
          .setStyle(ButtonStyle.Success),
        new ButtonBuilder()
          .setCustomId('search_export_csv')
          .setLabel('📄 Експорт CSV')
          .setStyle(ButtonStyle.Success),
        new ButtonBuilder()
          .setCustomId('search_refine')
          .setLabel('🔧 Уточнити пошук')
          .setStyle(ButtonStyle.Secondary)
      );

    components.push(actionsRow);

    return components;
  }

  /**
   * Отримання активних фільтрів для відображення
   */
  getActiveFilters(filters) {
    const active = [];

    if (filters.name) active.push(`📝 Назва: "${filters.name}"`);
    if (filters.client) active.push(`🏢 Контрагент: "${filters.client}"`);
    if (filters.series) active.push(`🔢 Серія: "${filters.series}"`);
    
    if (filters.priceFrom !== null || filters.priceTo !== null) {
      let priceFilter = '💰 Ціна: ';
      if (filters.priceFrom !== null && filters.priceTo !== null) {
        priceFilter += `${filters.priceFrom} - ${filters.priceTo}`;
      } else if (filters.priceFrom !== null) {
        priceFilter += `від ${filters.priceFrom}`;
      } else {
        priceFilter += `до ${filters.priceTo}`;
      }
      active.push(priceFilter);
    }

    if (filters.quantityFrom !== null || filters.quantityTo !== null) {
      let quantityFilter = '📦 Кількість: ';
      if (filters.quantityFrom !== null && filters.quantityTo !== null) {
        quantityFilter += `${filters.quantityFrom} - ${filters.quantityTo}`;
      } else if (filters.quantityFrom !== null) {
        quantityFilter += `від ${filters.quantityFrom}`;
      } else {
        quantityFilter += `до ${filters.quantityTo}`;
      }
      active.push(quantityFilter);
    }

    return active;
  }

  /**
   * Отримання індексу колонки
   */
  getColumnIndex(headers, field) {
    const headerMap = {
      назва: ['найменування номенклатури', 'назва', 'наименование номенклатуры'],
      серія: ['серійний номер', 'серйіний номер', 'серийный номер'],
      контрагент: ['контрагент', 'постачальник', 'поставщик'],
      кількість: ['кількість', 'залишок', 'остаток', 'количество'],
      ціна: ['ціна', 'цена', 'вартість', 'стоимость']
    };

    for (let i = 0; i < headers.length; i++) {
      const headerName = (headers[i] || '').toLowerCase().replace(/\s+/g, ' ').trim();
      if (headerMap[field]?.some(h => h.toLowerCase() === headerName)) {
        return i;
      }
    }
    return -1;
  }
}

module.exports = EnhancedSearch; 