/**
 * 📊 Команди аналітики та звітності для ЗСУ
 * Спеціалізовані звіти та аналіз даних
 */

const {
  SlashCommandBuilder,
  EmbedBuilder,
  ActionRowBuilder,
  ButtonBuilder,
  ButtonStyle,
} = require('discord.js');
const logger = require('../utils/logger');

module.exports = {
  data: new SlashCommandBuilder()
    .setName('аналітика')
    .setDescription('📊 Аналітика та звітність ЗСУ')
    .addSubcommand(subcommand =>
      subcommand
        .setName('звіт')
        .setDescription('📋 Генерація звітів')
        .addStringOption(option =>
          option
            .setName('тип')
            .setDescription('Тип звіту')
            .setRequired(true)
            .addChoices(
              { name: 'Щоденний звіт', value: 'daily' },
              { name: 'Тижневий звіт', value: 'weekly' },
              { name: 'Місячний звіт', value: 'monthly' },
              { name: 'Звіт по особовому складу', value: 'personnel' },
              { name: 'Звіт по техніці', value: 'equipment' },
              { name: 'Звіт по операціях', value: 'operations' },
              { name: 'Звіт по МТЗ', value: 'materials' },
              { name: 'Звіт по наказах', value: 'orders' }
            )
        )
        .addStringOption(option =>
          option
            .setName('формат')
            .setDescription('Формат звіту')
            .setRequired(false)
            .addChoices(
              { name: 'Текстовий', value: 'text' },
              { name: 'Excel', value: 'excel' },
              { name: 'PDF', value: 'pdf' }
            )
        )
    )
    .addSubcommand(subcommand =>
      subcommand
        .setName('статистика')
        .setDescription('📈 Статистика та метрики')
        .addStringOption(option =>
          option
            .setName('категорія')
            .setDescription('Категорія статистики')
            .setRequired(true)
            .addChoices(
              { name: 'Загальна статистика', value: 'general' },
              { name: 'Бойова готовність', value: 'combat' },
              { name: 'Особовий склад', value: 'personnel' },
              { name: 'Техніка', value: 'equipment' },
              { name: 'Операції', value: 'operations' },
              { name: 'МТЗ', value: 'materials' },
              { name: 'Ефективність', value: 'efficiency' }
            )
        )
    )
    .addSubcommand(subcommand =>
      subcommand
        .setName('прогноз')
        .setDescription('🔮 Прогнозування та планування')
        .addStringOption(option =>
          option
            .setName('тип')
            .setDescription('Тип прогнозу')
            .setRequired(true)
            .addChoices(
              { name: 'Потреби в МТЗ', value: 'materials' },
              { name: 'Ремонт техніки', value: 'repairs' },
              { name: 'Особовий склад', value: 'personnel' },
              { name: 'Оперативні потреби', value: 'operations' },
              { name: 'Бюджет', value: 'budget' }
            )
        )
        .addIntegerOption(option =>
          option
            .setName('період')
            .setDescription('Період прогнозування (днів)')
            .setRequired(false)
            .setMinValue(1)
            .setMaxValue(365)
        )
    )
    .addSubcommand(subcommand =>
      subcommand
        .setName('порівняння')
        .setDescription('⚖️ Порівняльний аналіз')
        .addStringOption(option =>
          option
            .setName("об'єкт")
            .setDescription("Об'єкт порівняння")
            .setRequired(true)
            .addChoices(
              { name: 'Підрозділи', value: 'units' },
              { name: 'Періоди', value: 'periods' },
              { name: 'Операції', value: 'operations' },
              { name: 'Техніка', value: 'equipment' }
            )
        )
        .addStringOption(option =>
          option
            .setName('метрика')
            .setDescription('Метрика для порівняння')
            .setRequired(true)
            .addChoices(
              { name: 'Ефективність', value: 'efficiency' },
              { name: 'Витрати', value: 'costs' },
              { name: 'Результати', value: 'results' },
              { name: 'Час', value: 'time' }
            )
        )
    ),

  async execute(interaction) {
    try {
      await interaction.deferReply();

      const subcommand = interaction.options.getSubcommand();

      switch (subcommand) {
        case 'звіт':
          await this.handleReport(interaction);
          break;
        case 'статистика':
          await this.handleStatistics(interaction);
          break;
        case 'прогноз':
          await this.handleForecast(interaction);
          break;
        case 'порівняння':
          await this.handleComparison(interaction);
          break;
        default:
          await interaction.editReply('❌ Невідомий тип аналітики');
      }
    } catch (error) {
      logger.error('Помилка команди аналітика:', error);
      await interaction.editReply({
        content: '❌ Помилка при виконанні аналітики. Спробуйте ще раз.',
        ephemeral: true,
      });
    }
  },

  /**
   * Обробка звітів
   */
  async handleReport(interaction) {
    const reportType = interaction.options.getString('тип');
    const format = interaction.options.getString('формат') || 'text';

    const embed = new EmbedBuilder()
      .setColor(0x3498db)
      .setTitle('📋 Генерація звіту')
      .setDescription(`**Тип звіту:** ${this.getReportTypeName(reportType)}`)
      .addFields(
        { name: '📄 Формат', value: format.toUpperCase(), inline: true },
        { name: '📅 Дата', value: new Date().toLocaleDateString('uk-UA'), inline: true },
        { name: '⏱️ Статус', value: 'Генерується...', inline: true }
      );

    // Симуляція генерації звіту
    setTimeout(async () => {
      const reportEmbed = new EmbedBuilder()
        .setColor(0x2ecc71)
        .setTitle('✅ Звіт готовий')
        .setDescription(`**${this.getReportTypeName(reportType)}** успішно згенеровано`)
        .addFields(
          { name: '📄 Формат', value: format.toUpperCase(), inline: true },
          { name: '📊 Розмір', value: '2.5 MB', inline: true },
          { name: '📅 Дата створення', value: new Date().toLocaleDateString('uk-UA'), inline: true }
        );

      const downloadRow = new ActionRowBuilder().addComponents(
        new ButtonBuilder()
          .setCustomId(`download_${reportType}_${format}`)
          .setLabel('📥 Завантажити')
          .setStyle(ButtonStyle.Success),
        new ButtonBuilder()
          .setCustomId(`share_${reportType}`)
          .setLabel('📤 Поділитися')
          .setStyle(ButtonStyle.Primary)
      );

      await interaction.editReply({
        embeds: [reportEmbed],
        components: [downloadRow],
      });
    }, 2000);

    await interaction.editReply({ embeds: [embed] });
  },

  /**
   * Обробка статистики
   */
  async handleStatistics(interaction) {
    const category = interaction.options.getString('категорія');

    const embed = new EmbedBuilder()
      .setColor(0x9b59b6)
      .setTitle('📈 Статистика ЗСУ')
      .setDescription(`**Категорія:** ${this.getCategoryName(category)}`);

    switch (category) {
      case 'general':
        embed.addFields(
          { name: '👥 Особовий склад', value: '1,250 осіб', inline: true },
          { name: '🚗 Техніка', value: '85 одиниць', inline: true },
          { name: '⚔️ Активні операції', value: '3 операції', inline: true },
          { name: '📦 МТЗ', value: '95% забезпечення', inline: true },
          { name: '📋 Накази', value: '45 активних', inline: true },
          { name: '🎯 Бойова готовність', value: '87%', inline: true }
        );
        break;

      case 'combat':
        embed.addFields(
          { name: '🟢 Повна готовність', value: '75%', inline: true },
          { name: '🟡 Часткова готовність', value: '15%', inline: true },
          { name: '🔴 Не готовий', value: '10%', inline: true },
          { name: '⚔️ Бойові операції', value: '3 активні', inline: true },
          { name: '🛡️ Оборонні операції', value: '5 активних', inline: true },
          { name: '📊 Ефективність', value: '92%', inline: true }
        );
        break;

      case 'personnel':
        embed.addFields(
          { name: '👨‍✈️ Офіцери', value: '180 осіб (14.4%)', inline: true },
          { name: '👨‍💼 Сержанти', value: '450 осіб (36%)', inline: true },
          { name: '👨‍🎖️ Солдати', value: '620 осіб (49.6%)', inline: true },
          { name: '📈 Прибуття', value: '+15 цього місяця', inline: true },
          { name: '📉 Відбуття', value: '-8 цього місяця', inline: true },
          { name: '🎓 Навчання', value: '45 осіб', inline: true }
        );
        break;

      case 'equipment':
        embed.addFields(
          { name: '🟢 Бойовий готовий', value: '68 одиниць (80%)', inline: true },
          { name: '🟡 На ремонті', value: '12 одиниць (14%)', inline: true },
          { name: '🔴 Не готовий', value: '5 одиниць (6%)', inline: true },
          { name: '🔧 Плановий ремонт', value: '8 одиниць', inline: true },
          { name: '⛽ Паливо', value: '85% запасів', inline: true },
          { name: '💥 Боєприпаси', value: '92% запасів', inline: true }
        );
        break;

      default:
        embed.addFields({ name: 'ℹ️ Інформація', value: 'Статистика в розробці' });
    }

    await interaction.editReply({ embeds: [embed] });
  },

  /**
   * Обробка прогнозування
   */
  async handleForecast(interaction) {
    const forecastType = interaction.options.getString('тип');
    const period = interaction.options.getInteger('період') || 30;

    const embed = new EmbedBuilder()
      .setColor(0xf39c12)
      .setTitle('🔮 Прогнозування')
      .setDescription(`**Тип прогнозу:** ${this.getForecastTypeName(forecastType)}`)
      .addFields(
        { name: '📅 Період', value: `${period} днів`, inline: true },
        { name: '📊 Точність', value: '85%', inline: true },
        { name: '🔄 Статус', value: 'Розраховується...', inline: true }
      );

    // Симуляція розрахунків
    setTimeout(async () => {
      const forecastEmbed = new EmbedBuilder()
        .setColor(0x2ecc71)
        .setTitle('✅ Прогноз готовий')
        .setDescription(`**${this.getForecastTypeName(forecastType)}** на ${period} днів`);

      switch (forecastType) {
        case 'materials':
          forecastEmbed.addFields(
            { name: '📦 Критичні потреби', value: 'Боєприпаси 5.45мм - 15,000 шт', inline: true },
            { name: '⛽ Паливо', value: 'ДП - 5,000 л', inline: true },
            { name: '💊 Медикаменти', value: "Перев'язувальні матеріали", inline: true },
            { name: '💰 Вартість', value: '3,200,000 грн', inline: true },
            { name: '📅 Термін', value: '7-14 днів', inline: true }
          );
          break;

        case 'repairs':
          forecastEmbed.addFields(
            { name: '🔧 Плановий ремонт', value: '8 одиниць техніки', inline: true },
            { name: '⏱️ Тривалість', value: '10-15 днів', inline: true },
            { name: '💰 Вартість', value: '1,500,000 грн', inline: true },
            { name: '👨‍🔧 Персонал', value: '12 техніків', inline: true }
          );
          break;

        case 'personnel':
          forecastEmbed.addFields(
            { name: '👥 Прибуття', value: '+25 осіб', inline: true },
            { name: '📉 Відбуття', value: '-12 осіб', inline: true },
            { name: '🎓 Навчання', value: '35 осіб', inline: true },
            { name: '📊 Чистий приріст', value: '+13 осіб', inline: true }
          );
          break;

        default:
          forecastEmbed.addFields({ name: 'ℹ️ Інформація', value: 'Прогноз в розробці' });
      }

      await interaction.editReply({ embeds: [forecastEmbed] });
    }, 3000);

    await interaction.editReply({ embeds: [embed] });
  },

  /**
   * Обробка порівняння
   */
  async handleComparison(interaction) {
    const object = interaction.options.getString("об'єкт");
    const metric = interaction.options.getString('метрика');

    const embed = new EmbedBuilder()
      .setColor(0xe74c3c)
      .setTitle('⚖️ Порівняльний аналіз')
      .setDescription(
        `**Об'єкт:** ${this.getObjectName(object)} | **Метрика:** ${this.getMetricName(metric)}`
      );

    switch (object) {
      case 'units':
        embed.addFields(
          { name: '1️⃣ 1-ша бригада', value: 'Ефективність: 92%', inline: true },
          { name: '2️⃣ 2-га бригада', value: 'Ефективність: 88%', inline: true },
          { name: '3️⃣ 3-тя бригада', value: 'Ефективність: 85%', inline: true },
          { name: '🏆 Лідер', value: '1-ша бригада', inline: true },
          { name: '📊 Середня ефективність', value: '88.3%', inline: true }
        );
        break;

      case 'periods':
        embed.addFields(
          { name: '📅 Січень', value: 'Результати: 85%', inline: true },
          { name: '📅 Лютий', value: 'Результати: 92%', inline: true },
          { name: '📅 Березень', value: 'Результати: 88%', inline: true },
          { name: '📈 Тренд', value: 'Покращення +3%', inline: true },
          { name: '🎯 Ціль', value: '95% до кінця місяця', inline: true }
        );
        break;

      default:
        embed.addFields({ name: 'ℹ️ Інформація', value: 'Порівняння в розробці' });
    }

    await interaction.editReply({ embeds: [embed] });
  },

  /**
   * Отримання назви типу звіту
   */
  getReportTypeName(type) {
    const types = {
      daily: 'Щоденний звіт',
      weekly: 'Тижневий звіт',
      monthly: 'Місячний звіт',
      personnel: 'Звіт по особовому складу',
      equipment: 'Звіт по техніці',
      operations: 'Звіт по операціях',
      materials: 'Звіт по МТЗ',
      orders: 'Звіт по наказах',
    };
    return types[type] || type;
  },

  /**
   * Отримання назви категорії
   */
  getCategoryName(category) {
    const categories = {
      general: 'Загальна статистика',
      combat: 'Бойова готовність',
      personnel: 'Особовий склад',
      equipment: 'Техніка',
      operations: 'Операції',
      materials: 'МТЗ',
      efficiency: 'Ефективність',
    };
    return categories[category] || category;
  },

  /**
   * Отримання назви типу прогнозу
   */
  getForecastTypeName(type) {
    const types = {
      materials: 'Потреби в МТЗ',
      repairs: 'Ремонт техніки',
      personnel: 'Особовий склад',
      operations: 'Оперативні потреби',
      budget: 'Бюджет',
    };
    return types[type] || type;
  },

  /**
   * Отримання назви об'єкта
   */
  getObjectName(object) {
    const objects = {
      units: 'Підрозділи',
      periods: 'Періоди',
      operations: 'Операції',
      equipment: 'Техніка',
    };
    return objects[object] || object;
  },

  /**
   * Отримання назви метрики
   */
  getMetricName(metric) {
    const metrics = {
      efficiency: 'Ефективність',
      costs: 'Витрати',
      results: 'Результати',
      time: 'Час',
    };
    return metrics[metric] || metric;
  },
};
