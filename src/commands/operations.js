/**
 * ⚔️ Команди оперативного управління ЗСУ
 * Спеціалізовані функції для оперативної роботи
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
    .setName('операції')
    .setDescription('⚔️ Оперативне управління ЗСУ')
    .addSubcommand(subcommand =>
      subcommand
        .setName('ситуація')
        .setDescription('📊 Поточна оперативна ситуація')
        .addStringOption(option =>
          option
            .setName('сектор')
            .setDescription('Оперативний сектор')
            .setRequired(false)
            .addChoices(
              { name: 'Всі сектори', value: 'all' },
              { name: 'Сектор А', value: 'A' },
              { name: 'Сектор Б', value: 'B' },
              { name: 'Сектор В', value: 'C' },
              { name: 'Сектор Г', value: 'D' }
            )
        )
    )
    .addSubcommand(subcommand =>
      subcommand
        .setName('завдання')
        .setDescription('🎯 Управління завданнями')
        .addStringOption(option =>
          option
            .setName('дія')
            .setDescription('Дія з завданнями')
            .setRequired(true)
            .addChoices(
              { name: 'Поточні завдання', value: 'current' },
              { name: 'Нове завдання', value: 'new' },
              { name: 'Оновити статус', value: 'update' },
              { name: 'Завершити завдання', value: 'complete' },
              { name: 'Архів завдань', value: 'archive' }
            )
        )
        .addStringOption(option =>
          option.setName('запит').setDescription('Пошуковий запит або дані').setRequired(false)
        )
    )
    .addSubcommand(subcommand =>
      subcommand
        .setName('координація')
        .setDescription('🔄 Координація між підрозділами')
        .addStringOption(option =>
          option
            .setName('тип')
            .setDescription('Тип координації')
            .setRequired(true)
            .addChoices(
              { name: 'Вогнева підтримка', value: 'fire_support' },
              { name: 'Логістика', value: 'logistics' },
              { name: 'Розвідка', value: 'intelligence' },
              { name: 'Медична допомога', value: 'medical' },
              { name: "Зв'язок", value: 'communications' }
            )
        )
        .addStringOption(option =>
          option.setName('підрозділ').setDescription('Підрозділ для координації').setRequired(false)
        )
    )
    .addSubcommand(subcommand =>
      subcommand
        .setName('розвідка')
        .setDescription('🔍 Розвідувальні дані')
        .addStringOption(option =>
          option
            .setName('тип')
            .setDescription('Тип розвідки')
            .setRequired(true)
            .addChoices(
              { name: 'Повітряна розвідка', value: 'air' },
              { name: 'Наземна розвідка', value: 'ground' },
              { name: 'Технічна розвідка', value: 'technical' },
              { name: 'Агентурна розвідка', value: 'agent' },
              { name: 'Зведена розвідка', value: 'summary' }
            )
        )
        .addStringOption(option =>
          option.setName('район').setDescription('Район розвідки').setRequired(false)
        )
    )
    .addSubcommand(subcommand =>
      subcommand
        .setName("зв'язок")
        .setDescription("📡 Система зв'язку")
        .addStringOption(option =>
          option
            .setName('дія')
            .setDescription("Дія з системою зв'язку")
            .setRequired(true)
            .addChoices(
              { name: "Статус зв'язку", value: 'status' },
              { name: 'Налаштування каналів', value: 'channels' },
              { name: 'Передача повідомлень', value: 'messages' },
              { name: 'Технічне обслуговування', value: 'maintenance' },
              { name: 'Резервні канали', value: 'backup' }
            )
        )
    ),

  async execute(interaction) {
    try {
      await interaction.deferReply();

      const subcommand = interaction.options.getSubcommand();

      switch (subcommand) {
        case 'ситуація':
          await this.handleSituation(interaction);
          break;
        case 'завдання':
          await this.handleTasks(interaction);
          break;
        case 'координація':
          await this.handleCoordination(interaction);
          break;
        case 'розвідка':
          await this.handleIntelligence(interaction);
          break;
        case "зв'язок":
          await this.handleCommunications(interaction);
          break;
        default:
          await interaction.editReply('❌ Невідомий тип операції');
      }
    } catch (error) {
      logger.error('Помилка команди операції:', error);
      await interaction.editReply({
        content: '❌ Помилка при виконанні операції. Спробуйте ще раз.',
        ephemeral: true,
      });
    }
  },

  /**
   * Обробка оперативної ситуації
   */
  async handleSituation(interaction) {
    const sector = interaction.options.getString('сектор') || 'all';

    const embed = new EmbedBuilder()
      .setColor(0xe74c3c)
      .setTitle('📊 Оперативна ситуація')
      .setDescription(`**Сектор:** ${sector === 'all' ? 'Всі сектори' : `Сектор ${sector}`}`)
      .addFields(
        { name: '🟢 Контрольовані території', value: '75%', inline: true },
        { name: '🟡 Контактна лінія', value: '20%', inline: true },
        { name: '🔴 Ворожий контроль', value: '5%', inline: true },
        { name: '⚔️ Активні операції', value: '3 операції', inline: true },
        { name: '🛡️ Оборонні позиції', value: '12 позицій', inline: true },
        { name: "📡 Зв'язок", value: 'Стабільний', inline: true }
      );

    if (sector !== 'all') {
      embed.addFields(
        { name: `📍 Сектор ${sector}`, value: 'Детальна інформація по сектору' },
        { name: '👥 Підрозділи', value: '2 роти, 1 батарея' },
        { name: '🚗 Техніка', value: '8 одиниць бойової техніки' },
        { name: '🎯 Завдання', value: 'Оборона позицій, патрулювання' }
      );
    }

    const actionRow = new ActionRowBuilder().addComponents(
      new ButtonBuilder()
        .setCustomId(`situation_update_${sector}`)
        .setLabel('🔄 Оновити')
        .setStyle(ButtonStyle.Primary),
      new ButtonBuilder()
        .setCustomId(`situation_report_${sector}`)
        .setLabel('📋 Звіт')
        .setStyle(ButtonStyle.Secondary)
    );

    await interaction.editReply({
      embeds: [embed],
      components: [actionRow],
    });
  },

  /**
   * Обробка завдань
   */
  async handleTasks(interaction) {
    const action = interaction.options.getString('дія');
    const query = interaction.options.getString('запит') || '';

    const embed = new EmbedBuilder()
      .setColor(0x3498db)
      .setTitle('🎯 Управління завданнями')
      .setDescription(`**Дія:** ${this.getTaskActionName(action)}`);

    switch (action) {
      case 'current':
        embed.addFields(
          { name: '📋 Активні завдання', value: '5 завдань', inline: true },
          { name: '⏳ В процесі', value: '3 завдання', inline: true },
          { name: '✅ Завершені', value: '12 завдань', inline: true }
        );
        embed.addFields({
          name: '🎯 Поточні завдання',
          value:
            '• Завдання #001 - Оборона позицій\n• Завдання #002 - Патрулювання\n• Завдання #003 - Розвідка',
        });
        break;

      case 'new':
        embed.addFields(
          { name: '➕ Нове завдання', value: 'Форма створення завдання', inline: true },
          { name: '📋 Поля', value: 'Назва, опис, термін, відповідальний', inline: true },
          { name: '🎯 Пріоритет', value: 'Високий/Середній/Низький', inline: true }
        );
        break;

      case 'update':
        embed.addFields(
          { name: '🔄 Оновлення статусу', value: 'Виберіть завдання для оновлення', inline: true },
          { name: '📊 Статуси', value: 'Нове/В процесі/Завершене/Скасоване', inline: true }
        );
        break;

      default:
        embed.addFields({ name: 'ℹ️ Інформація', value: 'Функція в розробці' });
    }

    await interaction.editReply({ embeds: [embed] });
  },

  /**
   * Обробка координації
   */
  async handleCoordination(interaction) {
    const type = interaction.options.getString('тип');
    const unit = interaction.options.getString('підрозділ') || 'Всі підрозділи';

    const embed = new EmbedBuilder()
      .setColor(0x9b59b6)
      .setTitle('🔄 Координація між підрозділами')
      .setDescription(`**Тип:** ${this.getCoordinationTypeName(type)} | **Підрозділ:** ${unit}`);

    switch (type) {
      case 'fire_support':
        embed.addFields(
          {
            name: '🎯 Вогнева підтримка',
            value: 'Координація артилерійського вогню',
            inline: true,
          },
          { name: "📡 Канали зв'язку", value: 'Радіо, телефон', inline: true },
          { name: '🎖️ Відповідальний', value: 'Артилерійський командир', inline: true }
        );
        embed.addFields({
          name: '📋 Процедури',
          value: '• Запит вогневої підтримки\n• Координація цілей\n• Контроль вогню',
        });
        break;

      case 'logistics':
        embed.addFields(
          { name: '📦 Логістика', value: 'Постачання та транспортування', inline: true },
          { name: '🚚 Транспорт', value: '5 одиниць', inline: true },
          { name: '📅 Графік', value: 'Щоденний', inline: true }
        );
        break;

      case 'intelligence':
        embed.addFields(
          { name: '🔍 Розвідка', value: 'Обмін розвідувальними даними', inline: true },
          { name: '📊 Частота', value: 'Кожні 2 години', inline: true },
          { name: '🔐 Класифікація', value: 'Секретно', inline: true }
        );
        break;

      default:
        embed.addFields({ name: 'ℹ️ Інформація', value: 'Координація в розробці' });
    }

    await interaction.editReply({ embeds: [embed] });
  },

  /**
   * Обробка розвідки
   */
  async handleIntelligence(interaction) {
    const type = interaction.options.getString('тип');
    const area = interaction.options.getString('район') || 'Загальний район';

    const embed = new EmbedBuilder()
      .setColor(0xf39c12)
      .setTitle('🔍 Розвідувальні дані')
      .setDescription(`**Тип:** ${this.getIntelligenceTypeName(type)} | **Район:** ${area}`);

    switch (type) {
      case 'air':
        embed.addFields(
          { name: '✈️ Повітряна розвідка', value: 'БПЛА, літаки', inline: true },
          { name: '📊 Останні дані', value: '2 години тому', inline: true },
          { name: '🎯 Цілі', value: '3 виявлені цілі', inline: true }
        );
        embed.addFields({
          name: '📋 Результати',
          value: '• Концентрація техніки\n• Позиції противника\n• Рух колон',
        });
        break;

      case 'ground':
        embed.addFields(
          { name: '👥 Наземна розвідка', value: 'Розвідувальні групи', inline: true },
          { name: '📍 Позиції', value: '5 активних груп', inline: true },
          { name: "📡 Зв'язок", value: 'Стабільний', inline: true }
        );
        break;

      case 'summary':
        embed.addFields(
          { name: '📊 Зведена розвідка', value: 'Комплексна оцінка', inline: true },
          { name: '⚠️ Загрози', value: 'Середній рівень', inline: true },
          { name: '🎯 Рекомендації', value: 'Посилення оборони', inline: true }
        );
        break;

      default:
        embed.addFields({ name: 'ℹ️ Інформація', value: 'Розвідка в розробці' });
    }

    await interaction.editReply({ embeds: [embed] });
  },

  /**
   * Обробка зв'язку
   */
  async handleCommunications(interaction) {
    const action = interaction.options.getString('дія');

    const embed = new EmbedBuilder()
      .setColor(0x2ecc71)
      .setTitle("📡 Система зв'язку")
      .setDescription(`**Дія:** ${this.getCommunicationActionName(action)}`);

    switch (action) {
      case 'status':
        embed.addFields(
          { name: "🟢 Основний зв'язок", value: 'Функціонує', inline: true },
          { name: "🟡 Резервний зв'язок", value: 'Готовий', inline: true },
          { name: '📡 Канали', value: '5 активних', inline: true },
          { name: '📊 Якість', value: 'Відмінна', inline: true },
          { name: '🔧 Технічне обслуговування', value: 'Не потрібне', inline: true }
        );
        break;

      case 'channels':
        embed.addFields(
          { name: '📻 Радіо канали', value: '3 канали', inline: true },
          { name: '📞 Телефон', value: '2 лінії', inline: true },
          { name: '📡 Супутниковий', value: '1 канал', inline: true }
        );
        break;

      case 'messages':
        embed.addFields(
          { name: '📨 Передача повідомлень', value: 'Система готова', inline: true },
          { name: '📊 Статус', value: 'Всі канали відкриті', inline: true },
          { name: '🔐 Безпека', value: 'Шифрування активне', inline: true }
        );
        break;

      default:
        embed.addFields({ name: 'ℹ️ Інформація', value: "Зв'язок в розробці" });
    }

    await interaction.editReply({ embeds: [embed] });
  },

  /**
   * Отримання назви дії завдання
   */
  getTaskActionName(action) {
    const actions = {
      current: 'Поточні завдання',
      new: 'Нове завдання',
      update: 'Оновити статус',
      complete: 'Завершити завдання',
      archive: 'Архів завдань',
    };
    return actions[action] || action;
  },

  /**
   * Отримання назви типу координації
   */
  getCoordinationTypeName(type) {
    const types = {
      fire_support: 'Вогнева підтримка',
      logistics: 'Логістика',
      intelligence: 'Розвідка',
      medical: 'Медична допомога',
      communications: "Зв'язок",
    };
    return types[type] || type;
  },

  /**
   * Отримання назви типу розвідки
   */
  getIntelligenceTypeName(type) {
    const types = {
      air: 'Повітряна розвідка',
      ground: 'Наземна розвідка',
      technical: 'Технічна розвідка',
      agent: 'Агентурна розвідка',
      summary: 'Зведена розвідка',
    };
    return types[type] || type;
  },

  /**
   * Отримання назви дії зв'язку
   */
  getCommunicationActionName(action) {
    const actions = {
      status: "Статус зв'язку",
      channels: 'Налаштування каналів',
      messages: 'Передача повідомлень',
      maintenance: 'Технічне обслуговування',
      backup: 'Резервні канали',
    };
    return actions[action] || action;
  },
};
