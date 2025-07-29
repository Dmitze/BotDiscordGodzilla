/**
 * 📄 Команди для роботи з військовими документами ЗСУ
 * Спеціалізовані функції для різних типів документів
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
    .setName('документи')
    .setDescription('📄 Робота з військовими документами ЗСУ')
    .addSubcommand(subcommand =>
      subcommand
        .setName('особовий-склад')
        .setDescription('👥 Робота з особовим складом')
        .addStringOption(option =>
          option
            .setName('дія')
            .setDescription('Дія з особовим складом')
            .setRequired(true)
            .addChoices(
              { name: 'Пошук особового складу', value: 'search' },
              { name: 'Додати особу', value: 'add' },
              { name: 'Оновити дані', value: 'update' },
              { name: 'Звіт по особовому складу', value: 'report' },
              { name: 'Перевірка наявності', value: 'check' }
            )
        )
        .addStringOption(option =>
          option.setName('запит').setDescription('Пошуковий запит або дані').setRequired(false)
        )
    )
    .addSubcommand(subcommand =>
      subcommand
        .setName('техніка')
        .setDescription('🚗 Робота з технікою та озброєнням')
        .addStringOption(option =>
          option
            .setName('дія')
            .setDescription('Дія з технікою')
            .setRequired(true)
            .addChoices(
              { name: 'Пошук техніки', value: 'search' },
              { name: 'Додати техніку', value: 'add' },
              { name: 'Стан техніки', value: 'status' },
              { name: 'Звіт по техніці', value: 'report' },
              { name: 'Технічне обслуговування', value: 'maintenance' }
            )
        )
        .addStringOption(option =>
          option.setName('запит').setDescription('Пошуковий запит або дані').setRequired(false)
        )
    )
    .addSubcommand(subcommand =>
      subcommand
        .setName('матеріали')
        .setDescription('📦 Робота з матеріально-технічним забезпеченням')
        .addStringOption(option =>
          option
            .setName('дія')
            .setDescription('Дія з матеріалами')
            .setRequired(true)
            .addChoices(
              { name: 'Пошук матеріалів', value: 'search' },
              { name: 'Додати матеріали', value: 'add' },
              { name: 'Залишки', value: 'stock' },
              { name: 'Звіт по МТЗ', value: 'report' },
              { name: 'Поповнення', value: 'replenish' }
            )
        )
        .addStringOption(option =>
          option.setName('запит').setDescription('Пошуковий запит або дані').setRequired(false)
        )
    )
    .addSubcommand(subcommand =>
      subcommand
        .setName('операції')
        .setDescription('⚔️ Робота з оперативними документами')
        .addStringOption(option =>
          option
            .setName('дія')
            .setDescription('Дія з операціями')
            .setRequired(true)
            .addChoices(
              { name: 'Пошук операцій', value: 'search' },
              { name: 'Додати операцію', value: 'add' },
              { name: 'Статус операцій', value: 'status' },
              { name: 'Звіт по операціях', value: 'report' },
              { name: 'Планування', value: 'planning' }
            )
        )
        .addStringOption(option =>
          option.setName('запит').setDescription('Пошуковий запит або дані').setRequired(false)
        )
    )
    .addSubcommand(subcommand =>
      subcommand
        .setName('накази')
        .setDescription('📋 Робота з наказами та розпорядженнями')
        .addStringOption(option =>
          option
            .setName('дія')
            .setDescription('Дія з наказами')
            .setRequired(true)
            .addChoices(
              { name: 'Пошук наказів', value: 'search' },
              { name: 'Додати наказ', value: 'add' },
              { name: 'Статус виконання', value: 'status' },
              { name: 'Звіт по наказах', value: 'report' },
              { name: 'Архів наказів', value: 'archive' }
            )
        )
        .addStringOption(option =>
          option.setName('запит').setDescription('Пошуковий запит або дані').setRequired(false)
        )
    ),

  async execute(interaction) {
    try {
      await interaction.deferReply();

      const subcommand = interaction.options.getSubcommand();
      const action = interaction.options.getString('дія');
      const query = interaction.options.getString('запит') || '';

      logger.info(`Команда документи: ${subcommand} - ${action} від ${interaction.user.tag}`);

      switch (subcommand) {
        case 'особовий-склад':
          await this.handlePersonnel(interaction, action, query);
          break;
        case 'техніка':
          await this.handleEquipment(interaction, action, query);
          break;
        case 'матеріали':
          await this.handleMaterials(interaction, action, query);
          break;
        case 'операції':
          await this.handleOperations(interaction, action, query);
          break;
        case 'накази':
          await this.handleOrders(interaction, action, query);
          break;
        default:
          await interaction.editReply('❌ Невідомий тип документа');
      }
    } catch (error) {
      logger.error('Помилка команди документи:', error);
      await interaction.editReply({
        content: '❌ Помилка при роботі з документами. Спробуйте ще раз.',
        ephemeral: true,
      });
    }
  },

  /**
   * Обробка особового складу
   */
  async handlePersonnel(interaction, action, query) {
    const embed = new EmbedBuilder()
      .setColor(0x3498db)
      .setTitle('👥 Особовий склад ЗСУ')
      .setDescription(`**Дія:** ${this.getActionName(action)}`);

    switch (action) {
      case 'search':
        embed.addFields(
          { name: '🔍 Пошук', value: query || 'Всі особи', inline: true },
          { name: '📊 Результат', value: 'Знайдено 15 осіб', inline: true },
          {
            name: '📅 Останнє оновлення',
            value: new Date().toLocaleDateString('uk-UA'),
            inline: true,
          }
        );
        embed.addFields({
          name: '👤 Приклад результатів',
          value:
            '• Капітан Іванов І.І. - 1-ша рота\n• Старший лейтенант Петров П.П. - 2-га рота\n• Лейтенант Сидоров С.С. - 3-тя рота',
        });
        break;

      case 'add':
        embed.addFields(
          { name: '➕ Додавання особи', value: 'Форма для додавання нової особи', inline: true },
          { name: '📋 Поля', value: 'ПІБ, звання, підрозділ, посада', inline: true }
        );
        break;

      case 'report':
        embed.addFields(
          {
            name: '📊 Звіт по особовому складу',
            value: 'Загальна кількість: 150 осіб',
            inline: true,
          },
          { name: '👨‍✈️ Офіцери', value: '45 осіб', inline: true },
          { name: '👨‍💼 Сержанти', value: '75 осіб', inline: true },
          { name: '👨‍🎖️ Солдати', value: '30 осіб', inline: true }
        );
        break;

      default:
        embed.addFields({ name: 'ℹ️ Інформація', value: 'Функція в розробці' });
    }

    await interaction.editReply({ embeds: [embed] });
  },

  /**
   * Обробка техніки
   */
  async handleEquipment(interaction, action, query) {
    const embed = new EmbedBuilder()
      .setColor(0xe74c3c)
      .setTitle('🚗 Техніка та озброєння')
      .setDescription(`**Дія:** ${this.getActionName(action)}`);

    switch (action) {
      case 'search':
        embed.addFields(
          { name: '🔍 Пошук техніки', value: query || 'Вся техніка', inline: true },
          { name: '📊 Результат', value: 'Знайдено 25 одиниць', inline: true },
          { name: '🛠️ Стан', value: '80% бойової готовності', inline: true }
        );
        embed.addFields({
          name: '🚗 Приклад результатів',
          value:
            '• БМП-2 #001 - Бойовий готовий\n• Т-72 #015 - На ремонті\n• БТР-80 #023 - Бойовий готовий',
        });
        break;

      case 'status':
        embed.addFields(
          { name: '🟢 Бойовий готовий', value: '20 одиниць', inline: true },
          { name: '🟡 На ремонті', value: '3 одиниці', inline: true },
          { name: '🔴 Не готовий', value: '2 одиниці', inline: true }
        );
        break;

      case 'maintenance':
        embed.addFields(
          {
            name: '🔧 Технічне обслуговування',
            value: 'Заплановано на наступний тиждень',
            inline: true,
          },
          { name: '📅 Дата', value: '15.01.2024', inline: true },
          { name: '👨‍🔧 Відповідальний', value: 'Технічна служба', inline: true }
        );
        break;

      default:
        embed.addFields({ name: 'ℹ️ Інформація', value: 'Функція в розробці' });
    }

    await interaction.editReply({ embeds: [embed] });
  },

  /**
   * Обробка матеріалів
   */
  async handleMaterials(interaction, action, query) {
    const embed = new EmbedBuilder()
      .setColor(0xf39c12)
      .setTitle('📦 Матеріально-технічне забезпечення')
      .setDescription(`**Дія:** ${this.getActionName(action)}`);

    switch (action) {
      case 'stock':
        embed.addFields(
          { name: '📦 Залишки МТЗ', value: 'Поточний стан запасів', inline: true },
          { name: '⚠️ Критичний рівень', value: '3 позиції', inline: true },
          { name: '✅ Нормальний рівень', value: '47 позицій', inline: true }
        );
        embed.addFields({
          name: '🔴 Критичні позиції',
          value:
            '• Боєприпаси 5.45мм - 15% залишку\n• Паливо ДП - 20% залишку\n• Медикаменти - 25% залишку',
        });
        break;

      case 'replenish':
        embed.addFields(
          { name: '📋 Заявка на поповнення', value: 'Сформовано нову заявку', inline: true },
          { name: '📅 Термін', value: '7 днів', inline: true },
          { name: '💰 Вартість', value: '2,500,000 грн', inline: true }
        );
        break;

      default:
        embed.addFields({ name: 'ℹ️ Інформація', value: 'Функція в розробці' });
    }

    await interaction.editReply({ embeds: [embed] });
  },

  /**
   * Обробка операцій
   */
  async handleOperations(interaction, action, query) {
    const embed = new EmbedBuilder()
      .setColor(0x9b59b6)
      .setTitle('⚔️ Оперативні документи')
      .setDescription(`**Дія:** ${this.getActionName(action)}`);

    switch (action) {
      case 'status':
        embed.addFields(
          { name: '🟢 Активні операції', value: '3 операції', inline: true },
          { name: '🟡 Планування', value: '2 операції', inline: true },
          { name: '🔴 Завершені', value: '15 операцій', inline: true }
        );
        embed.addFields({
          name: '⚔️ Активні операції',
          value:
            '• Операція "Щит" - В ході\n• Операція "Воля" - Планування\n• Операція "Захисник" - В ході',
        });
        break;

      case 'planning':
        embed.addFields(
          { name: '📋 Планування операцій', value: 'Нові операції в розробці', inline: true },
          { name: '📅 Термін', value: 'До кінця місяця', inline: true },
          { name: '👥 Участь', value: '3 підрозділи', inline: true }
        );
        break;

      default:
        embed.addFields({ name: 'ℹ️ Інформація', value: 'Функція в розробці' });
    }

    await interaction.editReply({ embeds: [embed] });
  },

  /**
   * Обробка наказів
   */
  async handleOrders(interaction, action, query) {
    const embed = new EmbedBuilder()
      .setColor(0x2ecc71)
      .setTitle('📋 Накази та розпорядження')
      .setDescription(`**Дія:** ${this.getActionName(action)}`);

    switch (action) {
      case 'search':
        embed.addFields(
          { name: '🔍 Пошук наказів', value: query || 'Всі накази', inline: true },
          { name: '📊 Результат', value: 'Знайдено 8 наказів', inline: true },
          { name: '📅 Період', value: 'Останній місяць', inline: true }
        );
        embed.addFields({
          name: '📋 Приклад результатів',
          value:
            '• Наказ №001 - Про організацію служби\n• Наказ №002 - Про призначення\n• Наказ №003 - Про заходи безпеки',
        });
        break;

      case 'status':
        embed.addFields(
          { name: '✅ Виконано', value: '5 наказів', inline: true },
          { name: '🔄 В процесі', value: '2 накази', inline: true },
          { name: '⏳ Очікує', value: '1 наказ', inline: true }
        );
        break;

      default:
        embed.addFields({ name: 'ℹ️ Інформація', value: 'Функція в розробці' });
    }

    await interaction.editReply({ embeds: [embed] });
  },

  /**
   * Отримання назви дії
   */
  getActionName(action) {
    const actions = {
      search: 'Пошук',
      add: 'Додавання',
      update: 'Оновлення',
      report: 'Звіт',
      check: 'Перевірка',
      status: 'Статус',
      maintenance: 'Обслуговування',
      stock: 'Залишки',
      replenish: 'Поповнення',
      planning: 'Планування',
      archive: 'Архів',
    };
    return actions[action] || action;
  },
};
