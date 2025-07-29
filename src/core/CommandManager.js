/**
 * Менеджер команд Discord бота
 * Централізоване управління всіма командами
 */

const fs = require('fs').promises;
const path = require('path');
const logger = require('../utils/logger');

class CommandManager {
  constructor(bot) {
    this.bot = bot;
    this.commands = new Map();
    this.commandHandlers = new Map();
    this.commandCategories = new Map();
  }

  /**
   * Ініціалізація менеджера команд
   */
  async initialize() {
    try {
      logger.info('📋 Ініціалізація менеджера команд...');

      // Завантаження команд
      await this.loadCommands();

      // Реєстрація обробників подій
      this.registerEventHandlers();

      logger.info(`✅ Завантажено ${this.commands.size} команд`);
    } catch (error) {
      logger.error('❌ Помилка ініціалізації менеджера команд:', error);
      throw error;
    }
  }

  /**
   * Завантаження команд з папки commands
   */
  async loadCommands() {
    const commandsPath = path.join(__dirname, '../commands');

    try {
      const files = await fs.readdir(commandsPath);
      const commandFiles = files.filter(file => file.endsWith('.js'));

      for (const file of commandFiles) {
        try {
          const filePath = path.join(commandsPath, file);
          const command = require(filePath);

          if (this.validateCommand(command)) {
            const commandName = command.data.name;
            this.commands.set(commandName, command);
            this.commandHandlers.set(commandName, command.execute.bind(command));

            // Категоризація команд
            const category = this.getCommandCategory(command);
            if (!this.commandCategories.has(category)) {
              this.commandCategories.set(category, []);
            }
            this.commandCategories.get(category).push(commandName);

            logger.debug(`📝 Завантажено команду: ${commandName}`);
          }
        } catch (error) {
          logger.error(`❌ Помилка завантаження команди ${file}:`, error);
        }
      }
    } catch (error) {
      logger.error('❌ Помилка читання папки команд:', error);
      throw error;
    }
  }

  /**
   * Валідація команди
   */
  validateCommand(command) {
    const required = ['data', 'execute'];
    const missing = required.filter(prop => !command[prop]);

    if (missing.length > 0) {
      logger.warn(`Команда містить відсутні властивості: ${missing.join(', ')}`);
      return false;
    }

    if (!command.data.name) {
      logger.warn('Команда не має назви');
      return false;
    }

    return true;
  }

  /**
   * Визначення категорії команди
   */
  getCommandCategory(command) {
    // Аналіз назви команди для визначення категорії
    const name = command.data.name.toLowerCase();

    if (name.includes('пошук') || name.includes('search')) return 'search';
    if (name.includes('документ') || name.includes('document')) return 'documents';
    if (name.includes('аналітик') || name.includes('analytics')) return 'analytics';
    if (name.includes('операці') || name.includes('operation')) return 'operations';
    if (name.includes('ai') || name.includes('штучний')) return 'ai';
    if (name.includes('файл') || name.includes('file')) return 'files';
    if (name.includes('статистик') || name.includes('stats')) return 'statistics';
    if (name.includes('допомог') || name.includes('help')) return 'help';

    return 'general';
  }

  /**
   * Реєстрація обробників подій
   */
  registerEventHandlers() {
    this.bot.client.on('interactionCreate', async interaction => {
      if (!interaction.isChatInputCommand()) return;

      await this.handleCommand(interaction);
    });
  }

  /**
   * Обробка команди
   */
  async handleCommand(interaction) {
    const commandName = interaction.commandName;
    const command = this.commands.get(commandName);

    if (!command) {
      logger.warn(`Незнайдена команда: ${commandName}`);
      await interaction.reply({
        content: '❌ Команда не знайдена',
        ephemeral: true,
      });
      return;
    }

    try {
      // Логування використання команди
      logger.info(`🎯 Команда ${commandName} виконана користувачем ${interaction.user.tag}`);

      // Перевірка прав доступу
      if (command.permissions && !this.checkPermissions(interaction, command.permissions)) {
        await interaction.reply({
          content: '❌ У вас немає прав для використання цієї команди',
          ephemeral: true,
        });
        return;
      }

      // Виконання команди
      await command.execute(interaction, this.bot);
    } catch (error) {
      logger.error(`❌ Помилка виконання команди ${commandName}:`, error);

      const errorMessage = '❌ Помилка виконання команди. Спробуйте ще раз.';

      if (interaction.deferred || interaction.replied) {
        await interaction.editReply({ content: errorMessage });
      } else {
        await interaction.reply({ content: errorMessage, ephemeral: true });
      }
    }
  }

  /**
   * Перевірка прав доступу
   */
  checkPermissions(interaction, requiredPermissions) {
    if (!interaction.guild) return false;

    const member = interaction.member;
    if (!member) return false;

    // Перевірка ролей
    if (requiredPermissions.roles) {
      const hasRole = member.roles.cache.some(role =>
        requiredPermissions.roles.includes(role.name)
      );
      if (!hasRole) return false;
    }

    // Перевірка прав
    if (requiredPermissions.permissions) {
      const hasPermission = member.permissions.has(requiredPermissions.permissions);
      if (!hasPermission) return false;
    }

    return true;
  }

  /**
   * Отримання команди за назвою
   */
  getCommand(name) {
    return this.commands.get(name);
  }

  /**
   * Отримання всіх команд
   */
  getAllCommands() {
    return Array.from(this.commands.values());
  }

  /**
   * Отримання команд за категорією
   */
  getCommandsByCategory(category) {
    const commandNames = this.commandCategories.get(category) || [];
    return commandNames.map(name => this.commands.get(name));
  }

  /**
   * Отримання всіх категорій
   */
  getCategories() {
    return Array.from(this.commandCategories.keys());
  }

  /**
   * Статистика команд
   */
  getStats() {
    return {
      total: this.commands.size,
      categories: this.commandCategories.size,
      byCategory: Object.fromEntries(
        Array.from(this.commandCategories.entries()).map(([category, commands]) => [
          category,
          commands.length,
        ])
      ),
    };
  }
}

module.exports = CommandManager;
