/**
 * Команда для роботи з Google Drive та різними форматами файлів
 * Включає пошук, читання та аналіз файлів
 */

const { SlashCommandBuilder, AttachmentBuilder } = require('discord.js');
const {
  checkPermission,
  ROLES,
  sanitizeInput,
  validateCommandOptions,
} = require('../utils/security');
const { fileProcessor } = require('../utils/fileProcessor');
const { aiEnhanced } = require('../utils/aiEnhanced');
const { logSecurityEvent } = require('../utils/security');
const logger = require('../utils/logger');
const fs = require('fs').promises;
const path = require('path');

module.exports = {
  data: new SlashCommandBuilder()
    .setName('файли')
    .setDescription('📁 Робота з файлами в Google Drive')
    .addSubcommand(subcommand =>
      subcommand
        .setName('пошук')
        .setDescription('Пошук файлів у Google Drive')
        .addStringOption(option =>
          option
            .setName('запит')
            .setDescription('Назва файлу для пошуку')
            .setRequired(true)
            .setMaxLength(200)
        )
        .addStringOption(option =>
          option
            .setName('папка')
            .setDescription('ID папки для пошуку (опціонально)')
            .setRequired(false)
            .setMaxLength(50)
        )
    )
    .addSubcommand(subcommand =>
      subcommand
        .setName('читати')
        .setDescription('Читати вміст файлу')
        .addStringOption(option =>
          option
            .setName('id')
            .setDescription('ID файлу в Google Drive')
            .setRequired(true)
            .setMaxLength(50)
        )
    )
    .addSubcommand(subcommand =>
      subcommand
        .setName('аналіз')
        .setDescription('AI-аналіз вмісту файлу')
        .addStringOption(option =>
          option
            .setName('id')
            .setDescription('ID файлу в Google Drive')
            .setRequired(true)
            .setMaxLength(50)
        )
        .addStringOption(option =>
          option
            .setName('тип')
            .setDescription('Тип аналізу')
            .setRequired(false)
            .addChoices(
              { name: 'Короткий зміст', value: 'summary' },
              { name: 'Детальний аналіз', value: 'detailed' },
              { name: 'Ключові моменти', value: 'key_points' }
            )
        )
    )
    .addSubcommand(subcommand =>
      subcommand
        .setName('звіт')
        .setDescription('Створити звіт на основі файлу')
        .addStringOption(option =>
          option
            .setName('id')
            .setDescription('ID файлу в Google Drive')
            .setRequired(true)
            .setMaxLength(50)
        )
        .addStringOption(option =>
          option
            .setName('формат')
            .setDescription('Формат звіту')
            .setRequired(false)
            .addChoices(
              { name: 'Текст', value: 'txt' },
              { name: 'PDF', value: 'pdf' },
              { name: 'Word', value: 'docx' }
            )
        )
    ),

  async execute(interaction) {
    try {
      // Перевірка прав доступу
      const hasAccess = await checkPermission(
        interaction,
        [ROLES.SHEETS_ACCESS, ROLES.ADMIN],
        'File Manager'
      );

      if (!hasAccess) {
        return;
      }

      const subcommand = interaction.options.getSubcommand();

      // Валідація параметрів
      const options = this.extractOptions(interaction, subcommand);
      const validation = this.validateOptions(options, subcommand);

      if (!validation.isValid) {
        await interaction.reply({
          content: `❌ Помилка валідації:\n${validation.errors.join('\n')}`,
          ephemeral: true,
        });
        return;
      }

      // Логування події
      logSecurityEvent('file_manager_command_executed', {
        userId: interaction.user.id,
        userTag: interaction.user.tag,
        subcommand,
        options: validation.data,
      });

      // Відповідь про обробку
      await interaction.deferReply();

      // Виконання підкоманди
      let result;
      switch (subcommand) {
        case 'пошук':
          result = await this.handleSearch(validation.data);
          break;
        case 'читати':
          result = await this.handleRead(validation.data);
          break;
        case 'аналіз':
          result = await this.handleAnalyze(validation.data);
          break;
        case 'звіт':
          result = await this.handleReport(validation.data);
          break;
        default:
          throw new Error(`Невідома підкоманда: ${subcommand}`);
      }

      // Відправка результату
      await this.sendResult(interaction, result, subcommand);

      // Логування успішного виконання
      logger.info(`File manager command executed successfully for ${interaction.user.tag}`, {
        subcommand,
        success: true,
      });
    } catch (error) {
      logger.error('File Manager command error:', error);

      const errorMessage =
        '❌ Помилка при роботі з файлами. Спробуйте ще раз або зверніться до адміністратора.';

      if (interaction.deferred) {
        await interaction.editReply({ content: errorMessage });
      } else {
        await interaction.reply({ content: errorMessage, ephemeral: true });
      }
    }
  },

  /**
   * Витяг параметрів з interaction
   */
  extractOptions(interaction, subcommand) {
    const options = {};

    switch (subcommand) {
      case 'пошук':
        options.query = interaction.options.getString('запит');
        options.folder = interaction.options.getString('папка');
        break;
      case 'читати':
      case 'аналіз':
      case 'звіт':
        options.fileId = interaction.options.getString('id');
        break;
    }

    if (subcommand === 'аналіз') {
      options.analysisType = interaction.options.getString('тип') || 'summary';
    }

    if (subcommand === 'звіт') {
      options.format = interaction.options.getString('формат') || 'txt';
    }

    return options;
  },

  /**
   * Валідація параметрів
   */
  validateOptions(options, subcommand) {
    const schema = {};

    switch (subcommand) {
      case 'пошук':
        schema.query = {
          required: true,
          type: 'string',
          maxLength: 200,
          sanitize: 'search',
        };
        if (options.folder) {
          schema.folder = {
            required: false,
            type: 'string',
            maxLength: 50,
          };
        }
        break;
      case 'читати':
      case 'аналіз':
      case 'звіт':
        schema.fileId = {
          required: true,
          type: 'string',
          maxLength: 50,
        };
        break;
    }

    return validateCommandOptions(options, schema);
  },

  /**
   * Обробка пошуку файлів
   */
  async handleSearch(options) {
    const files = await fileProcessor.searchFiles(options.query, options.folder);

    if (files.length === 0) {
      return {
        type: 'message',
        content: `🔍 **Пошук файлів**\n\nНе знайдено файлів за запитом: "${options.query}"`,
      };
    }

    let response = `🔍 **Результати пошуку**\n\n`;
    response += `**Запит:** ${options.query}\n`;
    response += `**Знайдено:** ${files.length} файлів\n\n`;

    files.slice(0, 10).forEach((file, index) => {
      const size = file.size ? `${(file.size / 1024).toFixed(1)} KB` : 'Невідомо';
      const modified = file.modifiedTime
        ? new Date(file.modifiedTime).toLocaleDateString('uk-UA')
        : 'Невідомо';

      response += `${index + 1}. **${file.name}**\n`;
      response += `   📄 Тип: ${this.getFileTypeName(file.mimeType)}\n`;
      response += `   📏 Розмір: ${size}\n`;
      response += `   📅 Змінено: ${modified}\n`;
      response += `   🔗 [Відкрити](${file.webViewLink})\n\n`;
    });

    if (files.length > 10) {
      response += `... та ще ${files.length - 10} файлів`;
    }

    return { type: 'message', content: response };
  },

  /**
   * Обробка читання файлу
   */
  async handleRead(options) {
    const fileContent = await fileProcessor.readFileContent(options.fileId);

    let response = `📖 **Читання файлу**\n\n`;
    response += `**Назва:** ${fileContent.metadata.name}\n`;
    response += `**Тип:** ${this.getFileTypeName(fileContent.metadata.mimeType)}\n`;
    response += `**Розмір:** ${
      fileContent.metadata.size ? `${(fileContent.metadata.size / 1024).toFixed(1)} KB` : 'Невідомо'
    }\n\n`;

    // Обмеження довжини вмісту для Discord
    const maxLength = 1500;
    const content =
      fileContent.content.length > maxLength
        ? fileContent.content.substring(0, maxLength) + '...'
        : fileContent.content;

    response += `**Вміст:**\n\`\`\`\n${content}\n\`\`\``;

    return { type: 'message', content: response };
  },

  /**
   * Обробка аналізу файлу
   */
  async handleAnalyze(options) {
    const fileContent = await fileProcessor.readFileContent(options.fileId);

    // AI-аналіз вмісту
    const analysis = await aiEnhanced.analyzeData(
      [fileContent.content], // Передаємо як масив для уніфікації
      options.analysisType
    );

    let response = `🤖 **AI-аналіз файлу**\n\n`;
    response += `**Файл:** ${fileContent.metadata.name}\n`;
    response += `**Тип аналізу:** ${this.getAnalysisTypeName(options.analysisType)}\n\n`;
    response += `**Результат аналізу:**\n${analysis}`;

    return { type: 'message', content: response };
  },

  /**
   * Обробка створення звіту
   */
  async handleReport(options) {
    const fileContent = await fileProcessor.readFileContent(options.fileId);

    // Створення звіту
    const reportData = {
      title: `Звіт по файлу: ${fileContent.metadata.name}`,
      content: `Аналіз файлу "${fileContent.metadata.name}"\n\nТип файлу: ${this.getFileTypeName(
        fileContent.metadata.mimeType
      )}\nРозмір: ${
        fileContent.metadata.size
          ? `${(fileContent.metadata.size / 1024).toFixed(1)} KB`
          : 'Невідомо'
      }\n\nВміст файлу:\n${fileContent.content.substring(0, 1000)}...`,
    };

    const reportPath = await fileProcessor.createReport(reportData, options.format);

    // Читання файлу для відправки
    const fileBuffer = await fs.readFile(reportPath);
    const attachment = new AttachmentBuilder(fileBuffer, {
      name: path.basename(reportPath),
    });

    let response = `📋 **Звіт створено**\n\n`;
    response += `**Файл:** ${fileContent.metadata.name}\n`;
    response += `**Формат звіту:** ${options.format.toUpperCase()}\n`;
    response += `**Розмір звіту:** ${(fileBuffer.length / 1024).toFixed(1)} KB`;

    // Очищення тимчасового файлу
    await fileProcessor.cleanupTempFile(reportPath);

    return {
      type: 'attachment',
      content: response,
      attachment,
    };
  },

  /**
   * Відправка результату
   */
  async sendResult(interaction, result, subcommand) {
    if (result.type === 'attachment') {
      await interaction.editReply({
        content: result.content,
        files: [result.attachment],
      });
    } else {
      await interaction.editReply({
        content: result.content,
      });
    }
  },

  /**
   * Отримання назви типу файлу
   */
  getFileTypeName(mimeType) {
    const types = {
      'application/vnd.google-apps.document': 'Google Docs',
      'application/pdf': 'PDF',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'Word (DOCX)',
      'application/msword': 'Word (DOC)',
      'text/plain': 'Текстовий файл',
    };
    return types[mimeType] || 'Невідомий тип';
  },

  /**
   * Отримання назви типу аналізу
   */
  getAnalysisTypeName(type) {
    const types = {
      summary: 'Короткий зміст',
      detailed: 'Детальний аналіз',
      key_points: 'Ключові моменти',
    };
    return types[type] || type;
  },
};
