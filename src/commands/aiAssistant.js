/**
 * AI-асистент команда з природномовним інтерфейсом
 * Використовує розширений AI-модуль та систему безпеки
 */

const { SlashCommandBuilder } = require('discord.js');
const {
  checkPermission,
  ROLES,
  sanitizeInput,
  validateCommandOptions,
} = require('../utils/security');
const { aiEnhanced } = require('../utils/aiEnhanced');
const { logSecurityEvent } = require('../utils/security');
const logger = require('../utils/logger');

module.exports = {
  data: new SlashCommandBuilder()
    .setName('ai')
    .setDescription('🤖 AI-асистент для роботи з Google Sheets')
    .addStringOption(option =>
      option
        .setName('запит')
        .setDescription(
          'Що ви хочете зробити? (наприклад: "знайди товари iPhone", "проаналізуй залишки")'
        )
        .setRequired(true)
        .setMaxLength(1000)
    )
    .addStringOption(option =>
      option
        .setName('контекст')
        .setDescription('Додатковий контекст для AI')
        .setRequired(false)
        .setMaxLength(500)
    ),

  async execute(interaction) {
    try {
      // Перевірка прав доступу
      const hasAccess = await checkPermission(
        interaction,
        [ROLES.AI_ACCESS, ROLES.ADMIN],
        'AI Assistant'
      );

      if (!hasAccess) {
        return;
      }

      // Валідація параметрів
      const options = {
        query: interaction.options.getString('запит'),
        context: interaction.options.getString('контекст'),
      };

      const validationSchema = {
        query: {
          required: true,
          type: 'string',
          maxLength: 1000,
          sanitize: 'ai_prompt',
        },
        context: {
          required: false,
          type: 'string',
          maxLength: 500,
          sanitize: 'ai_prompt',
        },
      };

      const validation = validateCommandOptions(options, validationSchema);
      if (!validation.isValid) {
        await interaction.reply({
          content: `❌ Помилка валідації:\n${validation.errors.join('\n')}`,
          ephemeral: true,
        });
        return;
      }

      // Логування події
      logSecurityEvent('ai_command_executed', {
        userId: interaction.user.id,
        userTag: interaction.user.tag,
        query: options.query,
        context: options.context,
      });

      // Відповідь про обробку
      await interaction.deferReply();

      // Обробка запиту через AI
      const result = await aiEnhanced.processNaturalLanguageQuery(
        interaction.user.id,
        options.query,
        null // sheetData буде передано пізніше
      );

      // Формування відповіді
      let response = `🤖 **AI-асистент**\n\n`;

      if (result.confidence < 0.7) {
        response += `⚠️ **Низька впевненість** (${Math.round(result.confidence * 100)}%)\n`;
      }

      response += `**Ваш запит:** ${options.query}\n\n`;
      response += `**Відповідь:**\n${result.response}`;

      // Додавання контексту якщо є
      if (options.context) {
        response += `\n\n**Контекст:** ${options.context}`;
      }

      // Додавання додаткової інформації
      if (result.actionData) {
        response += `\n\n**Дія:** ${result.actionData.type}`;
        if (result.actionData.format) {
          response += ` (формат: ${result.actionData.format})`;
        }
      }

      // Відправка відповіді
      await interaction.editReply({
        content: response,
        ephemeral: false,
      });

      // Логування успішного виконання
      logger.info(`AI command executed successfully for ${interaction.user.tag}`, {
        action: result.action,
        confidence: result.confidence,
        hasActionData: !!result.actionData,
      });
    } catch (error) {
      logger.error('AI Assistant command error:', error);

      const errorMessage =
        '❌ Помилка при обробці AI-запиту. Спробуйте ще раз або зверніться до адміністратора.';

      if (interaction.deferred) {
        await interaction.editReply({ content: errorMessage });
      } else {
        await interaction.reply({ content: errorMessage, ephemeral: true });
      }
    }
  },
};
