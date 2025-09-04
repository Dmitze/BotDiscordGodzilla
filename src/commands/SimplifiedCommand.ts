/**
 * Спрощена команда з трьома основними функціями
 * 1. AI Асистент
 * 2. Пошук у Google Drive
 * 3. OCR розпізнавання тексту
 */

import type { BotConfig, CommandExecuteOptions } from '@/types';
import { BaseCommand } from './BaseCommand';
import logger from '@/utils/logger';

import type { GoogleService } from '@/services/GoogleService';
import type { AIService } from '@/services/AIService';
import type {
  SlashCommandBuilder,
  SlashCommandStringOption,
  ChatInputCommandInteraction,
} from 'discord.js';
import { t } from '@/i18n';
import { buildSimpleActionRow, buildSimplePaginationRow, buildCloseRow, chunkTextForDiscord } from '@/ui/simpleComponents';
import { EmbedBuilder } from 'discord.js';

type CommandAction = 'ai' | 'search' | 'ocr' | 'close' | 'page';

interface CommandState {
  action: CommandAction;
  page: number;
  fileId?: string;
  [key: string]: any; // This allows any additional properties
}

export class SimplifiedCommand extends BaseCommand {
  private readonly googleService: GoogleService | undefined;
  private readonly aiService: AIService | undefined;
  
  constructor(
    config: BotConfig, 
    googleService?: GoogleService,
    aiService?: AIService
  ) {
    super('bot', t('simplified.command.description'), config, {
      i18n: { nameKey: 'commands.simplified.name', descriptionKey: 'simplified.command.description' }
    }, (builder: SlashCommandBuilder): SlashCommandBuilder => {
      builder.addStringOption((option: SlashCommandStringOption) =>
        option
          .setName('action')
          .setDescription(t('simplified.opt.action.description'))
          .setRequired(false)
          .addChoices(
            { name: '🤖 AI Асистент', value: 'ai' },
            { name: '🔍 Пошук у Google Drive', value: 'search' },
            { name: '📝 Розпізнати текст', value: 'ocr' }
          )
      );
      builder.addStringOption((option: SlashCommandStringOption) =>
        option
          .setName('query')
          .setDescription(t('simplified.opt.query.description'))
          .setRequired(false)
          .setMaxLength(1000)
      );
      return builder;
    });

    this.googleService = googleService;
    this.aiService = aiService;
  }

  /**
   * Генерація унікального ID для кнопок
   */
  private buildId(args: { action: CommandAction; page?: number; fileId?: string; timestamp?: number }): string {
    const timestamp = args.timestamp ?? Date.now();
    return `simplified_${args.action}_${args.page ?? 1}_${args.fileId ?? 'none'}_${timestamp}`;
  }

  /**
   * Парсинг ID кнопки
   */
  private parseId(customId: string): CommandState | null {
    const parts = customId.split('_');
    if (parts.length < 4) return null;
    
    if (parts[0] !== 'simplified') return null;
    
    const action = parts[1] as CommandAction;
    const page = parseInt(parts[2] ?? '1', 10) || 1;
    const fileId = parts[3] === 'none' ? undefined : parts[3];
    
    // Fix: Ensure page is a valid number and handle undefined fileId properly
    if (isNaN(page) || page < 1) return null;
    
    return { action, page, fileId: fileId !== undefined ? fileId : '' };
  }

  /**
   * Виконання команди
   */
  protected override async onExecute(options: CommandExecuteOptions): Promise<void> {
    try {
      // Отримуємо параметри
      const action = options.interaction.options.getString('action') as CommandAction | null;
      const query = options.interaction.options.getString('query');
      
      // Якщо вказана дія, виконуємо її
      if (action) {
        switch (action) {
          case 'ai':
            await this.handleAI(options.interaction, query);
            break;
          case 'search':
            await this.handleSearch(options.interaction, query);
            break;
          case 'ocr':
            await this.handleOCR(options.interaction);
            break;
        }
        return;
      }
      
      // Якщо дія не вказана, показуємо головне меню
      await this.showMainMenu(options.interaction);
    } catch (error) {
      logger.error('❌ Помилка виконання спрощеної команди:', { error });
      await this.sendError(options.interaction, 'Виникла помилка під час виконання команди');
    }
  }

  /**
   * Показ головного меню
   */
  private async showMainMenu(interaction: ChatInputCommandInteraction): Promise<void> {
    // Fix: Add description property with proper type
    const embed = new EmbedBuilder()
      .setTitle('🤖 Discord AI Бот')
      .setDescription('Оберіть одну з доступних функцій:')
      .addFields([
        {
          name: '🤖 AI Асистент',
          value: 'Задайте питання або отримайте допомогу з аналізу документів',
          inline: false
        },
        {
          name: '🔍 Пошук у Google Drive',
          value: 'Шукайте файли у вашому Google Drive',
          inline: false
        },
        {
          name: '📝 Розпізнати текст',
          value: 'Розпізнавайте текст з зображень',
          inline: false
        }
      ])
      .setColor(0x0099ff);

    const row = buildSimpleActionRow({ buildId: this.buildId.bind(this) });
    
    await interaction.reply({ embeds: [embed], components: [row] });
  }

  /**
   * Обробка AI запиту
   */
  private async handleAI(interaction: ChatInputCommandInteraction, query?: string | null): Promise<void> {
    if (!this.aiService) {
      await this.sendError(interaction, 'AI сервіс недоступний');
      return;
    }

    if (!query) {
      // Якщо запит не вказаний, просимо ввести його
      const embed = new EmbedBuilder()
        .setTitle('🤖 AI Асистент')
        .setDescription('Введіть ваше питання:')
        .setColor(0x0099ff);

      await interaction.reply({ embeds: [embed] });
      return;
    }

    try {
      await interaction.deferReply();
      
      // Генеруємо відповідь
      const response = await this.aiService.generateResponse(query);
      
      // Розбиваємо відповідь на частини, якщо вона занадто довга
      const chunks = chunkTextForDiscord(response.content);
      
      if (chunks.length === 1) {
        const embed = new EmbedBuilder()
          .setTitle('🤖 AI Відповідь')
          .setDescription(chunks[0] ?? '') // Fix: Handle undefined description
          .setColor(0x00ff00);
        
        await interaction.editReply({ embeds: [embed] });
      } else {
        // Для довгих відповідей показуємо першу частину з кнопками навігації
        const embed = new EmbedBuilder()
          .setTitle('🤖 AI Відповідь (1/' + chunks.length + ')')
          .setDescription(chunks[0] ?? '')
          .setColor(0x00ff00);
        
        const row = buildSimplePaginationRow({
          buildId: this.buildId.bind(this),
          currentPage: 1,
          totalPages: chunks.length,
          baseAction: 'ai' as 'search' // Type cast to match expected type
        });
        
        const closeRow = buildCloseRow({ buildId: this.buildId.bind(this) });
        
        await interaction.editReply({ 
          embeds: [embed], 
          components: [row, closeRow] 
        });
      }
    } catch (error) {
      logger.error('❌ Помилка AI запиту:', { error });
      await this.sendError(interaction, 'Виникла помилка під час обробки AI запиту');
    }
  }

  /**
   * Обробка пошуку
   */
  private async handleSearch(interaction: ChatInputCommandInteraction, query?: string | null): Promise<void> {
    if (!this.googleService) {
      await this.sendError(interaction, 'Google сервіс недоступний');
      return;
    }

    if (!query) {
      // Якщо запит не вказаний, просимо ввести його
      const embed = new EmbedBuilder()
        .setTitle('🔍 Пошук у Google Drive')
        .setDescription('Введіть пошуковий запит:')
        .setColor(0x0099ff);

      await interaction.reply({ embeds: [embed] });
      return;
    }

    try {
      await interaction.deferReply();
      
      // Виконуємо пошук
      const folderId = this.config.drive.folderId;
      const results = await this.googleService.listDriveFiles({
        folderId,
        query: query
      });
      
      // Розбиваємо результати на частини для пагінації
      const chunks = chunkTextForDiscord(JSON.stringify(results.files));
      
      if (results.files.length === 0) {
        const embed = new EmbedBuilder()
          .setTitle('🔍 Пошук у Google Drive')
          .setDescription('Нічого не знайдено за запитом: ' + query)
          .setColor(0xff0000);
        
        await interaction.editReply({ embeds: [embed] });
        return;
      }
      
      // Форматуємо результати
      let description = `Знайдено ${results.files.length} файлів:\n\n`;
      
      for (let i = 0; i < Math.min(10, results.files.length); i++) {
        const file = results.files[i];
        // Fix: Check if file is defined before accessing properties
        if (file) {
          description += `**${file.name || 'Без назви'}**\n`;
          description += `Тип: ${file.mimeType || 'Невідомо'}\n`;
          if (file.webViewLink) {
            description += `[Відкрити у Google Drive](${file.webViewLink})\n`;
          }
          description += '\n';
        }
      }
      
      if (results.files.length > 10) {
        description += `...і ще ${results.files.length - 10} файлів`;
      }
      
      const embed = new EmbedBuilder()
        .setTitle('🔍 Результати пошуку')
        .setDescription(description || '') // Fix: Handle undefined description
        .setColor(0x00ff00);
      
      const row = buildSimplePaginationRow({
        buildId: this.buildId.bind(this),
        currentPage: 1,
        totalPages: chunks.length,
        baseAction: 'search' as 'search' // Fix: Use 'search' instead of 'ai' to match expected type
      });
      
      const closeRow = buildCloseRow({ buildId: this.buildId.bind(this) });
      
      await interaction.editReply({ 
        embeds: [embed], 
        components: [row, closeRow] 
      });
    } catch (error) {
      logger.error('❌ Помилка пошуку:', { error });
      await this.sendError(interaction, 'Виникла помилка під час пошуку');
    }
  }

  /**
   * Обробка OCR
   */
  private async handleOCR(interaction: ChatInputCommandInteraction): Promise<void> {
    const embed = {
      title: '📝 Розпізнавання тексту',
      description: 'Для розпізнавання тексту завантажте зображення з текстом',
      color: 0x0099ff
    };

    await interaction.reply({ embeds: [embed] });
  }

  /**
   * Відправка помилки
   */
  private async sendError(interaction: ChatInputCommandInteraction, message: string): Promise<void> {
    const embed = new EmbedBuilder()
      .setTitle('❌ Помилка')
      .setDescription(message || '') // Fix: Handle undefined description
      .setColor(0xff0000);

    if (interaction.replied || interaction.deferred) {
      await interaction.editReply({ embeds: [embed] });
    } else {
      await interaction.reply({ embeds: [embed], ephemeral: true });
    }
  }

  /**
   * Обробка кнопок
   */
  public async handleButtonInteraction(interaction: any): Promise<void> {
    try {
      const state = this.parseId(interaction.customId);
      if (!state) {
        await interaction.reply({ 
          content: 'Невідома дія', 
          ephemeral: true 
        });
        return;
      }

      switch (state.action) {
        case 'ai':
          // Fix: Define embed variable before using it
          const aiEmbed = new EmbedBuilder()
            .setTitle('🤖 AI Відповідь')
            .setDescription('AI відповідь тут')
            .setColor(0x00ff00);
          await interaction.update({ embeds: [aiEmbed] });
          break;
        case 'search':
          // Fix: Define embed variable before using it
          const searchEmbed = new EmbedBuilder()
            .setTitle('🔍 Результати пошуку')
            .setDescription('Результати пошуку тут')
            .setColor(0x00ff00);
          await interaction.update({ 
            embeds: [searchEmbed],
            components: [
              buildSimplePaginationRow({
                buildId: this.buildId.bind(this),
                currentPage: state.page,
                totalPages: state['totalPages'] || 1,
                baseAction: 'search' // Fix: Remove type assertion
              }),
              buildCloseRow({ buildId: this.buildId.bind(this) })
            ]
          });
          break;
        case 'ocr':
          await this.handleOCR(interaction);
          break;
        case 'close':
          await interaction.message.delete();
          break;
        case 'page':
          await interaction.reply({ 
            content: 'Навігація між сторінками', 
            ephemeral: true 
          });
          break;
      }
    } catch (error) {
      logger.error('❌ Помилка обробки кнопки:', { error });
      await interaction.reply({ 
        content: 'Виникла помилка під час обробки дії', 
        ephemeral: true 
      });
    }
  }
}