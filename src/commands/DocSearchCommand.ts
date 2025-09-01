import {
  SlashCommandBuilder,
  EmbedBuilder,
  ActionRowBuilder,
  ButtonBuilder,
  ButtonStyle,
} from 'discord.js';
import { BaseCommand, CommandExecuteOptions } from './BaseCommand';
import { tUser } from '@/i18n';
import logger from '@/utils/logger';
import { signComponentId } from '@/security/componentId';

/**
 * DocSearchCommand - Команда для пошуку в Google Docs документах
 * Дозволяє користувачам шукати інформацію в завантажених документах
 */
export class DocSearchCommand extends BaseCommand {
  constructor(config: any) {
    super(
      'doc-search',
      'Пошук в завантажених Google Docs документах',
      config,
      { category: 'documents' }
    );
  }

  protected override async onExecute(options: CommandExecuteOptions): Promise<void> {
    const interaction = options.interaction;
    try {
      // Відкладена відповідь, оскільки операція може бути тривалою
      await interaction.deferReply();

      // Отримання параметрів команди
      const query = interaction.options.getString('query', true);
      const documentId = interaction.options.getString('document_id');
      const limit = interaction.options.getInteger('limit') || 5;
      
      logger.info('🔍 Початок пошуку в документах', {
        type: 'command',
        event: 'doc_search_start',
        component: 'DocSearchCommand',
        userId: interaction.user.id,
        guildId: interaction.guildId ?? 'unknown',
        query,
        documentId,
        limit,
      });

      // Отримання сервісу Google Docs
      const googleService = (this as any).container.get('google');
      if (!googleService) {
        throw new Error('Сервіс Google не доступний');
      }

      // Спроба отримати сервіс GoogleDocsService
      let googleDocsService;
      try {
        googleDocsService = googleService.getGoogleDocsService();
      } catch (error) {
        logger.error('❌ Не вдалося отримати GoogleDocsService', {
          type: 'service_error',
          event: 'doc_search_service_failed',
          component: 'DocSearchCommand',
          error: error instanceof Error ? error.message : String(error),
        });
        
        await interaction.editReply({
          content: tUser('doc-search.service_unavailable', interaction),
        });
        return;
      }

      // Якщо вказаний конкретний документ, шукаємо лише в ньому
      if (documentId) {
        await this.searchInSpecificDocument(interaction, googleDocsService, documentId, query, limit);
      } else {
        // TODO: Реалізувати глобальний пошук по всіх проіндексованих документах
        // Поки що виводимо повідомлення про обмеження
        await interaction.editReply({
          content: tUser('doc-search.global_search_not_implemented', interaction),
        });
      }

    } catch (error) {
      logger.error('❌ Помилка виконання команди doc-search', {
        type: 'command_error',
        event: 'doc_search_failed',
        component: 'DocSearchCommand',
        userId: interaction.user.id,
        guildId: interaction.guildId ?? 'unknown',
        error: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });

      try {
        if (!interaction.replied && !interaction.deferred) {
          await interaction.reply({
            content: tUser('doc-search.error', interaction),
            ephemeral: true,
          });
        } else if (interaction.deferred) {
          await interaction.editReply({
            content: tUser('doc-search.error', interaction),
          });
        }
      } catch (replyError) {
        logger.error('❌ Не вдалося надіслати повідомлення про помилку', {
          type: 'reply_error',
          event: 'doc_search_reply_failed',
          component: 'DocSearchCommand',
          error: replyError instanceof Error ? replyError.message : String(replyError),
        });
      }
    }
  }

  /**
   * Пошук в конкретному документі
   */
  private async searchInSpecificDocument(
    interaction: any,
    googleDocsService: any,
    documentId: string,
    query: string,
    limit: number
  ): Promise<void> {
    try {
      // Отримання метаданих документа для відображення
      const googleService = (this as any).container.get('google');
      let docMetadata;
      try {
        docMetadata = await googleService.getDriveFileMetadata(documentId);
      } catch (error) {
        logger.error('❌ Не вдалося отримати метадані документа для пошуку', {
          type: 'api_error',
          event: 'doc_search_metadata_failed',
          component: 'DocSearchCommand',
          documentId,
          error: error instanceof Error ? error.message : String(error),
        });
        
        await interaction.editReply({
          content: tUser('doc-search.metadata_failed', interaction, { documentId }),
        });
        return;
      }

      // Виконання пошуку
      const searchResults = await googleDocsService.searchDoc(documentId, query);
      
      // Обмеження кількості результатів
      const limitedResults = searchResults.slice(0, limit);
      
      if (limitedResults.length === 0) {
        await interaction.editReply({
          content: tUser('doc-search.no_results', interaction, {
            query,
            documentName: docMetadata.name || documentId,
          }),
        });
        return;
      }

      // Створення вбудованих повідомлень для результатів
      const embeds: EmbedBuilder[] = [];
      
      // Основне вбудоване повідомлення з інформацією про пошук
      const mainEmbed = new EmbedBuilder()
        .setTitle(tUser('doc-search.results.title', interaction))
        .setDescription(tUser('doc-search.results.description', interaction, {
          query,
          documentName: docMetadata.name || documentId,
          resultsCount: limitedResults.length.toString(),
          totalResults: searchResults.length.toString(),
        }))
        .setColor('#0099FF')
        .setTimestamp();

      embeds.push(mainEmbed);

      // Вбудовані повідомлення для кожного результату
      for (let i = 0; i < Math.min(limitedResults.length, 3); i++) {
        const result = limitedResults[i];
        const resultEmbed = new EmbedBuilder()
          .setTitle(tUser('doc-search.result_item.title', interaction, {
            index: (i + 1).toString(),
          }))
          .setDescription(this.truncateText(result.content, 500))
          .addFields(
            {
              name: tUser('doc-search.result_item.type', interaction),
              value: result.blockType,
              inline: true,
            },
            {
              name: tUser('doc-search.result_item.relevance', interaction),
              value: `${Math.round(result.relevanceScore * 100)}%`,
              inline: true,
            }
          )
          .setColor('#00CCFF');

        embeds.push(resultEmbed);
      }

      // Якщо є більше ніж 3 результати, додаємо підсумкове повідомлення
      if (limitedResults.length > 3) {
        const summaryEmbed = new EmbedBuilder()
          .setTitle(tUser('doc-search.more_results.title', interaction))
          .setDescription(tUser('doc-search.more_results.description', interaction, {
            count: (limitedResults.length - 3).toString(),
          }))
          .setColor('#0066CC');

        embeds.push(summaryEmbed);
      }

      // Створення кнопок для додаткових дій
      const actionRow = new ActionRowBuilder<ButtonBuilder>()
        .addComponents(
          new ButtonBuilder()
            .setCustomId(signComponentId({
              kind: 'doc-search-action',
              action: 'refine',
              documentId,
              query,
              ts: Math.floor(Date.now() / 1000),
            }))
            .setLabel(tUser('doc-search.buttons.refine', interaction))
            .setStyle(ButtonStyle.Primary),
          new ButtonBuilder()
            .setCustomId(signComponentId({
              kind: 'doc-search-action',
              action: 'summary',
              documentId,
              ts: Math.floor(Date.now() / 1000),
            }))
            .setLabel(tUser('doc-search.buttons.summary', interaction))
            .setStyle(ButtonStyle.Secondary),
        );

      await interaction.editReply({
        embeds,
        components: embeds.length > 1 ? [actionRow] : [],
      });

      logger.info('✅ Успішно виконано пошук в документі', {
        type: 'command',
        event: 'doc_search_success',
        component: 'DocSearchCommand',
        userId: interaction.user.id,
        guildId: interaction.guildId ?? 'unknown',
        documentId,
        query,
        resultsCount: limitedResults.length,
        duration: Date.now() - interaction.createdTimestamp,
      });
    } catch (error) {
      logger.error('❌ Помилка пошуку в документі', {
        type: 'search_error',
        event: 'doc_search_in_document_failed',
        component: 'DocSearchCommand',
        documentId,
        query,
        error: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });
      
      throw error;
    }
  }

  /**
   * Обрізка тексту до вказаної довжини
   */
  private truncateText(text: string, maxLength: number): string {
    if (text.length <= maxLength) {
      return text;
    }
    
    return text.substring(0, maxLength - 3) + '...';
  }

  /**
   * Реєстрація слеш-команди
   */
  public register(): Omit<SlashCommandBuilder, 'addSubcommand' | 'addSubcommandGroup'> {
    const builder = new SlashCommandBuilder()
      .setName(this.name)
      .setDescription(this.description)
      .setDescriptionLocalizations({
        uk: 'Пошук в завантажених Google Docs документах',
        'en-US': 'Search in loaded Google Docs documents',
      } as any)
      .addStringOption(option =>
        option
          .setName('query')
          .setDescription('Пошуковий запит')
          .setDescriptionLocalizations({
            uk: 'Пошуковий запит',
            'en-US': 'Search query',
          } as any)
          .setRequired(true)
      )
      .addStringOption(option =>
        option
          .setName('document_id')
          .setDescription('ID документа (опціонально, для пошуку в конкретному документі)')
          .setDescriptionLocalizations({
            uk: 'ID документа (опціонально, для пошуку в конкретному документі)',
            'en-US': 'Document ID (optional, for search in specific document)',
          } as any)
          .setRequired(false)
      )
      .addIntegerOption(option =>
        option
          .setName('limit')
          .setDescription('Максимальна кількість результатів (за замовчуванням: 5)')
          .setDescriptionLocalizations({
            uk: 'Максимальна кількість результатів (за замовчуванням: 5)',
            'en-US': 'Maximum number of results (default: 5)',
          } as any)
          .setMinValue(1)
          .setMaxValue(20)
          .setRequired(false)
      )
      .setDMPermission(false);
    
    // Create a new builder with the same properties to ensure correct type
    const result = new SlashCommandBuilder()
      .setName(builder.name)
      .setDescription(builder.description)
      .setDescriptionLocalizations(builder.description_localizations ?? {})
      .setDMPermission(builder.dm_permission ?? false);
    
    // Copy options
    if (builder.options) {
      for (const option of builder.options) {
        (result as any).options.push(option);
      }
    }
    
    return result;
  }
}