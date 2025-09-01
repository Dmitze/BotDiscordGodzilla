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
 * DocSummaryCommand - Команда для генерації резюме Google Docs документів
 * Дозволяє користувачам отримати коротке резюме документа
 */
export class DocSummaryCommand extends BaseCommand {
  constructor(config: any) {
    super(
      'doc-summary',
      'Згенерувати резюме Google Docs документа',
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
      const documentId = interaction.options.getString('document_id', true);
      
      logger.info('📝 Початок генерації резюме документа', {
        type: 'command',
        event: 'doc_summary_start',
        component: 'DocSummaryCommand',
        userId: interaction.user.id,
        guildId: interaction.guildId ?? 'unknown',
        documentId,
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
          event: 'doc_summary_service_failed',
          component: 'DocSummaryCommand',
          error: error instanceof Error ? error.message : String(error),
        });
        
        await interaction.editReply({
          content: tUser('doc-summary.service_unavailable', interaction),
        });
        return;
      }

      // Отримання метаданих документа
      let docMetadata;
      try {
        docMetadata = await googleService.getDriveFileMetadata(documentId);
      } catch (error) {
        logger.error('❌ Не вдалося отримати метадані документа для резюме', {
          type: 'api_error',
          event: 'doc_summary_metadata_failed',
          component: 'DocSummaryCommand',
          documentId,
          error: error instanceof Error ? error.message : String(error),
        });
        
        await interaction.editReply({
          content: tUser('doc-summary.metadata_failed', interaction, { documentId }),
        });
        return;
      }

      // Перевірка, чи це дійсно Google Docs документ
      if (docMetadata.mimeType !== 'application/vnd.google-apps.document') {
        await interaction.editReply({
          content: tUser('doc-summary.not_google_doc', interaction, { 
            mimeType: docMetadata.mimeType 
          }),
        });
        return;
      }

      // Генерація резюме
      const summaryResult = await googleDocsService.summarizeDoc(documentId);

      // Створення вбудованого повідомлення з резюме
      const mainEmbed = new EmbedBuilder()
        .setTitle(tUser('doc-summary.title', interaction, {
          documentName: summaryResult.title,
        }))
        .setDescription(summaryResult.summary || tUser('doc-summary.no_summary', interaction))
        .addFields(
          {
            name: tUser('doc-summary.fields.word_count', interaction),
            value: summaryResult.wordCount.toString(),
            inline: true,
          },
          {
            name: tUser('doc-summary.fields.reading_time', interaction),
            value: tUser('doc-summary.reading_time_value', interaction, {
              minutes: summaryResult.readingTimeMinutes.toString(),
            }),
            inline: true,
          }
        )
        .setColor('#FF6600')
        .setTimestamp();

      // Додавання ключових точок, якщо вони є
      if (summaryResult.keyPoints.length > 0) {
        const keyPointsText = summaryResult.keyPoints
          .slice(0, 10) // Обмежуємо 10 ключовими точками
          .map((point: string, index: number) => `${index + 1}. ${this.truncateText(point, 100)}`)
          .join('\n');
        
        mainEmbed.addFields({
          name: tUser('doc-summary.fields.key_points', interaction),
          value: keyPointsText,
        });
      }

      // Створення кнопок для додаткових дій
      const actionRow = new ActionRowBuilder<ButtonBuilder>()
        .addComponents(
          new ButtonBuilder()
            .setCustomId(signComponentId({
              kind: 'doc-summary-action',
              action: 'search',
              documentId,
              ts: Math.floor(Date.now() / 1000),
            }))
            .setLabel(tUser('doc-summary.buttons.search', interaction))
            .setStyle(ButtonStyle.Primary),
          new ButtonBuilder()
            .setCustomId(signComponentId({
              kind: 'doc-summary-action',
              action: 'details',
              documentId,
              ts: Math.floor(Date.now() / 1000),
            }))
            .setLabel(tUser('doc-summary.buttons.details', interaction))
            .setStyle(ButtonStyle.Secondary),
          new ButtonBuilder()
            .setCustomId(signComponentId({
              kind: 'doc-summary-action',
              action: 'export',
              documentId,
              ts: Math.floor(Date.now() / 1000),
            }))
            .setLabel(tUser('doc-summary.buttons.export', interaction))
            .setStyle(ButtonStyle.Secondary),
        );

      await interaction.editReply({
        embeds: [mainEmbed],
        components: [actionRow],
      });

      logger.info('✅ Успішно згенеровано резюме документа', {
        type: 'command',
        event: 'doc_summary_success',
        component: 'DocSummaryCommand',
        userId: interaction.user.id,
        guildId: interaction.guildId ?? 'unknown',
        documentId,
        wordCount: summaryResult.wordCount,
        readingTime: summaryResult.readingTimeMinutes,
        keyPointsCount: summaryResult.keyPoints.length,
        duration: Date.now() - interaction.createdTimestamp,
      });
    } catch (error) {
      logger.error('❌ Помилка виконання команди doc-summary', {
        type: 'command_error',
        event: 'doc_summary_failed',
        component: 'DocSummaryCommand',
        userId: interaction.user.id,
        guildId: interaction.guildId ?? 'unknown',
        documentId: interaction.options.getString('document_id'),
        error: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });

      try {
        if (!interaction.replied && !interaction.deferred) {
          await interaction.reply({
            content: tUser('doc-summary.error', interaction),
            ephemeral: true,
          });
        } else if (interaction.deferred) {
          await interaction.editReply({
            content: tUser('doc-summary.error', interaction),
          });
        }
      } catch (replyError) {
        logger.error('❌ Не вдалося надіслати повідомлення про помилку', {
          type: 'reply_error',
          event: 'doc_summary_reply_failed',
          component: 'DocSummaryCommand',
          error: replyError instanceof Error ? replyError.message : String(replyError),
        });
      }
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
        uk: 'Згенерувати резюме Google Docs документа',
        'en-US': 'Generate summary of Google Docs document',
      } as any)
      .addStringOption(option =>
        option
          .setName('document_id')
          .setDescription('ID Google Docs документа')
          .setDescriptionLocalizations({
            uk: 'ID Google Docs документа',
            'en-US': 'Google Docs document ID',
          } as any)
          .setRequired(true)
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