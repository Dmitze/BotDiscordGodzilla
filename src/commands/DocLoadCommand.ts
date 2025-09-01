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
 * DocLoadCommand - Команда для завантаження та індексації Google Docs документів
 * Дозволяє користувачам завантажувати документи для подальшого пошуку та аналізу
 */
export class DocLoadCommand extends BaseCommand {
  constructor(config: any) {
    super(
      'doc-load',
      'Завантажити та проіндексувати Google Docs документ',
      config,
      { category: 'documents' }
    );
  }

  protected override async onExecute(options: CommandExecuteOptions): Promise<void> {
    const interaction = options.interaction;
    try {
      // Відкладена відповідь, оскільки операція може бути тривалою
      await interaction.deferReply({ ephemeral: true });

      // Отримання параметрів команди
      const documentUrl = interaction.options.getString('document_url', true);
      const folderId = interaction.options.getString('folder_id');
      
      logger.info('📥 Початок завантаження документа', {
        type: 'command',
        event: 'doc_load_start',
        component: 'DocLoadCommand',
        userId: interaction.user.id,
        guildId: interaction.guildId ?? 'unknown',
        documentUrl,
        folderId,
      });

      // Вилучення ID документа з URL
      const documentId = this.extractDocumentId(documentUrl);
      
      if (!documentId) {
        await interaction.editReply({
          content: tUser('doc-load.invalid_url', interaction, { url: documentUrl }),
        });
        return;
      }

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
          event: 'doc_load_service_failed',
          component: 'DocLoadCommand',
          error: error instanceof Error ? error.message : String(error),
        });
        
        await interaction.editReply({
          content: tUser('doc-load.service_unavailable', interaction),
        });
        return;
      }

      // Отримання метаданих документа
      let docMetadata;
      try {
        docMetadata = await googleService.getDriveFileMetadata(documentId);
      } catch (error) {
        logger.error('❌ Не вдалося отримати метадані документа', {
          type: 'api_error',
          event: 'doc_load_metadata_failed',
          component: 'DocLoadCommand',
          documentId,
          error: error instanceof Error ? error.message : String(error),
        });
        
        await interaction.editReply({
          content: tUser('doc-load.metadata_failed', interaction, { documentId }),
        });
        return;
      }

      // Перевірка, чи це дійсно Google Docs документ
      if (docMetadata.mimeType !== 'application/vnd.google-apps.document') {
        await interaction.editReply({
          content: tUser('doc-load.not_google_doc', interaction, { 
            mimeType: docMetadata.mimeType 
          }),
        });
        return;
      }

      // Індексація документа
      const indexResult = await googleDocsService.indexDoc(documentId);
      
      if (!indexResult.success) {
        await interaction.editReply({
          content: tUser('doc-load.index_failed', interaction, { documentId }),
        });
        return;
      }

      // Створення вбудованого повідомлення з результатами
      const embed = new EmbedBuilder()
        .setTitle(tUser('doc-load.success.title', interaction))
        .setDescription(tUser('doc-load.success.description', interaction, {
          documentName: docMetadata.name || documentId,
          wordCount: indexResult.wordCount.toString(),
        }))
        .addFields(
          {
            name: tUser('doc-load.success.fields.document_id', interaction),
            value: documentId,
            inline: true,
          },
          {
            name: tUser('doc-load.success.fields.indexed_at', interaction),
            value: `<t:${Math.floor(new Date(indexResult.indexedAt).getTime() / 1000)}:F>`,
            inline: true,
          },
          {
            name: tUser('doc-load.success.fields.content_hash', interaction),
            value: `\`${indexResult.contentHash.substring(0, 8)}...\``,
            inline: true,
          }
        )
        .setColor('#00FF00')
        .setTimestamp();

      // Створення кнопок для додаткових дій
      const actionRow = new ActionRowBuilder<ButtonBuilder>()
        .addComponents(
          new ButtonBuilder()
            .setCustomId(signComponentId({
              kind: 'doc-action',
              action: 'search',
              documentId,
              ts: Math.floor(Date.now() / 1000),
            }))
            .setLabel(tUser('doc-load.buttons.search', interaction))
            .setStyle(ButtonStyle.Primary),
          new ButtonBuilder()
            .setCustomId(signComponentId({
              kind: 'doc-action',
              action: 'summary',
              documentId,
              ts: Math.floor(Date.now() / 1000),
            }))
            .setLabel(tUser('doc-load.buttons.summary', interaction))
            .setStyle(ButtonStyle.Secondary),
          new ButtonBuilder()
            .setCustomId(signComponentId({
              kind: 'doc-action',
              action: 'info',
              documentId,
              ts: Math.floor(Date.now() / 1000),
            }))
            .setLabel(tUser('doc-load.buttons.info', interaction))
            .setStyle(ButtonStyle.Secondary),
        );

      await interaction.editReply({
        content: tUser('doc-load.success.message', interaction, {
          documentName: docMetadata.name || documentId,
        }),
        embeds: [embed],
        components: [actionRow],
      });

      logger.info('✅ Успішно завершено завантаження документа', {
        type: 'command',
        event: 'doc_load_success',
        component: 'DocLoadCommand',
        userId: interaction.user.id,
        guildId: interaction.guildId ?? 'unknown',
        documentId,
        wordCount: indexResult.wordCount,
        duration: Date.now() - interaction.createdTimestamp,
      });
    } catch (error) {
      logger.error('❌ Помилка виконання команди doc-load', {
        type: 'command_error',
        event: 'doc_load_failed',
        component: 'DocLoadCommand',
        userId: interaction.user.id,
        guildId: interaction.guildId ?? 'unknown',
        error: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });

      try {
        if (!interaction.replied && !interaction.deferred) {
          await interaction.reply({
            content: tUser('doc-load.error', interaction),
            ephemeral: true,
          });
        } else if (interaction.deferred) {
          await interaction.editReply({
            content: tUser('doc-load.error', interaction),
          });
        }
      } catch (replyError) {
        logger.error('❌ Не вдалося надіслати повідомлення про помилку', {
          type: 'reply_error',
          event: 'doc_load_reply_failed',
          component: 'DocLoadCommand',
          error: replyError instanceof Error ? replyError.message : String(replyError),
        });
      }
    }
  }

  /**
   * Вилучення ID документа з URL Google Docs
   * @param url - URL Google Docs документа
   * @returns ID документа або null, якщо не вдалося вилучити
   */
  private extractDocumentId(url: string): string | null {
    try {
      // Спроба розпарсити URL
      const parsedUrl = new URL(url);
      
      // Перевірка, чи це Google Docs URL
      if (parsedUrl.hostname !== 'docs.google.com') {
        return null;
      }
      
      // Вилучення ID з шляху
      // Приклад: https://docs.google.com/document/d/1a2b3c4d5e6f7g8h9i0j/edit
      const match = parsedUrl.pathname.match(/\/document\/d\/([a-zA-Z0-9_-]+)/);
      
      if (match && match[1]) {
        return match[1];
      }
      
      return null;
    } catch (error) {
      logger.debug('⚠️ Не вдалося розпарсити URL документа', {
        type: 'parsing',
        event: 'doc_url_parse_failed',
        component: 'DocLoadCommand',
        url,
        error: error instanceof Error ? error.message : String(error),
      });
      return null;
    }
  }

  /**
   * Реєстрація слеш-команди
   */
  public register(): Omit<SlashCommandBuilder, 'addSubcommand' | 'addSubcommandGroup'> {
    return new SlashCommandBuilder()
      .setName(this.name)
      .setDescription(this.description)
      .setDescriptionLocalizations({
        uk: 'Завантажити та проіндексувати Google Docs документ',
        'en-US': 'Load and index Google Docs document',
      } as any)
      .addStringOption(option =>
        option
          .setName('document_url')
          .setDescription('URL Google Docs документа')
          .setDescriptionLocalizations({
            uk: 'URL Google Docs документа',
            'en-US': 'Google Docs document URL',
          } as any)
          .setRequired(true)
      )
      .addStringOption(option =>
        option
          .setName('folder_id')
          .setDescription('ID папки Google Drive (опціонально)')
          .setDescriptionLocalizations({
            uk: 'ID папки Google Drive (опціонально)',
            'en-US': 'Google Drive folder ID (optional)',
          } as any)
          .setRequired(false)
      )
      .setDMPermission(false) as Omit<SlashCommandBuilder, 'addSubcommand' | 'addSubcommandGroup'>;
  }
}