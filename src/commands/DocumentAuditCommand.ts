import { BaseCommand, type CommandExecuteOptions } from '@/commands/BaseCommand';
import type { BotConfig } from '@/types';
import type { ChatInputCommandInteraction } from 'discord.js';
import { DocumentAccessAuditService } from '@/services/DocumentAccessAuditService';
import logger from '@/utils/logger';
import { t } from '@/i18n';
import { replyWithPrivacy } from '@/ui/reply';

export class DocumentAuditCommand extends BaseCommand {
  private auditService: DocumentAccessAuditService | null = null;

  constructor(config: BotConfig) {
    super(
      'document-audit',
      t('document-audit.command.description'),
      config,
      { 
        category: 'security', 
        i18n: { 
          nameKey: 'commands.document-audit.name', 
          descriptionKey: 'document-audit.command.description' 
        } 
      },
      builder => {
        builder
          .addSubcommand(sub =>
            sub
              .setName('view')
              .setDescription(t('document-audit.sub.view.description'))
              .addStringOption(option =>
                option
                  .setName('user')
                  .setDescription(t('document-audit.opt.user.description'))
                  .setRequired(false)
              )
              .addStringOption(option =>
                option
                  .setName('file')
                  .setDescription(t('document-audit.opt.file.description'))
                  .setRequired(false)
              )
              .addStringOption(option =>
                option
                  .setName('type')
                  .setDescription(t('document-audit.opt.type.description'))
                  .setRequired(false)
                  .addChoices(
                    { name: t('document-audit.choices.type.view'), value: 'view' },
                    { name: t('document-audit.choices.type.edit'), value: 'edit' },
                    { name: t('document-audit.choices.type.download'), value: 'download' },
                    { name: t('document-audit.choices.type.share'), value: 'share' },
                    { name: t('document-audit.choices.type.delete'), value: 'delete' }
                  )
              )
              .addIntegerOption(option =>
                option
                  .setName('limit')
                  .setDescription(t('document-audit.opt.limit.description'))
                  .setRequired(false)
                  .setMinValue(1)
                  .setMaxValue(100)
              )
          )
          .addSubcommand(sub =>
            sub
              .setName('stats')
              .setDescription(t('document-audit.sub.stats.description'))
              .addStringOption(option =>
                option
                  .setName('user')
                  .setDescription(t('document-audit.opt.user.description'))
                  .setRequired(false)
              )
              .addStringOption(option =>
                option
                  .setName('file')
                  .setDescription(t('document-audit.opt.file.description'))
                  .setRequired(false)
              )
              .addStringOption(option =>
                option
                  .setName('type')
                  .setDescription(t('document-audit.opt.type.description'))
                  .setRequired(false)
                  .addChoices(
                    { name: t('document-audit.choices.type.view'), value: 'view' },
                    { name: t('document-audit.choices.type.edit'), value: 'edit' },
                    { name: t('document-audit.choices.type.download'), value: 'download' },
                    { name: t('document-audit.choices.type.share'), value: 'share' },
                    { name: t('document-audit.choices.type.delete'), value: 'delete' }
                  )
              )
          )
          .addSubcommand(sub =>
            sub
              .setName('export')
              .setDescription(t('document-audit.sub.export.description'))
              .addStringOption(option =>
                option
                  .setName('user')
                  .setDescription(t('document-audit.opt.user.description'))
                  .setRequired(false)
              )
              .addStringOption(option =>
                option
                  .setName('file')
                  .setDescription(t('document-audit.opt.file.description'))
                  .setRequired(false)
              )
              .addStringOption(option =>
                option
                  .setName('type')
                  .setDescription(t('document-audit.opt.type.description'))
                  .setRequired(false)
                  .addChoices(
                    { name: t('document-audit.choices.type.view'), value: 'view' },
                    { name: t('document-audit.choices.type.edit'), value: 'edit' },
                    { name: t('document-audit.choices.type.download'), value: 'download' },
                    { name: t('document-audit.choices.type.share'), value: 'share' },
                    { name: t('document-audit.choices.type.delete'), value: 'delete' }
                  )
              )
          );
        return builder;
      }
    );
  }

  /**
   * Initialize the command with required services
   */
  initializeServices(auditService: DocumentAccessAuditService): void {
    this.auditService = auditService;
  }

  /**
   * Execute the command
   */
  protected async onExecute(options: CommandExecuteOptions): Promise<void> {
    const { interaction } = options;
    
    try {
      if (!this.auditService) {
        await replyWithPrivacy(interaction, {
          content: '❌ Audit service not initialized',
          ephemeral: true
        });
        return;
      }

      const subcommand = interaction.options.getSubcommand();
      
      switch (subcommand) {
        case 'view':
          await this.handleViewLogs(interaction);
          break;
        case 'stats':
          await this.handleViewStats(interaction);
          break;
        case 'export':
          await this.handleExportLogs(interaction);
          break;
        default:
          await replyWithPrivacy(interaction, {
            content: '❌ Unknown subcommand',
            ephemeral: true
          });
      }
    } catch (error) {
      logger.error('Error executing document audit command', {
        component: 'DocumentAuditCommand',
        error: error instanceof Error ? error.message : String(error)
      });
      
      await replyWithPrivacy(interaction, {
        content: '❌ An error occurred while processing your request',
        ephemeral: true
      });
    }
  }

  /**
   * Handle viewing access logs
   */
  private async handleViewLogs(interaction: ChatInputCommandInteraction): Promise<void> {
    await replyWithPrivacy(interaction, { content: '⏳ Processing...' });

    try {
      const userId = interaction.options.getString('user') || undefined;
      const fileId = interaction.options.getString('file') || undefined;
      const accessType = interaction.options.getString('type') || undefined;
      const limit = interaction.options.getInteger('limit') || 10;

      const logs = await this.auditService.getAccessLogs({
        userId,
        fileId,
        accessType,
        limit
      });

      if (logs.length === 0) {
        await replyWithPrivacy(interaction, {
          content: t('document-audit.view.noLogs')
        });
        return;
      }

      // Format logs for display
      const logEntries = logs.map(log => {
        const timestamp = log.timestamp.toLocaleString();
        const status = log.success ? t('document-audit.view.success') : t('document-audit.view.failure');
        const size = log.fileSize ? `${(log.fileSize / 1024).toFixed(1)}KB` : 'N/A';
        
        return `**${log.fileName}** (${log.fileId.substring(0, 8)}...)
${status} ${log.accessType} by ${log.userName} (${log.userId.substring(0, 8)}...)
📅 ${timestamp}
💾 ${size} | 📄 ${log.fileType || 'N/A'}`;
      });

      const response = `## 📋 ${t('document-audit.view.title')} (${logs.length} entries)

${logEntries.join('\n\n')}`;

      await replyWithPrivacy(interaction, {
        content: response
      });
    } catch (error) {
      logger.error('Error viewing access logs', {
        component: 'DocumentAuditCommand',
        error: error instanceof Error ? error.message : String(error)
      });
      
      await replyWithPrivacy(interaction, {
        content: t('document-audit.export.failure')
      });
    }
  }

  /**
   * Handle viewing access statistics
   */
  private async handleViewStats(interaction: ChatInputCommandInteraction): Promise<void> {
    await replyWithPrivacy(interaction, { content: '⏳ Processing...' });

    try {
      const userId = interaction.options.getString('user') || undefined;
      const fileId = interaction.options.getString('file') || undefined;
      const accessType = interaction.options.getString('type') || undefined;

      const stats = await this.auditService.getAccessStats({
        userId,
        fileId,
        accessType
      });

      // Format access by type
      const accessByType = Object.entries(stats.accessByType)
        .map(([type, count]) => `${type}: ${count}`)
        .join('\n');

      // Format access by time
      const accessByTime = Object.entries(stats.accessByTime)
        .sort(([a], [b]) => parseInt(a) - parseInt(b))
        .map(([hour, count]) => `${hour}:00 - ${count}`)
        .join('\n');

      const response = `## 📊 ${t('document-audit.stats.title')}

**${t('document-audit.stats.totalAccesses')}:** ${stats.totalAccesses}
**${t('document-audit.stats.successful')}:** ${stats.successfulAccesses}
**${t('document-audit.stats.failed')}:** ${stats.failedAccesses}
**${t('document-audit.stats.uniqueUsers')}:** ${stats.uniqueUsers}

### ${t('document-audit.stats.byType')}:
${accessByType || t('document-audit.stats.noData')}

### ${t('document-audit.stats.byTime')}:
${accessByTime || t('document-audit.stats.noData')}`;

      await replyWithPrivacy(interaction, {
        content: response
      });
    } catch (error) {
      logger.error('Error viewing access stats', {
        component: 'DocumentAuditCommand',
        error: error instanceof Error ? error.message : String(error)
      });
      
      await replyWithPrivacy(interaction, {
        content: t('document-audit.export.failure')
      });
    }
  }

  /**
   * Handle exporting access logs
   */
  private async handleExportLogs(interaction: ChatInputCommandInteraction): Promise<void> {
    await replyWithPrivacy(interaction, { content: '⏳ Processing...' });

    try {
      const userId = interaction.options.getString('user') || undefined;
      const fileId = interaction.options.getString('file') || undefined;
      const accessType = interaction.options.getString('type') || undefined;

      const exportedData = await this.auditService.exportAccessLogs({
        userId,
        fileId,
        accessType
      });

      // In a real implementation, this would create and send a file
      // For now, we'll just show a confirmation message
      await replyWithPrivacy(interaction, {
        content: `${t('document-audit.export.success')}
📋 ${exportedData.length} characters of data prepared for export.`
      });
    } catch (error) {
      logger.error('Error exporting access logs', {
        component: 'DocumentAuditCommand',
        error: error instanceof Error ? error.message : String(error)
      });
      
      await replyWithPrivacy(interaction, {
        content: t('document-audit.export.failure')
      });
    }
  }
}