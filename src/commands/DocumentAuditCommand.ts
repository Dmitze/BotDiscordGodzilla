import { BaseCommand } from '@/core/BaseCommand';
import type { CommandInteraction } from '@/types';
import { DocumentAccessAuditService } from '@/services/DocumentAccessAuditService';
import logger from '@/utils/logger';
import i18n from '@/i18n';

export class DocumentAuditCommand extends BaseCommand {
  private auditService: DocumentAccessAuditService | null = null;

  constructor() {
    super('document-audit', 'Audit document access and view security reports');
    
    // Add subcommands
    this.addSubcommand('view', 'View recent document access logs')
      .addStringOption(option => 
        option.setName('user')
          .setDescription('Filter by user ID')
          .setRequired(false)
      )
      .addStringOption(option => 
        option.setName('file')
          .setDescription('Filter by file ID')
          .setRequired(false)
      )
      .addStringOption(option => 
        option.setName('type')
          .setDescription('Filter by access type')
          .setRequired(false)
          .addChoices(
            { name: 'View', value: 'view' },
            { name: 'Edit', value: 'edit' },
            { name: 'Download', value: 'download' },
            { name: 'Share', value: 'share' },
            { name: 'Delete', value: 'delete' }
          )
      )
      .addIntegerOption(option => 
        option.setName('limit')
          .setDescription('Number of logs to show (max 100)')
          .setRequired(false)
          .setMinValue(1)
          .setMaxValue(100)
      );

    this.addSubcommand('stats', 'View document access statistics')
      .addStringOption(option => 
        option.setName('user')
          .setDescription('Filter by user ID')
          .setRequired(false)
      )
      .addStringOption(option => 
        option.setName('file')
          .setDescription('Filter by file ID')
          .setRequired(false)
      )
      .addStringOption(option => 
        option.setName('type')
          .setDescription('Filter by access type')
          .setRequired(false)
          .addChoices(
            { name: 'View', value: 'view' },
            { name: 'Edit', value: 'edit' },
            { name: 'Download', value: 'download' },
            { name: 'Share', value: 'share' },
            { name: 'Delete', value: 'delete' }
          )
      );

    this.addSubcommand('export', 'Export document access logs')
      .addStringOption(option => 
        option.setName('user')
          .setDescription('Filter by user ID')
          .setRequired(false)
      )
      .addStringOption(option => 
        option.setName('file')
          .setDescription('Filter by file ID')
          .setRequired(false)
      )
      .addStringOption(option => 
        option.setName('type')
          .setDescription('Filter by access type')
          .setRequired(false)
          .addChoices(
            { name: 'View', value: 'view' },
            { name: 'Edit', value: 'edit' },
            { name: 'Download', value: 'download' },
            { name: 'Share', value: 'share' },
            { name: 'Delete', value: 'delete' }
          )
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
  async execute(interaction: CommandInteraction): Promise<void> {
    try {
      if (!this.auditService) {
        await interaction.reply({
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
          await interaction.reply({
            content: '❌ Unknown subcommand',
            ephemeral: true
          });
      }
    } catch (error) {
      logger.error('Error executing document audit command', {
        component: 'DocumentAuditCommand',
        error: error instanceof Error ? error.message : String(error)
      });
      
      await interaction.reply({
        content: '❌ An error occurred while processing your request',
        ephemeral: true
      });
    }
  }

  /**
   * Handle viewing access logs
   */
  private async handleViewLogs(interaction: CommandInteraction): Promise<void> {
    await interaction.deferReply();

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
        await interaction.editReply({
          content: '🔍 No document access logs found matching your criteria'
        });
        return;
      }

      // Format logs for display
      const logEntries = logs.map(log => {
        const timestamp = log.timestamp.toLocaleString();
        const status = log.success ? '✅' : '❌';
        const size = log.fileSize ? `${(log.fileSize / 1024).toFixed(1)}KB` : 'N/A';
        
        return `**${log.fileName}** (${log.fileId.substring(0, 8)}...)
${status} ${log.accessType} by ${log.userName} (${log.userId.substring(0, 8)}...)
📅 ${timestamp}
💾 ${size} | 📄 ${log.fileType || 'N/A'}`;
      });

      const response = `## 📋 Document Access Logs (${logs.length} entries)

${logEntries.join('\n\n')}`;

      await interaction.editReply({
        content: response
      });
    } catch (error) {
      logger.error('Error viewing access logs', {
        component: 'DocumentAuditCommand',
        error: error instanceof Error ? error.message : String(error)
      });
      
      await interaction.editReply({
        content: '❌ Failed to retrieve document access logs'
      });
    }
  }

  /**
   * Handle viewing access statistics
   */
  private async handleViewStats(interaction: CommandInteraction): Promise<void> {
    await interaction.deferReply();

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

      const response = `## 📊 Document Access Statistics

**Total Accesses:** ${stats.totalAccesses}
**Successful:** ${stats.successfulAccesses}
**Failed:** ${stats.failedAccesses}
**Unique Users:** ${stats.uniqueUsers}

### Access by Type:
${accessByType || 'No data'}

### Access by Time (Hour):
${accessByTime || 'No data'}`;

      await interaction.editReply({
        content: response
      });
    } catch (error) {
      logger.error('Error viewing access stats', {
        component: 'DocumentAuditCommand',
        error: error instanceof Error ? error.message : String(error)
      });
      
      await interaction.editReply({
        content: '❌ Failed to retrieve document access statistics'
      });
    }
  }

  /**
   * Handle exporting access logs
   */
  private async handleExportLogs(interaction: CommandInteraction): Promise<void> {
    await interaction.deferReply();

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
      await interaction.editReply({
        content: `✅ Document access logs exported successfully!
📋 ${exportedData.length} characters of data prepared for export.`
      });
    } catch (error) {
      logger.error('Error exporting access logs', {
        component: 'DocumentAuditCommand',
        error: error instanceof Error ? error.message : String(error)
      });
      
      await interaction.editReply({
        content: '❌ Failed to export document access logs'
      });
    }
  }
}