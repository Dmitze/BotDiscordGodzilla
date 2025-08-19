import type { BotConfig, CommandExecuteOptions } from '@/types';
import { BaseCommand } from './BaseCommand';
import { t } from '@/i18n';
import logger from '@/utils/logger';
import type { GoogleService } from '@/services/GoogleService';
import { defaultWorkspaceService, WorkspaceService } from '@/services/WorkspaceService';
import type { DriveListQuery } from '@/types/drive';

function getWorkspace(interaction: any): WorkspaceService {
  const svc = (interaction.client as any)?.serviceContainer?.get?.('workspace');
  return (svc as WorkspaceService) || defaultWorkspaceService;
}

function getGoogle(interaction: any): GoogleService | undefined {
  return ((interaction.client as any)?.serviceContainer?.get?.('google') as GoogleService) || undefined;
}

export class SavedSearchCommand extends BaseCommand {
  constructor(config: BotConfig) {
    super('search', t('workspace.search.command.description'), config, {}, (builder: any) => {
      return builder
        .addSubcommand((sub: any) =>
          sub
            .setName('save')
            .setDescription(t('workspace.search.sub.save'))
            .addStringOption((opt: any) => opt.setName('name').setDescription(t('workspace.search.opt.name')).setRequired(true))
            .addStringOption((opt: any) => opt.setName('query').setDescription(t('workspace.search.opt.query')).setRequired(false))
            .addStringOption((opt: any) => opt.setName('folder').setDescription(t('workspace.search.opt.folder')).setRequired(false))
            .addStringOption((opt: any) => opt.setName('mime').setDescription(t('workspace.search.opt.mime')).setRequired(false))
            .addStringOption((opt: any) => opt.setName('owner').setDescription(t('workspace.search.opt.owner')).setRequired(false))
            .addStringOption((opt: any) => opt.setName('date_from').setDescription('YYYY-MM-DD').setRequired(false))
            .addStringOption((opt: any) => opt.setName('date_to').setDescription('YYYY-MM-DD').setRequired(false))
            .addIntegerOption((opt: any) => opt.setName('size_min').setDescription('MB').setRequired(false).setMinValue(0))
            .addIntegerOption((opt: any) => opt.setName('size_max').setDescription('MB').setRequired(false).setMinValue(0))
            .addIntegerOption((opt: any) => opt.setName('limit').setDescription('page size').setRequired(false).setMinValue(1).setMaxValue(25))
            .addStringOption((opt: any) =>
              opt
                .setName('sort_by')
                .setDescription('name | modifiedTime')
                .setRequired(false)
                .addChoices({ name: 'name', value: 'name' }, { name: 'modifiedTime', value: 'modifiedTime' })
            )
            .addStringOption((opt: any) => opt.setName('sort_dir').setDescription('asc|desc').setRequired(false))
        )
        .addSubcommand((sub: any) =>
          sub
            .setName('run')
            .setDescription(t('workspace.search.sub.run'))
            .addStringOption((opt: any) => opt.setName('name').setDescription(t('workspace.search.opt.name')).setRequired(true))
        )
        .addSubcommand((sub: any) =>
          sub
            .setName('list')
            .setDescription(t('workspace.search.sub.list'))
        )
        .addSubcommand((sub: any) =>
          sub
            .setName('remove')
            .setDescription(t('workspace.search.sub.remove'))
            .addStringOption((opt: any) => opt.setName('name').setDescription(t('workspace.search.opt.name')).setRequired(true))
        );
    });
  }

  protected async onExecute(options: CommandExecuteOptions): Promise<void> {
    const { interaction } = options as any;
    try {
      const sub = interaction.options.getSubcommand();
      if (sub === 'save') return this.handleSave(interaction);
      if (sub === 'run') return this.handleRun(interaction);
      if (sub === 'list') return this.handleList(interaction);
      if (sub === 'remove') return this.handleRemove(interaction);
      await interaction.reply({ content: t('workspace.common.unknownSub'), ephemeral: true });
    } catch (error) {
      logger.error('SavedSearchCommand failed', { error: String(error) });
      if (interaction.deferred || interaction.replied) {
        await interaction.editReply({ content: t('workspace.common.execError') });
      } else {
        await interaction.reply({ content: t('workspace.common.execError'), ephemeral: true });
      }
    }
  }

  private extractFilters(interaction: any): DriveListQuery {
    const q: DriveListQuery = {
      folderId: interaction.options.getString('folder') || undefined,
      query: interaction.options.getString('query') || undefined,
      mimeIncludes: interaction.options.getString('mime') ? [interaction.options.getString('mime')] : [],
      ownerAllowlist: interaction.options.getString('owner') ? [interaction.options.getString('owner')] : [],
      dateFrom: interaction.options.getString('date_from') || undefined,
      dateTo: interaction.options.getString('date_to') || undefined,
      sizeMin: interaction.options.getInteger('size_min') || undefined,
      sizeMax: interaction.options.getInteger('size_max') || undefined,
      pageSize: interaction.options.getInteger('limit') || undefined,
      sortBy: interaction.options.getString('sort_by') || undefined,
      sortDir: interaction.options.getString('sort_dir') || undefined,
    } as any;
    return q;
  }

  private async handleSave(interaction: any): Promise<void> {
    const userId = interaction.user.id;
    const name = interaction.options.getString('name', true);
    const ws = getWorkspace(interaction);
    const filters = this.extractFilters(interaction);
    const res = await ws.saveSearch(userId, name, filters);
    await interaction.reply({ content: t('workspace.search.saved', { name: res.search.name }), ephemeral: true });
  }

  private async handleRun(interaction: any): Promise<void> {
    const userId = interaction.user.id;
    const name = interaction.options.getString('name', true);
    const ws = getWorkspace(interaction);
    const google = getGoogle(interaction);
    if (!google) {
      await interaction.reply({ content: t('files.error.serviceUnavailable'), ephemeral: true });
      return;
    }
    const result = await ws.runSearch(userId, name, { google, config: this.config });
    if (!result) {
      await interaction.reply({ content: t('workspace.search.runNotFound'), ephemeral: true });
      return;
    }
    const items = Array.isArray((result as any).files) ? (result as any).files : [];
    if (!items.length) {
      await interaction.reply({ content: t('files.result.searchEmpty', { query: '' }), ephemeral: true });
      return;
    }
    const pageSize = this.config.drive?.pageSize ?? 10;
    const lines = items.slice(0, pageSize).map((f: any) => `• ${f.name || f.id}`);
    await interaction.reply({ content: `${t('workspace.search.listTitle')}\n${lines.join('\n')}`, ephemeral: true });
  }

  private async handleList(interaction: any): Promise<void> {
    const userId = interaction.user.id;
    const ws = getWorkspace(interaction);
    const list = await ws.listSearches(userId);
    if (!list.length) {
      await interaction.reply({ content: t('workspace.search.listEmpty'), ephemeral: true });
      return;
    }
    const lines = list.map(s => `• ${s.name}`);
    await interaction.reply({ content: `${t('workspace.search.listTitle')}\n${lines.join('\n')}`, ephemeral: true });
  }

  private async handleRemove(interaction: any): Promise<void> {
    const userId = interaction.user.id;
    const name = interaction.options.getString('name', true);
    const ws = getWorkspace(interaction);
    const ok = await ws.removeSearch(userId, name);
    await interaction.reply({ content: ok ? t('workspace.search.removed') : t('workspace.common.notFound'), ephemeral: true });
  }
}
