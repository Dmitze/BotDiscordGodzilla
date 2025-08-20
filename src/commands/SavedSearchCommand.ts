import type { BotConfig, CommandExecuteOptions } from '@/types';
import { BaseCommand } from './BaseCommand';
import { t } from '@/i18n';
import logger from '@/utils/logger';
import type { GoogleService } from '@/services/GoogleService';
import type { SearchIndex, SearchQuery } from '@/search/SearchIndex';
import { defaultWorkspaceService, WorkspaceService } from '@/services/WorkspaceService';
import type { DriveListQuery } from '@/types/drive';

function getWorkspace(interaction: any): WorkspaceService {
  const svc = (interaction.client as any)?.serviceContainer?.get?.('workspace');
  return (svc as WorkspaceService) || defaultWorkspaceService;
}

function getGoogle(interaction: any): GoogleService | undefined {
  return ((interaction.client as any)?.serviceContainer?.get?.('google') as GoogleService) || undefined;
}

function getSearchIndex(interaction: any): SearchIndex | undefined {
  return ((interaction.client as any)?.serviceContainer?.get?.('searchIndex') as SearchIndex) || undefined;
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
    // Попробуем выполнить через персистентный индекс (SQLite), если доступен
    const searchIndex = getSearchIndex(interaction);
    if (searchIndex) {
      const saved = ws.getSavedSearch(userId, name);
      if (!saved) {
        await interaction.reply({ content: t('workspace.search.runNotFound'), ephemeral: true });
        return;
      }
      // Применяем политики из конфига
      const cfg = this.config.drive || {};
      const f = { ...(saved.filters || {}) } as any;
      // Политики allowedMime/ownerAllowlist
      if (Array.isArray(cfg.allowedMime) && cfg.allowedMime.length) {
        f.mimeIncludes = Array.isArray(f.mimeIncludes) && f.mimeIncludes.length
          ? f.mimeIncludes.filter((m: string) => cfg.allowedMime!.includes(m))
          : cfg.allowedMime;
      }
      if (Array.isArray(cfg.ownerAllowlist) && cfg.ownerAllowlist.length) {
        f.ownerAllowlist = cfg.ownerAllowlist;
      }
      // Маппинг DriveListQuery -> SearchQuery
      const limit = Math.max(1, Math.min(25, f.pageSize ?? (this.config.drive?.pageSize ?? 10)));
      const filterObj: any = {};
      const mime = Array.isArray(f.mimeIncludes) && f.mimeIncludes.length ? f.mimeIncludes : null;
      const owner = Array.isArray(f.ownerAllowlist) && f.ownerAllowlist.length ? f.ownerAllowlist : null;
      const modifiedFrom = f.dateFrom ? new Date(f.dateFrom).getTime() : null;
      const modifiedTo = f.dateTo ? new Date(f.dateTo).getTime() : null;
      const sizeFrom = typeof f.sizeMin === 'number' ? Math.max(0, f.sizeMin) * 1024 * 1024 : null;
      const sizeTo = typeof f.sizeMax === 'number' ? Math.max(0, f.sizeMax) * 1024 * 1024 : null;
      const tags = Array.isArray(f.tags) && f.tags.length ? f.tags : null;
      if (mime) filterObj.mime = mime;
      if (owner) filterObj.owner = owner;
      if (typeof modifiedFrom === 'number' && !Number.isNaN(modifiedFrom)) filterObj.modifiedFrom = modifiedFrom;
      if (typeof modifiedTo === 'number' && !Number.isNaN(modifiedTo)) filterObj.modifiedTo = modifiedTo;
      if (typeof sizeFrom === 'number') filterObj.sizeFrom = sizeFrom;
      if (typeof sizeTo === 'number') filterObj.sizeTo = sizeTo;
      if (tags) filterObj.tags = tags;

      const q: SearchQuery = {
        text: (f.query || '').toString(),
        limit,
        ...(Object.keys(filterObj).length ? { filters: filterObj } as any : {}),
      };
      let hits: Awaited<ReturnType<SearchIndex['search']>>;
      try {
        hits = await searchIndex.search(q);
      } catch (e) {
        // Фоллбек на Google, если индекс недоступен
        const google = getGoogle(interaction);
        if (!google) {
          await interaction.reply({ content: t('files.error.serviceUnavailable'), ephemeral: true });
          return;
        }
        const result = await ws.runSearch(userId, name, { google, config: this.config });
        const items = Array.isArray((result as any)?.files) ? (result as any).files : [];
        const pageSize = this.config.drive?.pageSize ?? 10;
        const lines = items.slice(0, pageSize).map((f: any) => `• ${f.name || f.id}`);
        await interaction.reply({ content: `${t('workspace.search.listTitle')}` + "\n" + lines.join('\n'), ephemeral: true });
        return;
      }
      const items = hits?.hits || [];
      if (!items.length) {
        await interaction.reply({ content: t('files.result.searchEmpty', { query: '' }), ephemeral: true });
        return;
      }
      const lines = items.map((h: any) => `• ${h.name || h.fileId}`);
      await interaction.reply({ content: `${t('workspace.search.listTitle')}` + "\n" + lines.join('\n'), ephemeral: true });
      return;
    }

    // Если индекса нет — прежнее поведение через Google
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
    await interaction.reply({ content: `${t('workspace.search.listTitle')}` + "\n" + lines.join('\n'), ephemeral: true });
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
