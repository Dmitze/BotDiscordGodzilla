import type { ChatInputCommandInteraction, SlashCommandSubcommandBuilder } from 'discord.js';
import { BaseCommand, type CommandExecuteOptions } from '../BaseCommand';
import type { BotConfig } from '@/types';
import { t } from '@/i18n';
import { dashboardViewService } from '@/services/DashboardViewService';
import { replyWithPrivacy } from '@/ui/reply';
import { UIHelper } from '@/utils/uiHelpers';

export class DashboardCommand extends BaseCommand {
  constructor(config: BotConfig) {
    super(
      'dashboard',
      t('dashboard.command.description'),
      config,
      { category: 'files', i18n: { nameKey: 'commands.dashboard.name', descriptionKey: 'dashboard.command.description' } },
      builder => {
        builder
          .addSubcommand(sub =>
            sub
              .setName('view')
              .setDescription(t('dashboard.sub.view.description'))
              .addStringOption(option =>
                option
                  .setName('name')
                  .setDescription(t('dashboard.opt.name.description'))
                  .setRequired(false)
              )
          )
          .addSubcommand(sub =>
            sub
              .setName('list')
              .setDescription(t('dashboard.sub.list.description'))
          )
          .addSubcommand(sub =>
            sub
              .setName('create')
              .setDescription(t('dashboard.sub.create.description'))
              .addStringOption(option =>
                option
                  .setName('name')
                  .setDescription(t('dashboard.opt.name.description'))
                  .setRequired(true)
              )
              .addStringOption(option =>
                option
                  .setName('layout')
                  .setDescription(t('dashboard.opt.layout.description'))
                  .setRequired(false)
                  .addChoices(
                    { name: 'List', value: 'list' },
                    { name: 'Grid', value: 'grid' },
                    { name: 'Compact', value: 'compact' }
                  )
              )
              .addStringOption(option =>
                option
                  .setName('sort_by')
                  .setDescription(t('dashboard.opt.sortBy.description'))
                  .setRequired(false)
                  .addChoices(
                    { name: 'Name', value: 'name' },
                    { name: 'Modified Time', value: 'modifiedTime' },
                    { name: 'Size', value: 'size' },
                    { name: 'Type', value: 'mimeType' }
                  )
              )
              .addStringOption(option =>
                option
                  .setName('sort_order')
                  .setDescription(t('dashboard.opt.sortOrder.description'))
                  .setRequired(false)
                  .addChoices(
                    { name: 'Ascending', value: 'asc' },
                    { name: 'Descending', value: 'desc' }
                  )
              )
              .addIntegerOption(option =>
                option
                  .setName('items_per_page')
                  .setDescription(t('dashboard.opt.itemsPerPage.description'))
                  .setRequired(false)
                  .setMinValue(5)
                  .setMaxValue(100)
              )
          )
          .addSubcommand(sub =>
            sub
              .setName('delete')
              .setDescription(t('dashboard.sub.delete.description'))
              .addStringOption(option =>
                option
                  .setName('name')
                  .setDescription(t('dashboard.opt.name.description'))
                  .setRequired(true)
              )
          );
        return builder;
      }
    );
  }

  protected async onExecute(options: CommandExecuteOptions): Promise<void> {
    const { interaction } = options;
    const subcommand = interaction.options.getSubcommand();

    switch (subcommand) {
      case 'view':
        await this.handleView(interaction);
        break;
      case 'list':
        await this.handleList(interaction);
        break;
      case 'create':
        await this.handleCreate(interaction);
        break;
      case 'delete':
        await this.handleDelete(interaction);
        break;
      default:
        await replyWithPrivacy(interaction, t('dashboard.common.unknownSub'));
    }
  }

  private async handleView(interaction: ChatInputCommandInteraction): Promise<void> {
    const userId = interaction.user.id;
    const viewName = interaction.options.getString('name') || 'default';
    
    const viewConfig = dashboardViewService.getViewConfig(userId, viewName);
    
    if (!viewConfig) {
      await replyWithPrivacy(interaction, t('dashboard.view.notFound', { name: viewName }));
      return;
    }
    
    const embed = UIHelper.createBaseEmbed(
      t('dashboard.view.title', { name: viewConfig.viewName }),
      t('dashboard.view.description'),
      UIHelper.COLORS.INFO
    );
    
    // Add view details
    embed.addFields(
      {
        name: t('dashboard.view.layout'),
        value: viewConfig.layout,
        inline: true
      },
      {
        name: t('dashboard.view.sortBy'),
        value: `${viewConfig.sortBy} (${viewConfig.sortOrder})`,
        inline: true
      },
      {
        name: t('dashboard.view.itemsPerPage'),
        value: viewConfig.itemsPerPage.toString(),
        inline: true
      }
    );
    
    // Add display options
    const displayOptions = [
      viewConfig.showPreview ? t('dashboard.view.showPreview') : '',
      viewConfig.showTags ? t('dashboard.view.showTags') : '',
      viewConfig.showOwner ? t('dashboard.view.showOwner') : '',
      viewConfig.showDates ? t('dashboard.view.showDates') : ''
    ].filter(Boolean).join(', ') || t('dashboard.view.none');
    
    embed.addFields({
      name: t('dashboard.view.displayOptions'),
      value: displayOptions,
      inline: false
    });
    
    await replyWithPrivacy(interaction, { embeds: [embed] });
  }

  private async handleList(interaction: ChatInputCommandInteraction): Promise<void> {
    const userId = interaction.user.id;
    const prefs = dashboardViewService.getUserPreferences(userId);
    
    if (prefs.views.length === 0) {
      await replyWithPrivacy(interaction, t('dashboard.list.empty'));
      return;
    }
    
    const embed = UIHelper.createBaseEmbed(
      t('dashboard.list.title'),
      t('dashboard.list.description'),
      UIHelper.COLORS.INFO
    );
    
    // Add views list
    const viewList = prefs.views.map((view, index) => {
      const isDefault = view.viewName === prefs.defaultView ? ` (${t('dashboard.list.default')})` : '';
      return `${index + 1}. **${view.viewName}**${isDefault} - ${view.layout} (${view.itemsPerPage} items/page)`;
    }).join('\n');
    
    embed.setDescription(viewList);
    
    await replyWithPrivacy(interaction, { embeds: [embed] });
  }

  private async handleCreate(interaction: ChatInputCommandInteraction): Promise<void> {
    const userId = interaction.user.id;
    const viewName = interaction.options.getString('name', true);
    const layout = interaction.options.getString('layout') as 'list' | 'grid' | 'compact' | null;
    const sortBy = interaction.options.getString('sort_by') as 'name' | 'modifiedTime' | 'size' | 'mimeType' | null;
    const sortOrder = interaction.options.getString('sort_order') as 'asc' | 'desc' | null;
    const itemsPerPage = interaction.options.getInteger('items_per_page');
    
    // Get existing preferences or create new ones
    const prefs = dashboardViewService.getUserPreferences(userId);
    
    // Create new view config
    const newViewConfig = {
      userId,
      viewName,
      layout: layout || 'list',
      sortBy: sortBy || 'modifiedTime',
      sortOrder: sortOrder || 'desc',
      showPreview: true,
      showTags: true,
      showOwner: true,
      showDates: true,
      itemsPerPage: itemsPerPage || 25,
      fileFilters: {}
    };
    
    // Save the new view
    dashboardViewService.saveViewConfig(userId, newViewConfig);
    
    await replyWithPrivacy(interaction, t('dashboard.create.success', { name: viewName }));
  }

  private async handleDelete(interaction: ChatInputCommandInteraction): Promise<void> {
    const userId = interaction.user.id;
    const viewName = interaction.options.getString('name', true);
    
    const result = dashboardViewService.deleteViewConfig(userId, viewName);
    
    if (result) {
      await replyWithPrivacy(interaction, t('dashboard.delete.success', { name: viewName }));
    } else {
      await replyWithPrivacy(interaction, t('dashboard.delete.failed', { name: viewName }));
    }
  }
}