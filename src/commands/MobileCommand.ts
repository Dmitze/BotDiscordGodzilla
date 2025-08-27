import type { ChatInputCommandInteraction, SlashCommandSubcommandBuilder } from 'discord.js';
import { BaseCommand, type CommandExecuteOptions } from '../BaseCommand';
import type { BotConfig } from '@/types';
import { t } from '@/i18n';
import { mobileOptimizationService } from '@/services/MobileOptimizationService';
import { replyWithPrivacy } from '@/ui/reply';
import { UIHelper } from '@/utils/uiHelpers';

export class MobileCommand extends BaseCommand {
  constructor(config: BotConfig) {
    super(
      'mobile',
      t('mobile.command.description'),
      config,
      { category: 'settings', i18n: { nameKey: 'commands.mobile.name', descriptionKey: 'mobile.command.description' } },
      builder => {
        builder
          .addSubcommand(sub =>
            sub
              .setName('enable')
              .setDescription(t('mobile.sub.enable.description'))
          )
          .addSubcommand(sub =>
            sub
              .setName('disable')
              .setDescription(t('mobile.sub.disable.description'))
          )
          .addSubcommand(sub =>
            sub
              .setName('status')
              .setDescription(t('mobile.sub.status.description'))
          )
          .addSubcommand(sub =>
            sub
              .setName('compact')
              .setDescription(t('mobile.sub.compact.description'))
              .addBooleanOption(option =>
                option
                  .setName('mode')
                  .setDescription(t('mobile.opt.mode.description'))
                  .setRequired(true)
              )
          )
          .addSubcommand(sub =>
            sub
              .setName('contrast')
              .setDescription(t('mobile.sub.contrast.description'))
              .addBooleanOption(option =>
                option
                  .setName('mode')
                  .setDescription(t('mobile.opt.contrast.description'))
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
      case 'enable':
        await this.handleEnable(interaction);
        break;
      case 'disable':
        await this.handleDisable(interaction);
        break;
      case 'status':
        await this.handleStatus(interaction);
        break;
      case 'compact':
        await this.handleCompact(interaction);
        break;
      case 'contrast':
        await this.handleContrast(interaction);
        break;
      default:
        await replyWithPrivacy(interaction, t('mobile.common.unknownSub'));
    }
  }

  private async handleEnable(interaction: ChatInputCommandInteraction): Promise<void> {
    const userId = interaction.user.id;
    
    mobileOptimizationService.setEnabled(userId, true);
    
    await replyWithPrivacy(interaction, t('mobile.enable.success'));
  }

  private async handleDisable(interaction: ChatInputCommandInteraction): Promise<void> {
    const userId = interaction.user.id;
    
    mobileOptimizationService.setEnabled(userId, false);
    
    await replyWithPrivacy(interaction, t('mobile.disable.success'));
  }

  private async handleStatus(interaction: ChatInputCommandInteraction): Promise<void> {
    const userId = interaction.user.id;
    const status = mobileOptimizationService.getStatus(userId);
    const prefs = mobileOptimizationService.getUserPreferences(userId);
    
    const embed = UIHelper.createBaseEmbed(
      t('mobile.status.title'),
      t('mobile.status.description'),
      status.enabled ? UIHelper.COLORS.SUCCESS : UIHelper.COLORS.WARNING
    );
    
    embed.addFields({
      name: t('mobile.status.current'),
      value: status.enabled 
        ? `✅ ${t('mobile.status.enabled')}` 
        : `❌ ${t('mobile.status.disabled')}`,
      inline: false
    });
    
    if (status.enabled) {
      embed.addFields({
        name: t('mobile.status.mode'),
        value: prefs.compactMode 
          ? `📱 ${t('mobile.status.compact')}` 
          : `📄 ${t('mobile.status.normal')}`,
        inline: true
      });
      
      embed.addFields({
        name: t('mobile.status.contrast'),
        value: prefs.contrastMode 
          ? `⚫ ${t('mobile.status.highContrast')}` 
          : `⚪ ${t('mobile.status.normalContrast')}`,
        inline: true
      });
      
      embed.addFields({
        name: t('mobile.status.limits'),
        value: `${t('mobile.status.componentsPerRow')}: ${prefs.maxComponentsPerRow}\n${t('mobile.status.actionRows')}: ${prefs.maxActionRows}`,
        inline: false
      });
    }
    
    await replyWithPrivacy(interaction, { embeds: [embed] });
  }

  private async handleCompact(interaction: ChatInputCommandInteraction): Promise<void> {
    const userId = interaction.user.id;
    const mode = interaction.options.getBoolean('mode', true);
    
    const prefs = mobileOptimizationService.getUserPreferences(userId);
    prefs.compactMode = mode;
    mobileOptimizationService.setUserPreferences(userId, prefs);
    
    const message = mode 
      ? t('mobile.compact.enabled') 
      : t('mobile.compact.disabled');
    
    await replyWithPrivacy(interaction, message);
  }

  private async handleContrast(interaction: ChatInputCommandInteraction): Promise<void> {
    const userId = interaction.user.id;
    const mode = interaction.options.getBoolean('mode', true);
    
    const prefs = mobileOptimizationService.getUserPreferences(userId);
    prefs.contrastMode = mode;
    mobileOptimizationService.setUserPreferences(userId, prefs);
    
    const message = mode 
      ? t('mobile.contrast.high') 
      : t('mobile.contrast.normal');
    
    await replyWithPrivacy(interaction, message);
  }
}