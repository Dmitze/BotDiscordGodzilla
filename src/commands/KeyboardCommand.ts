import type { ChatInputCommandInteraction, SlashCommandSubcommandBuilder } from 'discord.js';
import { BaseCommand, type CommandExecuteOptions } from '../BaseCommand';
import type { BotConfig } from '@/types';
import { t } from '@/i18n';
import { keyboardNavigationService } from '@/services/KeyboardNavigationService';
import { replyWithPrivacy } from '@/ui/reply';
import { UIHelper } from '@/utils/uiHelpers';

export class KeyboardCommand extends BaseCommand {
  constructor(config: BotConfig) {
    super(
      'keyboard',
      t('keyboard.command.description'),
      config,
      { category: 'settings', i18n: { nameKey: 'commands.keyboard.name', descriptionKey: 'keyboard.command.description' } },
      builder => {
        builder
          .addSubcommand(sub =>
            sub
              .setName('enable')
              .setDescription(t('keyboard.sub.enable.description'))
          )
          .addSubcommand(sub =>
            sub
              .setName('disable')
              .setDescription(t('keyboard.sub.disable.description'))
          )
          .addSubcommand(sub =>
            sub
              .setName('status')
              .setDescription(t('keyboard.sub.status.description'))
          )
          .addSubcommand(sub =>
            sub
              .setName('help')
              .setDescription(t('keyboard.sub.help.description'))
          )
          .addSubcommand(sub =>
            sub
              .setName('hints')
              .setDescription(t('keyboard.sub.hints.description'))
              .addBooleanOption(option =>
                option
                  .setName('show')
                  .setDescription(t('keyboard.opt.show.description'))
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
      case 'help':
        await this.handleHelp(interaction);
        break;
      case 'hints':
        await this.handleHints(interaction);
        break;
      default:
        await replyWithPrivacy(interaction, t('keyboard.common.unknownSub'));
    }
  }

  private async handleEnable(interaction: ChatInputCommandInteraction): Promise<void> {
    const userId = interaction.user.id;
    
    keyboardNavigationService.setEnabled(userId, true);
    
    await replyWithPrivacy(interaction, t('keyboard.enable.success'));
  }

  private async handleDisable(interaction: ChatInputCommandInteraction): Promise<void> {
    const userId = interaction.user.id;
    
    keyboardNavigationService.setEnabled(userId, false);
    
    await replyWithPrivacy(interaction, t('keyboard.disable.success'));
  }

  private async handleStatus(interaction: ChatInputCommandInteraction): Promise<void> {
    const userId = interaction.user.id;
    const prefs = keyboardNavigationService.getUserPreferences(userId);
    
    const embed = UIHelper.createBaseEmbed(
      t('keyboard.status.title'),
      t('keyboard.status.description'),
      prefs.enabled ? UIHelper.COLORS.SUCCESS : UIHelper.COLORS.WARNING
    );
    
    embed.addFields({
      name: t('keyboard.status.current'),
      value: prefs.enabled 
        ? `✅ ${t('keyboard.status.enabled')}` 
        : `❌ ${t('keyboard.status.disabled')}`,
      inline: false
    });
    
    embed.addFields({
      name: t('keyboard.status.hints'),
      value: prefs.showHints 
        ? `✅ ${t('keyboard.status.hintsShown')}` 
        : `❌ ${t('keyboard.status.hintsHidden')}`,
      inline: true
    });
    
    embed.addFields({
      name: t('keyboard.status.shortcuts'),
      value: t('keyboard.status.shortcutsCount', { count: prefs.shortcuts.length }),
      inline: true
    });
    
    await replyWithPrivacy(interaction, { embeds: [embed] });
  }

  private async handleHelp(interaction: ChatInputCommandInteraction): Promise<void> {
    const userId = interaction.user.id;
    const helpText = keyboardNavigationService.generateHelpText(userId);
    
    const embed = UIHelper.createBaseEmbed(
      t('keyboard.help.title'),
      helpText,
      UIHelper.COLORS.INFO
    );
    
    await replyWithPrivacy(interaction, { embeds: [embed] });
  }

  private async handleHints(interaction: ChatInputCommandInteraction): Promise<void> {
    const userId = interaction.user.id;
    const show = interaction.options.getBoolean('show', true);
    
    const prefs = keyboardNavigationService.getUserPreferences(userId);
    prefs.showHints = show;
    keyboardNavigationService.setUserPreferences(userId, prefs);
    
    const message = show 
      ? t('keyboard.hints.enabled') 
      : t('keyboard.hints.disabled');
    
    await replyWithPrivacy(interaction, message);
  }
}