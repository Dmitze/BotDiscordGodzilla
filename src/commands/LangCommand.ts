import { SlashCommandStringOption } from 'discord.js';
import { BaseCommand } from './BaseCommand';
import type { BotConfig } from '@/types';
import type { CommandExecuteOptions } from './BaseCommand';
import { t } from '@/i18n';
import { UserPreferencesService, type SupportedLocale } from '@/services/UserPreferencesService';

function toSupportedLocale(input: string): SupportedLocale {
  const lc = input.toLowerCase();
  if (lc === 'uk' || lc === 'uk-ua') return 'uk';
  if (lc === 'en' || lc === 'en-us') return 'en-US';
  // default
  return 'uk';
}

export class LangCommand extends BaseCommand {
  constructor(config: BotConfig) {
    super(
      'lang',
      t('lang.command.description'),
      config,
      { category: 'settings' },
      builder => {
        builder
          .addSubcommand(sub =>
            sub
              .setName('set')
              .setDescription(t('lang.sub.set.description'))
              .addStringOption((opt: SlashCommandStringOption) =>
                opt
                  .setName('locale')
                  .setDescription(t('lang.opt.locale.description'))
                  .setRequired(true)
                  .addChoices(
                    { name: 'Українська', value: 'uk' },
                    { name: 'English', value: 'en' },
                  )
              )
          )
          .addSubcommand(sub =>
            sub
              .setName('show')
              .setDescription(t('lang.sub.show.description'))
          );
        return builder;
      }
    );
  }

  protected async onExecute(options: CommandExecuteOptions): Promise<void> {
    const { interaction } = options;

    const sub = interaction.options.getSubcommand();

    if (sub === 'set') {
      const userId = interaction.user.id;
      const localeInput = interaction.options.getString('locale', true);
      const normalized = toSupportedLocale(localeInput);
      UserPreferencesService.setLocale(userId, normalized);
      await interaction.reply({
        content: t('lang.reply.setOk', { locale: normalized }),
        ephemeral: true,
      });
      return;
    }

    if (sub === 'show') {
      const userId = interaction.user.id;
      const current = UserPreferencesService.getLocale(userId);
      await interaction.reply({
        content: t('lang.reply.current', { locale: current }),
        ephemeral: true,
      });
      return;
    }

    await interaction.reply({ content: t('lang.reply.unknownSub'), ephemeral: true });
  }
}
