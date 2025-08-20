import type { SlashCommandStringOption } from 'discord.js';
import { BaseCommand, type CommandExecuteOptions } from './BaseCommand';
import type { BotConfig } from '@/types';
import { t } from '@/i18n';
import { UserPreferencesService, type SupportedLocale } from '@/services/UserPreferencesService';
import { replyWithPrivacy } from '@/ui/reply';

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
      await replyWithPrivacy(interaction, t('lang.reply.setOk', { locale: normalized }));
      return;
    }

    if (sub === 'show') {
      const userId = interaction.user.id;
      const current = UserPreferencesService.getLocale(userId);
      await replyWithPrivacy(interaction, t('lang.reply.current', { locale: current }));
      return;
    }

    await replyWithPrivacy(interaction, t('lang.reply.unknownSub'));
  }
}
