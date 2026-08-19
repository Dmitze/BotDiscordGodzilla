import { SlashCommandBuilder } from 'discord.js';
import { BaseCommand } from './BaseCommand';
import type { BotConfig, CommandExecuteOptions } from '@/types';
import { StandupService } from '@/services/StandupService';

export default class StandupCommand extends BaseCommand {
  private standupService: StandupService | null = null;

  constructor(config: BotConfig) {
    super('standup', 'Daily Standup Assistant', config, {
      category: 'business'
    });
  }

  protected buildCommand(): Omit<SlashCommandBuilder, 'addSubcommand' | 'addSubcommandGroup'> {
    const builder = new SlashCommandBuilder()
      .setName(this.name)
      .setDescription(this.description)
      .addSubcommand(subcommand =>
        subcommand
          .setName('trigger')
          .setDescription('Запустити Daily Standup в цьому каналі')
      ) as SlashCommandBuilder;
      
    return builder;
  }

  protected override async onExecute(options: CommandExecuteOptions): Promise<void> {
    const { interaction } = options;
    const subcommand = interaction.options.getSubcommand();
    
    // Лінива ініціалізація сервісу (lazy loading)
    if (!this.standupService) {
      const anyClient = interaction.client as any;
      if (anyClient.standupService) {
         this.standupService = anyClient.standupService;
      } else {
         this.standupService = new StandupService(interaction.client);
         anyClient.standupService = this.standupService;
      }
    }

    if (subcommand === 'trigger') {
      await interaction.deferReply({ ephemeral: true });
      await this.standupService?.triggerStandup(interaction.channelId);
      await interaction.editReply('✅ Daily Standup успішно запущено в цьому каналі!');
    }
  }
}
