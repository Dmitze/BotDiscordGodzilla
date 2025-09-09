import { BaseCommand } from './BaseCommand';
import type { CommandExecuteOptions } from '@/types';
import { t } from '@/i18n';
import type { SlashCommandBuilder } from '@discordjs/builders';
import type { SlashCommandStringOption, SlashCommandBooleanOption } from 'discord.js';
import logger from '@/utils/logger';

export class OllamaCommand extends BaseCommand {
  constructor() {
    super(
      'ollama',
      t('commands.ollama.description'),
      {} as any, // config
      {
        i18n: {
          nameKey: 'commands.ollama.name',
          descriptionKey: 'commands.ollama.description',
        },
      },
      (builder: SlashCommandBuilder): SlashCommandBuilder => {
        builder.addStringOption((option: SlashCommandStringOption) =>
          option
            .setName('prompt')
            .setDescription(t('commands.ollama.options.prompt.description'))
            .setRequired(true)
            .setMaxLength(2000)
        );
        builder.addStringOption((option: SlashCommandStringOption) =>
          option
            .setName('model')
            .setDescription(t('commands.ollama.options.model.description'))
            .setRequired(false)
        );
        builder.addBooleanOption((option: SlashCommandBooleanOption) =>
          option
            .setName('reset')
            .setDescription(t('commands.ollama.options.reset.description'))
            .setRequired(false)
        );
        return builder;
      }
    );
  }

  protected async onExecute(options: CommandExecuteOptions): Promise<void> {
    const { interaction } = options;
    
    try {
      // Defer the reply to allow time for processing
      await interaction.deferReply();
      
      // Get the prompt, model, and reset flag from the interaction
      const prompt = interaction.options.getString('prompt', true);
      const model = interaction.options.getString('model') || undefined;
      const reset = interaction.options.getBoolean('reset') || false;
      
      // Get the Ollama service
      const ollamaService = interaction.client.serviceContainer?.get('ollama');
      
      if (!ollamaService) {
        await interaction.editReply({
          content: t('commands.ollama.errors.service_unavailable'),
        });
        return;
      }
      
      // Handle reset command
      if (reset) {
        await ollamaService.resetChannelHistory(interaction.channelId);
        await interaction.editReply({
          content: t('commands.ollama.messages.history_reset'),
        });
        return;
      }
      
      // Generate response from Ollama
      const response = await ollamaService.generate(prompt, {
        model,
        channelId: interaction.channelId,
      });
      
      // Truncate response if too long for Discord (2000 character limit)
      let responseText = response;
      if (responseText.length > 1900) {
        responseText = responseText.substring(0, 1900) + '\n\n... *(response truncated)*';
      }
      
      // Send the response
      await interaction.editReply({
        content: responseText,
      });
    } catch (error) {
      logger.error('Error in Ollama command', { error });
      await interaction.editReply({
        content: t('commands.ollama.errors.generation_failed'),
      });
    }
  }
}