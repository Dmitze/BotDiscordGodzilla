import { BaseCommand } from './BaseCommand';
import type { CommandExecuteOptions } from '@/types';
import { t } from '@/i18n';
import type { SlashCommandBuilder } from '@discordjs/builders';
import type { SlashCommandStringOption } from 'discord.js';
import logger from '@/utils/logger';

export class MarkdownCommand extends BaseCommand {
  constructor() {
    super(
      'markdown',
      t('commands.markdown.description'),
      {} as any, // config
      {
        i18n: {
          nameKey: 'commands.markdown.name',
          descriptionKey: 'commands.markdown.description',
        },
      },
      (builder: SlashCommandBuilder): SlashCommandBuilder => {
        builder.addStringOption((option: SlashCommandStringOption) =>
          option
            .setName('content')
            .setDescription(t('commands.markdown.options.content.description'))
            .setRequired(true)
            .setMaxLength(1000)
        );
        builder.addStringOption((option: SlashCommandStringOption) =>
          option
            .setName('format')
            .setDescription(t('commands.markdown.options.format.description'))
            .setRequired(false)
            .addChoices(
              { name: 'Text', value: 'text' },
              { name: 'Image', value: 'image' }
            )
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
      
      // Get the markdown content and format from the interaction
      const content = interaction.options.getString('content', true);
      const format = interaction.options.getString('format') || 'text';
      
      // Get the markdown rendering service
      const markdownService = interaction.client.serviceContainer?.get('markdownRendering');
      
      if (!markdownService) {
        await interaction.editReply({
          content: t('commands.markdown.errors.service_unavailable'),
        });
        return;
      }
      
      if (format === 'image') {
        // Render as image
        const attachment = await markdownService.renderToImage(content);
        await interaction.editReply({
          files: [attachment],
        });
      } else {
        // Render as text
        const renderedText = await markdownService.renderToText(content);
        await interaction.editReply({
          content: renderedText,
        });
      }
    } catch (error) {
      logger.error('Error in markdown command', { error });
      await interaction.editReply({
        content: t('commands.markdown.errors.rendering_failed'),
      });
    }
  }
}