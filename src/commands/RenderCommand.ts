import type { SlashCommandStringOption } from 'discord.js';
import { BaseCommand, type CommandExecuteOptions } from './BaseCommand';
import type { BotConfig } from '@/types';
import { t } from '@/i18n';
import { MarkdownRenderService } from '@/services/MarkdownRenderService';
import { replyWithPrivacy } from '@/ui/reply';

export class RenderCommand extends BaseCommand {
  constructor(config: BotConfig) {
    super(
      'render',
      t('commands.render.description'),
      config,
      { 
        category: 'utility', 
        i18n: { 
          nameKey: 'commands.render.name', 
          descriptionKey: 'commands.render.description' 
        } 
      },
      (builder) => {
        builder
          .addStringOption((option: SlashCommandStringOption) =>
            option
              .setName('text')
              .setDescription(t('commands.render.opt.text.description'))
              .setRequired(true)
              .setMaxLength(2000)
          )
          .addStringOption((option: SlashCommandStringOption) =>
            option
              .setName('theme')
              .setDescription(t('commands.render.opt.theme.description'))
              .setRequired(false)
              .addChoices(
                { name: 'Discord Dark', value: 'dark' },
                { name: 'Light', value: 'light' },
                { name: 'Default', value: 'default' }
              )
          );
        return builder;
      }
    );
  }

  protected async onExecute(options: CommandExecuteOptions): Promise<void> {
    const { interaction } = options;
    
    try {
      await interaction.deferReply();

      const markdown = interaction.options.getString('text', true);
      const theme = interaction.options.getString('theme') || 'dark';
      
      // Access service through bot's service manager
      const bot = (interaction.client as any).bot;
      const markdownService = bot?.getService('markdownRender') as MarkdownRenderService | undefined;
      
      if (!markdownService) {
        throw new Error('Markdown render service is not available');
      }
      
      let attachment;
      switch (theme) {
        case 'light':
          attachment = await markdownService.renderLightTheme(markdown);
          break;
        case 'dark':
          attachment = await markdownService.renderDiscordDarkTheme(markdown);
          break;
        default:
          attachment = await markdownService.renderToImage(markdown);
      }
      
      await replyWithPrivacy(interaction, {
        content: t('commands.render.reply.success'),
        files: [attachment] 
      });
    } catch (error) {
      console.error('Render command failed', error);
      await replyWithPrivacy(interaction, { 
        content: t('commands.render.reply.error') 
      });
    }
  }
}