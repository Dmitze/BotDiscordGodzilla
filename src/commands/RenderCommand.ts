import type { BotConfig } from '@/types';
import type { CommandExecuteOptions } from '@/types/commands';
import { BaseCommand } from './BaseCommand';
import type { ServiceContainer } from '../core/ServiceContainer';
import { MarkdownRenderService } from '../services/MarkdownRenderService';
import { i18n } from '../i18n';
import type { 
  SlashCommandBuilder, 
  SlashCommandStringOption, 
  CommandInteraction 
} from 'discord.js';

export class RenderCommand extends BaseCommand {
  constructor(config: BotConfig) {
    super('render', i18n.__('commands.render.description'), config, {
      i18n: { nameKey: 'commands.render.name', descriptionKey: 'commands.render.description' }
    }, (builder: SlashCommandBuilder): SlashCommandBuilder => {
      builder.addStringOption((option: SlashCommandStringOption) =>
        option
          .setName('markdown')
          .setDescription(i18n.__('commands.render.opt.markdown.description'))
          .setRequired(true)
          .setMaxLength(2000)
      );
      builder.addStringOption((option: SlashCommandStringOption) =>
        option
          .setName('theme')
          .setDescription(i18n.__('commands.render.opt.theme.description'))
          .setRequired(false)
          .addChoices(
            { name: 'Discord Dark', value: 'dark' },
            { name: 'Light', value: 'light' },
            { name: 'Default', value: 'default' }
          )
      );
      return builder;
    });
  }

  async onExecute(options: CommandExecuteOptions): Promise<void> {
    const { interaction } = options;
    
    try {
      await interaction.deferReply();

      const markdown = interaction.options.get('markdown')?.value as string;
      const theme = interaction.options.get('theme')?.value as string || 'dark';
      
      const markdownService = this.container.get('markdownRender') as MarkdownRenderService;
      
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
      
      await interaction.editReply({ 
        content: i18n.__('commands.render.reply.success'),
        files: [attachment] 
      });
    } catch (error) {
      this.logger.error('Render command failed', error);
      await interaction.editReply({ 
        content: i18n.__('commands.render.reply.error') 
      });
    }
  }
}