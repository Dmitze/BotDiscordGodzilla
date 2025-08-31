import { AttachmentBuilder } from 'discord.js';
import { renderMarkdown, RenderOptions } from 'cordmd';
import logger from '../utils/logger';

export class MarkdownRenderService {
  constructor() {
    // Logger is a singleton, so we don't need to instantiate it
  }

  /**
   * Renders markdown to an image attachment
   * @param markdown - Markdown text to render
   * @param options - Customization options
   * @returns Promise<AttachmentBuilder> - Discord attachment with rendered image
   */
  async renderToImage(markdown: string, options?: RenderOptions): Promise<AttachmentBuilder> {
    try {
      // Default options for Discord theme
      const defaultOptions: RenderOptions = {
        backgroundColor: '#36393F',
        textColor: '#FFFFFF',
        width: 800,
        height: 600,
        fontSize: 16,
        fontFamily: 'sans-serif'
      };

      // Merge user options with defaults
      const renderOptions = { ...defaultOptions, ...options };

      // Render markdown to buffer
      const buffer = await renderMarkdown(markdown, renderOptions);
      
      // Create Discord attachment
      const attachment = new AttachmentBuilder(buffer, { name: 'markdown-render.png' });
      
      logger.info('Markdown rendered successfully', { component: 'MarkdownRenderService' });
      return attachment;
    } catch (error) {
      logger.error('Failed to render markdown', { 
        component: 'MarkdownRenderService',
        error: error instanceof Error ? error.message : String(error)
      });
      throw new Error('Не вдалося відобразити markdown контент');
    }
  }

  /**
   * Renders markdown with Discord dark theme customization
   */
  async renderDiscordDarkTheme(markdown: string): Promise<AttachmentBuilder> {
    const options: RenderOptions = {
      backgroundColor: '#36393F',
      textColor: '#FFFFFF',
      headingColors: ['#FF7F7F', '#7FBF7F', '#7F7FBF', '#BFBF7F', '#BF7FBF', '#7FBFBF'],
      borderColor: '#44475A',
      codeBackgroundColor: '#1E1F22',
      codeTextColor: '#D1B57B',
      linkColor: '#58A6FF'
    };
    
    return this.renderToImage(markdown, options);
  }

  /**
   * Renders markdown with light theme customization
   */
  async renderLightTheme(markdown: string): Promise<AttachmentBuilder> {
    const options: RenderOptions = {
      backgroundColor: '#FFFFFF',
      textColor: '#000000',
      headingColors: ['#D32F2F', '#388E3C', '#1976D2', '#F57C00', '#7B1FA2', '#0097A7'],
      borderColor: '#E0E0E0',
      codeBackgroundColor: '#F5F5F5',
      codeTextColor: '#D81B60',
      linkColor: '#1976D2'
    };
    
    return this.renderToImage(markdown, options);
  }
}