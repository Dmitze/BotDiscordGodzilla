import { AttachmentBuilder, EmbedBuilder } from 'discord.js';
import { MarkdownRenderingService } from '@/services/MarkdownRenderingService';
import type { BotConfig } from '@/types';

/**
 * Enhanced Discord message formatter with markdown rendering capabilities
 * Provides beautiful formatting for Discord responses using the CordMd library
 */

export interface FormattedResponse {
  content?: string;
  embeds?: EmbedBuilder[];
  files?: AttachmentBuilder[];
}

export interface FormatOptions {
  format?: 'text' | 'image' | 'embed' | 'mixed';
  theme?: 'light' | 'dark';
  fontSize?: number;
  maxWidth?: number;
  maxHeight?: number;
  language?: string;
}

export class DiscordMarkdownFormatter {
  private markdownService: MarkdownRenderingService;

  constructor(config: BotConfig) {
    this.markdownService = MarkdownRenderingService.getInstance(config);
  }

  /**
   * Format markdown content for Discord with enhanced rendering
   * @param content The markdown content to format
   * @param options Formatting options
   * @returns Formatted response with appropriate Discord message components
   */
  public async formatMarkdown(content: string, options: FormatOptions = {}): Promise<FormattedResponse> {
    const format = options.format || 'text';
    
    switch (format) {
      case 'image':
        return await this.formatAsImage(content, options);
      case 'embed':
        return await this.formatAsEmbed(content, options);
      case 'mixed':
        return await this.formatAsMixed(content, options);
      case 'text':
      default:
        return await this.formatAsText(content, options);
    }
  }

  /**
   * Format markdown as enhanced text with Discord-native formatting
   * @param content The markdown content to format
   * @param options Formatting options
   * @returns Formatted response with enhanced text
   */
  private async formatAsText(content: string, _options: FormatOptions): Promise<FormattedResponse> {
    // Use the markdown service to render to text with enhancements
    const renderedText = await this.markdownService.renderToText(content);
    
    return {
      content: renderedText
    };
  }

  /**
   * Format markdown as an image attachment
   * @param content The markdown content to format
   * @param options Formatting options
   * @returns Formatted response with image attachment
   */
  private async formatAsImage(content: string, options: FormatOptions): Promise<FormattedResponse> {
    // Use the markdown service to render to image
    const attachment = await this.markdownService.renderToImage(content, {
      theme: options.theme,
      fontSize: options.fontSize,
      maxWidth: options.maxWidth,
      maxHeight: options.maxHeight
    });
    
    return {
      files: [attachment]
    };
  }

  /**
   * Format markdown as a Discord embed
   * @param content The markdown content to format
   * @param options Formatting options
   * @returns Formatted response with embed
   */
  private async formatAsEmbed(content: string, options: FormatOptions): Promise<FormattedResponse> {
    // Extract title from markdown (first heading)
    const titleMatch = content.match(/^#\s+(.+)$/m);
    const title = titleMatch ? titleMatch[1].substring(0, 256) : 'Document';
    
    // Extract description (first paragraph or content without headings)
    const descriptionMatch = content.match(/^[^#].*$/m);
    const description = descriptionMatch ? descriptionMatch[0].substring(0, 4096) : '';
    
    // Create embed with enhanced formatting
    const embed = new EmbedBuilder()
      .setTitle(title)
      .setDescription(await this.markdownService.renderToText(description || content))
      .setColor(options.theme === 'dark' ? '#404EED' : '#FFFFFF')
      .setTimestamp();
    
    // Add fields for code blocks if present
    const codeBlocks = this.markdownService.extractCodeBlocks(content);
    if (codeBlocks.length > 0) {
      // Add up to 25 fields (Discord limit)
      const limitedBlocks = codeBlocks.slice(0, 25);
      for (const block of limitedBlocks) {
        const fieldName = block.language ? `\`${block.language}\`` : 'Code Block';
        const fieldValue = `\`\`\`${block.language}\n${block.content.substring(0, 1000)}\n\`\`\``;
        embed.addFields({ name: fieldName, value: fieldValue, inline: false });
      }
    }
    
    return {
      embeds: [embed]
    };
  }

  /**
   * Format markdown as a mixed response with both text and image
   * @param content The markdown content to format
   * @param options Formatting options
   * @returns Formatted response with both text and image
   */
  private async formatAsMixed(content: string, options: FormatOptions): Promise<FormattedResponse> {
    // Format as text for the main content
    const textResponse = await this.formatAsText(content, options);
    
    // Format complex sections as image
    const imageResponse = await this.formatAsImage(content, options);
    
    return {
      content: textResponse.content,
      files: imageResponse.files
    };
  }

  /**
   * Create interactive buttons for enhanced user experience
   * @param content The markdown content
   * @returns Array of action buttons
   */
  public createActionButtons(content: string): Array<{ label: string; customId: string; style: number }> {
    const buttons = [];
    
    // Add "Show Full Document" button if content is long
    if (content.length > 2000) {
      buttons.push({
        label: 'Show Full Document',
        customId: 'show_full_document',
        style: 1 // Primary
      });
    }
    
    // Add "Generate Summary" button
    buttons.push({
      label: 'Generate Summary',
      customId: 'generate_summary',
      style: 2 // Secondary
    });
    
    // Add "Find Similar Documents" button if we have search capabilities
    buttons.push({
      label: 'Find Similar Documents',
      customId: 'find_similar',
      style: 2 // Secondary
    });
    
    return buttons.slice(0, 5); // Discord limit is 5 buttons per row
  }

  /**
   * Extract and format code blocks with syntax highlighting
   * @param content The markdown content
   * @returns Formatted code blocks
   */
  public formatCodeBlocks(content: string): string {
    const codeBlocks = this.markdownService.extractCodeBlocks(content);
    if (codeBlocks.length === 0) return content;
    
    let formattedContent = content;
    
    // Replace code blocks with enhanced formatting
    for (const block of codeBlocks) {
      const originalBlock = `\`\`\`${block.language}\n${block.content}\n\`\`\``;
      const enhancedBlock = `\`\`\`${block.language}\n${block.content}\n\`\`\``;
      formattedContent = formattedContent.replace(originalBlock, enhancedBlock);
    }
    
    return formattedContent;
  }

  /**
   * Format tables for better Discord display
   * @param content The markdown content
   * @returns Formatted content with enhanced tables
   */
  public formatTables(content: string): string {
    // Simple table formatting enhancement
    return content.replace(/(\|[^\n]*\|\n\|[^\n]*\|\n\|[^\n]*\|)/g, (table) => {
      return `\n${table}\n`;
    });
  }

  /**
   * Apply all formatting enhancements to content
   * @param content The markdown content
   * @returns Fully formatted content
   */
  public async applyAllFormatting(content: string, options: FormatOptions = {}): Promise<FormattedResponse> {
    return await this.formatMarkdown(content, { format: 'mixed', ...options });
  }
}