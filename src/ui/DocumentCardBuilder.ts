import { EmbedBuilder, ActionRowBuilder, ButtonBuilder, ButtonStyle } from 'discord.js';
import type { DriveFile } from '@/types/drive';
import { signComponentId } from '@/security/componentId';

export interface DocumentCardOptions {
  showPreview?: boolean;
  showStats?: boolean;
  showActions?: boolean;
  maxPreviewLength?: number;
  // New options for enhanced functionality
  showHistory?: boolean;
  showQuickActions?: boolean;
}

// New interface for interaction history
export interface DocumentInteraction {
  userId: string;
  action: string;
  timestamp: Date;
  details?: any;
}

export class DocumentCardBuilder {
  private file: DriveFile;
  private options: DocumentCardOptions;
  private previewContent?: string;
  private tags: string[] = [];
  // New properties for enhanced functionality
  private interactions: DocumentInteraction[] = [];
  private quickActions: { label: string; action: string; emoji: string }[] = [];

  constructor(file: DriveFile, options: DocumentCardOptions = {}) {
    this.file = file;
    this.options = {
      showPreview: options.showPreview ?? true,
      showStats: options.showStats ?? true,
      showActions: options.showActions ?? true,
      maxPreviewLength: options.maxPreviewLength ?? 500,
      // New options with defaults
      showHistory: options.showHistory ?? true,
      showQuickActions: options.showQuickActions ?? true
    };
  }

  /**
   * Встановлює вміст попереднього перегляду
   */
  setPreview(content: string): this {
    this.previewContent = content;
    return this;
  }

  /**
   * Встановлює теги документа
   */
  setTags(tags: string[]): this {
    this.tags = tags;
    return this;
  }

  /**
   * Встановлює історію взаємодій з документом
   */
  setInteractions(interactions: DocumentInteraction[]): this {
    this.interactions = interactions;
    return this;
  }

  /**
   * Встановлює швидкі дії для документа
   */
  setQuickActions(actions: { label: string; action: string; emoji: string }[]): this {
    this.quickActions = actions;
    return this;
  }

  /**
   * Створює картку документа
   */
  build(sessionId: string): { embed: EmbedBuilder; components: ActionRowBuilder<any>[] } {
    const embed = this.createEmbed();
    const components = this.createComponents(sessionId);
    
    return { embed, components };
  }

  private createEmbed(): EmbedBuilder {
    const embed = new EmbedBuilder()
      .setTitle(this.getDocumentTitle())
      .setDescription(this.getDocumentDescription())
      .setColor(this.getDocumentColor())
      .setTimestamp(this.file.modifiedTime ? new Date(this.file.modifiedTime) : undefined);

    // Додаємо попередній перегляд якщо потрібно
    if (this.options.showPreview && this.previewContent) {
      const preview = this.truncateText(this.previewContent, this.options.maxPreviewLength);
      embed.addFields({
        name: '🔍 Попередній перегляд',
        value: `\`\`\`${preview}\`\`\``
      });
    }

    // Додаємо статистику якщо потрібно
    if (this.options.showStats) {
      const stats = this.getDocumentStats();
      if (stats) {
        embed.addFields({
          name: '📊 Статистика',
          value: stats
        });
      }
    }

    // Додаємо історію взаємодій якщо потрібно
    if (this.options.showHistory && this.interactions.length > 0) {
      const history = this.formatInteractionHistory();
      if (history) {
        embed.addFields({
          name: '🕒 Історія взаємодій',
          value: history
        });
      }
    }

    // Додаємо теги якщо є
    if (this.tags.length > 0) {
      embed.addFields({
        name: '🏷️ Теги',
        value: this.tags.map(tag => `\`${tag}\``).join(' ')
      });
    }

    return embed;
  }

  private formatInteractionHistory(): string | null {
    if (this.interactions.length === 0) return null;
    
    // Sort interactions by timestamp (newest first)
    const sortedInteractions = [...this.interactions].sort((a, b) => 
      b.timestamp.getTime() - a.timestamp.getTime()
    );
    
    // Take only the last 5 interactions
    const recentInteractions = sortedInteractions.slice(0, 5);
    
    const historyItems = recentInteractions.map(interaction => {
      const actionEmoji = this.getActionEmoji(interaction.action);
      const timeAgo = this.getTimeAgo(interaction.timestamp);
      return `${actionEmoji} **${interaction.action}** ${timeAgo}`;
    });
    
    return historyItems.join('\n');
  }

  private getActionEmoji(action: string): string {
    const emojiMap: Record<string, string> = {
      'view': '👁️',
      'download': '📥',
      'edit': '✏️',
      'share': '🔗',
      'comment': '💬',
      'analyze': '🧠',
      'export': '📤',
      'tag': '🏷️',
      'history': '🕒'
    };
    
    return emojiMap[action.toLowerCase()] || '📋';
  }

  private getTimeAgo(timestamp: Date): string {
    const now = new Date();
    const diffMs = now.getTime() - timestamp.getTime();
    const diffMins = Math.floor(diffMs / 60000);
    const diffHours = Math.floor(diffMins / 60);
    const diffDays = Math.floor(diffHours / 24);
    
    if (diffMins < 1) return 'тільки що';
    if (diffMins < 60) return `${diffMins} хв. тому`;
    if (diffHours < 24) return `${diffHours} год. тому`;
    return `${diffDays} дн. тому`;
  }

  private createComponents(sessionId: string): ActionRowBuilder<any>[] {
    const components: ActionRowBuilder<any>[] = [];

    if (this.options.showActions) {
      const actionRow = new ActionRowBuilder();

      // Кнопка аналізу
      const analyzeButton = new ButtonBuilder()
        .setCustomId(signComponentId({ kind: 'doc-analyze', fileId: this.file.id, sid: sessionId }))
        .setLabel('Аналіз')
        .setStyle(ButtonStyle.Primary)
        .setEmoji('🧠');

      // Кнопка експорту
      const exportButton = new ButtonBuilder()
        .setCustomId(signComponentId({ kind: 'doc-export', fileId: this.file.id, sid: sessionId }))
        .setLabel('Експорт')
        .setStyle(ButtonStyle.Secondary)
        .setEmoji('📤');

      // Кнопка тегування
      const tagButton = new ButtonBuilder()
        .setCustomId(signComponentId({ kind: 'doc-tag', fileId: this.file.id, sid: sessionId }))
        .setLabel('Теги')
        .setStyle(ButtonStyle.Secondary)
        .setEmoji('🏷️');

      // Кнопка історії
      const historyButton = new ButtonBuilder()
        .setCustomId(signComponentId({ kind: 'doc-history', fileId: this.file.id, sid: sessionId }))
        .setLabel('Історія')
        .setStyle(ButtonStyle.Secondary)
        .setEmoji('🕒');

      actionRow.addComponents(analyzeButton, exportButton, tagButton, historyButton);
      components.push(actionRow);
    }

    // Add quick actions if enabled
    if (this.options.showQuickActions && this.quickActions.length > 0) {
      const quickActionRow = new ActionRowBuilder();
      
      // Add up to 4 quick actions
      const actionsToAdd = this.quickActions.slice(0, 4);
      
      for (const action of actionsToAdd) {
        const button = new ButtonBuilder()
          .setCustomId(signComponentId({ kind: 'doc-quick', action: action.action, fileId: this.file.id, sid: sessionId }))
          .setLabel(action.label)
          .setStyle(ButtonStyle.Success)
          .setEmoji(action.emoji);
        
        quickActionRow.addComponents(button);
      }
      
      if (components.length < 4) { // Discord limit is 5 action rows, reserve one for pagination
        components.push(quickActionRow);
      }
    }

    return components;
  }

  private getDocumentTitle(): string {
    const icon = this.getMimeTypeIcon();
    const name = this.file.name || 'Без назви';
    return `${icon} ${name}`;
  }

  private getDocumentDescription(): string {
    const type = this.getMimeTypeLabel();
    const owner = this.file.owners && this.file.owners.length > 0 
      ? `Власник: ${this.file.owners[0]}` 
      : '';
    
    return `${type}${owner ? `\n${owner}` : ''}`;
  }

  private getDocumentStats(): string | null {
    const stats = [];

    // Розмір файлу
    if (typeof this.file.size === 'number') {
      stats.push(`⚖️ Розмір: ${this.formatFileSize(this.file.size)}`);
    }

    // Дата зміни
    if (this.file.modifiedTime) {
      stats.push(`🕒 Змінено: <t:${Math.floor(new Date(this.file.modifiedTime).getTime() / 1000)}:R>`);
    }

    // Посилання на перегляд
    if (this.file.webViewLink) {
      stats.push(`🔗 [Відкрити в Google Drive](${this.file.webViewLink})`);
    }

    // Add interaction count if history is enabled
    if (this.options.showHistory && this.interactions.length > 0) {
      stats.push(`👥 Взаємодій: ${this.interactions.length}`);
    }

    return stats.length > 0 ? stats.join('\n') : null;
  }

  private getMimeTypeIcon(): string {
    const mimeType = this.file.mimeType || '';
    
    const iconMap: Record<string, string> = {
      'application/pdf': '📄',
      'application/vnd.google-apps.document': '📝',
      'application/vnd.google-apps.spreadsheet': '📊',
      'application/vnd.google-apps.presentation': '📽️',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document': '📝',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': '📊',
      'application/vnd.openxmlformats-officedocument.presentationml.presentation': '📽️',
      'text/plain': '📄',
      'image/': '🖼️',
      'video/': '🎬',
      'audio/': '🎵'
    };
    
    for (const [key, icon] of Object.entries(iconMap)) {
      if (mimeType.startsWith(key) || mimeType.includes(key)) {
        return icon;
      }
    }
    
    return '📎'; // Default file icon
  }

  private getMimeTypeLabel(): string {
    const mimeType = this.file.mimeType || '';
    
    const labelMap: Record<string, string> = {
      'application/pdf': 'PDF документ',
      'application/vnd.google-apps.document': 'Google Docs',
      'application/vnd.google-apps.spreadsheet': 'Google Sheets',
      'application/vnd.google-apps.presentation': 'Google Slides',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'Word документ',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'Excel таблиця',
      'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'PowerPoint презентація',
      'text/plain': 'Текстовий файл',
      'image/': 'Зображення',
      'video/': 'Відео файл',
      'audio/': 'Аудіо файл'
    };
    
    for (const [key, label] of Object.entries(labelMap)) {
      if (mimeType.startsWith(key) || mimeType.includes(key)) {
        return label;
      }
    }
    
    return 'Файл';
  }

  private getDocumentColor(): number {
    const mimeType = this.file.mimeType || '';
    
    const colorMap: Record<string, number> = {
      'application/pdf': 0xff0000, // Red
      'application/vnd.google-apps.document': 0x4285f4, // Blue
      'application/vnd.google-apps.spreadsheet': 0x0f9d58, // Green
      'application/vnd.google-apps.presentation': 0xf4b400, // Yellow
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 0x4285f4, // Blue
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 0x0f9d58, // Green
      'application/vnd.openxmlformats-officedocument.presentationml.presentation': 0xf4b400, // Yellow
      'text/plain': 0x9e9e9e, // Gray
      'image/': 0x9c27b0, // Purple
      'video/': 0xff9800, // Orange
      'audio/': 0x3f51b5 // Indigo
    };
    
    for (const [key, color] of Object.entries(colorMap)) {
      if (mimeType.startsWith(key) || mimeType.includes(key)) {
        return color;
      }
    }
    
    return 0x4285f4; // Default blue
  }

  private formatFileSize(bytes: number): string {
    if (bytes === 0) return '0 Bytes';
    
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  }

  private truncateText(text: string, maxLength: number = 500): string {
    if (text.length <= maxLength) return text;
    
    // Обрізаємо текст та додаємо три крапки
    return text.substring(0, maxLength - 3) + '...';
  }
}