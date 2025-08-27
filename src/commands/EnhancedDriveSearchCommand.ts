import {
  SlashCommandBuilder,
  SlashCommandStringOption,
  SlashCommandIntegerOption,
  ChatInputCommandInteraction,
  EmbedBuilder,
  ActionRowBuilder,
  StringSelectMenuBuilder,
  ButtonBuilder,
  ButtonStyle
} from 'discord.js';
import { BaseCommand } from './BaseCommand';
import type { BotConfig, CommandExecuteOptions } from '@/types';
import type { GoogleService } from '@/services/GoogleService';
import type { DriveFile, DriveListQuery, DriveListResult } from '@/types/drive';
import type { SmartDocumentClassifier, ClassifiedDocument } from '@/services/SmartDocumentClassifier';
import logger from '@/utils/logger';
import { signComponentId } from '@/security/componentId';

interface SearchState {
  query: string;
  folderId?: string;
  mimeTypes: string[];
  dateFrom?: string;
  dateTo?: string;
  sizeMin?: number;
  sizeMax?: number;
  sortBy: 'relevance' | 'name' | 'modifiedTime' | 'size';
  sortDir: 'asc' | 'desc';
  pageToken?: string;
  pageSize: number;
  categoryId?: string;
  tags: string[];
  highlightTerms: boolean;
  findSimilar: boolean;
  searchType: 'name' | 'content' | 'both';
}

export class EnhancedDriveSearchCommand extends BaseCommand {
  private readonly google: GoogleService | null;
  private readonly classifier: SmartDocumentClassifier | null;
  private static sessions = new Map<string, SearchState>();

  constructor(config: BotConfig, google?: GoogleService, classifier?: SmartDocumentClassifier) {
    super(
      'drive-search',
      'Розширений пошук документів Google Drive',
      config,
      {
        category: 'documents',
        usage: '/drive-search [query] [filters]',
        examples: [
          '/drive-search query:бюджет',
          '/drive-search query:наказ mime:application/pdf',
          '/drive-search query:звіт date_from:2024-01-01 date_to:2024-12-31',
          '/drive-search query:бюджет highlight:true',
          '/drive-search query:бюджет similar:true'
        ]
      },
      (builder: SlashCommandBuilder) => {
        builder
          .addStringOption((option: SlashCommandStringOption) =>
            option
              .setName('query')
              .setDescription('Пошуковий запит')
              .setRequired(true)
              .setMaxLength(100)
          )
          .addStringOption((option: SlashCommandStringOption) =>
            option
              .setName('folder')
              .setDescription('ID папки для пошуку (за замовчуванням коренева папка)')
              .setRequired(false)
              .setMaxLength(100)
          )
          .addStringOption((option: SlashCommandStringOption) =>
            option
              .setName('mime')
              .setDescription('Фільтр за типом файлу (наприклад, application/pdf)')
              .setRequired(false)
              .setMaxLength(100)
          )
          .addStringOption((option: SlashCommandStringOption) =>
            option
              .setName('date_from')
              .setDescription('Фільтр за датою: від (YYYY-MM-DD)')
              .setRequired(false)
              .setMaxLength(10)
          )
          .addStringOption((option: SlashCommandStringOption) =>
            option
              .setName('date_to')
              .setDescription('Фільтр за датою: до (YYYY-MM-DD)')
              .setRequired(false)
              .setMaxLength(10)
          )
          .addIntegerOption((option: SlashCommandIntegerOption) =>
            option
              .setName('size_min')
              .setDescription('Мінімальний розмір файлу (в байтах)')
              .setRequired(false)
              .setMinValue(0)
          )
          .addIntegerOption((option: SlashCommandIntegerOption) =>
            option
              .setName('size_max')
              .setDescription('Максимальний розмір файлу (в байтах)')
              .setRequired(false)
              .setMinValue(0)
          )
          .addStringOption((option: SlashCommandStringOption) =>
            option
              .setName('sort')
              .setDescription('Сортування результатів')
              .setRequired(false)
              .addChoices(
                { name: 'Релевантність', value: 'relevance' },
                { name: 'Назва (А-Я)', value: 'name_asc' },
                { name: 'Назва (Я-А)', value: 'name_desc' },
                { name: 'Дата зміни (новіші)', value: 'modifiedTime_desc' },
                { name: 'Дата зміни (старіші)', value: 'modifiedTime_asc' },
                { name: 'Розмір (зростання)', value: 'size_asc' },
                { name: 'Розмір (спадання)', value: 'size_desc' }
              )
          )
          .addIntegerOption((option: SlashCommandIntegerOption) =>
            option
              .setName('page_size')
              .setDescription('Кількість елементів на сторінці (5-25)')
              .setRequired(false)
              .setMinValue(5)
              .setMaxValue(25)
          )
          // New options for enhanced functionality
          .addBooleanOption((option: SlashCommandStringOption) =>
            option
              .setName('highlight')
              .setDescription('Виділяти ключові фрази в результатах')
              .setRequired(false)
          )
          .addBooleanOption((option: SlashCommandStringOption) =>
            option
              .setName('similar')
              .setDescription('Шукати схожі документи')
              .setRequired(false)
          )
          .addStringOption((option: SlashCommandStringOption) =>
            option
              .setName('search-type')
              .setDescription('Тип пошуку: в назві, вмісті або в обох')
              .setRequired(false)
              .addChoices(
                { name: 'Назва', value: 'name' },
                { name: 'Вміст', value: 'content' },
                { name: 'Обоє', value: 'both' }
              )
          );
        return builder;
      }
    );

    this.google = google ?? null;
    this.classifier = classifier ?? null;
  }

  protected override async onExecute(options: CommandExecuteOptions): Promise<void> {
    const interaction = options.interaction;
    await interaction.deferReply({ ephemeral: false });

    try {
      if (!this.google) {
        await interaction.editReply('GoogleService недоступен. Обратитесь к администратору.');
        return;
      }

      // Get command options
      const query = interaction.options.getString('query', true);
      const folderId = interaction.options.getString('folder') || this.config.drive?.folderId || 'root';
      const mimeFilter = interaction.options.getString('mime') || undefined;
      const dateFrom = interaction.options.getString('date_from') || undefined;
      const dateTo = interaction.options.getString('date_to') || undefined;
      const sizeMin = interaction.options.getInteger('size_min') || undefined;
      const sizeMax = interaction.options.getInteger('size_max') || undefined;
      const sortOption = interaction.options.getString('sort') || 'relevance';
      const pageSize = interaction.options.getInteger('page_size') || 10;
      const highlight = interaction.options.getBoolean('highlight') || false;
      const similar = interaction.options.getBoolean('similar') || false;
      const searchType = interaction.options.getString('search-type') as 'name' | 'content' | 'both' || 'both';

      // Parse sort option
      let sortBy: 'relevance' | 'name' | 'modifiedTime' | 'size' = 'relevance';
      let sortDir: 'asc' | 'desc' = 'desc';
      
      if (sortOption === 'relevance') {
        sortBy = 'relevance';
        sortDir = 'desc';
      } else if (sortOption.includes('_')) {
        const [field, direction] = sortOption.split('_');
        sortBy = field as 'name' | 'modifiedTime' | 'size';
        sortDir = direction as 'asc' | 'desc';
      }

      // Create search state
      const state: SearchState = {
        query,
        folderId,
        mimeTypes: mimeFilter ? [mimeFilter] : [],
        dateFrom,
        dateTo,
        sizeMin,
        sizeMax,
        sortBy,
        sortDir,
        pageSize: Math.min(25, Math.max(5, pageSize)),
        tags: [],
        highlightTerms: highlight,
        findSimilar: similar,
        searchType
      };

      // Store session
      const sessionId = `${interaction.user.id}_${Date.now()}`;
      EnhancedDriveSearchCommand.sessions.set(sessionId, state);

      // Perform search
      await this.performSearch(interaction, sessionId, state);

    } catch (error) {
      logger.error('Помилка виконання команди /drive-search', {
        type: 'command',
        command: 'drive-search',
        error: error instanceof Error ? error.message : String(error),
      });
      await interaction.editReply('❌ Помилка при пошуку. Спробуйте позже.');
    }
  }

  private async performSearch(
    interaction: ChatInputCommandInteraction,
    sessionId: string,
    state: SearchState
  ): Promise<void> {
    try {
      if (!this.google) {
        await interaction.editReply('GoogleService недоступен.');
        return;
      }

      // Build query for Google Drive API
      const queryParts: string[] = [];
      
      // Add folder constraint if specified
      if (state.folderId && state.folderId !== 'root') {
        queryParts.push(`'${state.folderId}' in parents`);
      }
      
      // Add search query based on search type
      if (state.query) {
        switch (state.searchType) {
          case 'name':
            queryParts.push(`name contains '${state.query}'`);
            break;
          case 'content':
            queryParts.push(`fullText contains '${state.query}'`);
            break;
          case 'both':
          default:
            queryParts.push(`name contains '${state.query}' or fullText contains '${state.query}'`);
            break;
        }
      }
      
      // Add MIME filters
      if (state.mimeTypes.length > 0) {
        const mimeConditions = state.mimeTypes
          .map(mime => `mimeType = '${mime}'`)
          .join(' or ');
        queryParts.push(`(${mimeConditions})`);
      }
      
      // Add date filters
      if (state.dateFrom) {
        const fromDate = new Date(state.dateFrom);
        if (!isNaN(fromDate.getTime())) {
          queryParts.push(`modifiedTime >= '${fromDate.toISOString()}'`);
        }
      }
      
      if (state.dateTo) {
        const toDate = new Date(state.dateTo);
        if (!isNaN(toDate.getTime())) {
          toDate.setHours(23, 59, 59, 999);
          queryParts.push(`modifiedTime <= '${toDate.toISOString()}'`);
        }
      }
      
      // Build final query
      const finalQuery = queryParts.length > 0 ? queryParts.join(' and ') : undefined;

      // Prepare Drive API query
      const driveQuery: DriveListQuery = {
        folderId: state.folderId,
        query: finalQuery,
        pageSize: state.pageSize,
        pageToken: state.pageToken,
        sortBy: state.sortBy !== 'relevance' ? state.sortBy : undefined,
        sortDir: state.sortBy !== 'relevance' ? state.sortDir : undefined,
        mimeIncludes: state.mimeTypes.length > 0 ? state.mimeTypes : undefined,
        dateFrom: state.dateFrom,
        dateTo: state.dateTo,
        sizeMin: state.sizeMin,
        sizeMax: state.sizeMax
      };

      // Fetch files
      const result: DriveListResult = await this.google.listDriveFiles(driveQuery);

      // Find similar documents if requested
      let similarDocuments: DriveFile[] = [];
      if (state.findSimilar && result.files.length > 0) {
        similarDocuments = await this.findSimilarDocuments(result.files[0]);
      }

      // Classify documents if classifier is available
      let classifiedDocuments: ClassifiedDocument[] = [];
      if (this.classifier && result.files.length > 0) {
        // For demo purposes, we'll use a simplified approach
        // In a real implementation, you would extract content and classify
        classifiedDocuments = result.files.map(file => ({
          file,
          categories: [],
          confidence: 0,
          tags: [],
          projectThemes: [],
          relationships: []
        }));
      }

      // Create embed with search results
      const embed = new EmbedBuilder()
        .setTitle('🔍 Розширений пошук Google Drive')
        .setDescription(`Запит: **${state.query}**\nЗнайдено: **${result.files.length}** елементів${state.findSimilar ? ` | Схожих: **${similarDocuments.length}**` : ''}`)
        .setColor(0x4285f4);

      // Add filter info if filtering is active
      if (state.mimeTypes.length > 0 || state.dateFrom || state.dateTo || 
          state.sizeMin !== undefined || state.sizeMax !== undefined ||
          state.searchType !== 'both' || state.highlightTerms) {
        const filters = [];
        if (state.mimeTypes.length > 0) filters.push(`📎 Типи: ${state.mimeTypes.map(m => this.getMimeTypeLabel(m)).join(', ')}`);
        if (state.dateFrom) filters.push(`📅 Від: ${state.dateFrom}`);
        if (state.dateTo) filters.push(`📅 До: ${state.dateTo}`);
        if (state.sizeMin !== undefined) filters.push(`📊 Мін. розмір: ${this.formatFileSize(state.sizeMin)}`);
        if (state.sizeMax !== undefined) filters.push(`📊 Макс. розмір: ${this.formatFileSize(state.sizeMax)}`);
        if (state.searchType !== 'both') filters.push(`🔍 Тип пошуку: ${state.searchType === 'name' ? 'Назва' : 'Вміст'}`);
        if (state.highlightTerms) filters.push(`✨ Виділення: увімкнено`);
        if (state.findSimilar) filters.push(`🔗 Схожі: увімкнено`);
        
        embed.addFields({
          name: '📊 Фільтри',
          value: filters.join('\n')
        });
      }

      // Add search results
      if (result.files.length > 0) {
        embed.addFields({
          name: `📄 Результати пошуку`,
          value: this.formatSearchResults(result.files, state.highlightTerms ? state.query : undefined)
        });
      } else {
        embed.addFields({
          name: '📄 Результати пошуку',
          value: 'Нічого не знайдено за заданими критеріями'
        });
      }

      // Add similar documents if requested
      if (state.findSimilar && similarDocuments.length > 0) {
        embed.addFields({
          name: `🔗 Схожі документи`,
          value: this.formatSearchResults(similarDocuments, undefined, 'Схожі документи')
        });
      }

      // Create action components
      const components = this.createSearchComponents(sessionId, state, result, similarDocuments);

      await interaction.editReply({
        embeds: [embed],
        components: components
      });

    } catch (error) {
      logger.error('Помилка виконання пошуку', {
        type: 'command',
        command: 'drive-search',
        sessionId,
        error: error instanceof Error ? error.message : String(error),
      });
      await interaction.editReply('❌ Помилка при виконанні пошуку.');
    }
  }

  private async findSimilarDocuments(file: DriveFile): Promise<DriveFile[]> {
    // This is a simplified implementation
    // In a real application, you would use content analysis or embeddings to find similar documents
    try {
      if (!this.google) return [];
      
      // For demo purposes, we'll just search for files with similar names
      const keywords = (file.name || '').split(' ')
        .filter(word => word.length > 3)
        .slice(0, 3);
      
      if (keywords.length === 0) return [];
      
      const query = keywords.map(k => `name contains '${k}'`).join(' or ');
      
      const result = await this.google.listDriveFiles({
        query,
        pageSize: 5
      });
      
      // Filter out the original file
      return result.files.filter(f => f.id !== file.id);
    } catch (error) {
      logger.error('Помилка пошуку схожих документів', {
        type: 'command',
        command: 'drive-search',
        fileId: file.id,
        error: error instanceof Error ? error.message : String(error),
      });
      return [];
    }
  }

  private formatSearchResults(files: DriveFile[], highlightQuery?: string, title: string = 'Результати пошуку'): string {
    if (files.length === 0) return 'Нічого не знайдено';
    
    const items = files.slice(0, 15).map(file => {
      const isFolder = file.mimeType === 'application/vnd.google-apps.folder';
      const icon = isFolder ? '📁' : this.getMimeTypeIcon(file.mimeType);
      const size = file.size ? this.formatFileSize(file.size) : '';
      const modified = file.modifiedTime 
        ? `<t:${Math.floor(new Date(file.modifiedTime).getTime() / 1000)}:R>` 
        : '';
      
      // Truncate long names
      let displayName = file.name && file.name.length > 30 
        ? file.name.substring(0, 27) + '...' 
        : file.name || 'Без назви';
      
      // Highlight query terms if requested
      if (highlightQuery && !isFolder) {
        const queryTerms = highlightQuery.toLowerCase().split(/\s+/);
        for (const term of queryTerms) {
          if (term.length > 1) {
            const regex = new RegExp(`(${term})`, 'gi');
            displayName = displayName.replace(regex, '**$1**');
          }
        }
      }
      
      return `${icon} **${displayName}** ${size} ${modified}`;
    });
    
    let result = items.join('\n');
    
    // Add "and X more" if there are more items
    if (files.length > 15) {
      result += `\n\n...і ще ${files.length - 15} елементів`;
    }
    
    return result.length > 1024 ? result.substring(0, 1021) + '...' : result;
  }

  private getMimeTypeIcon(mimeType: string = ''): string {
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

  private getMimeTypeLabel(mimeType: string): string {
    const labelMap: Record<string, string> = {
      'application/pdf': 'PDF',
      'application/vnd.google-apps.document': 'Google Docs',
      'application/vnd.google-apps.spreadsheet': 'Google Sheets',
      'application/vnd.google-apps.presentation': 'Google Slides',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'Word',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'Excel',
      'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'PowerPoint',
      'text/plain': 'Текстовий файл',
      'image/': 'Зображення',
      'video/': 'Відео',
      'audio/': 'Аудіо'
    };
    
    for (const [key, label] of Object.entries(labelMap)) {
      if (mimeType.startsWith(key) || mimeType.includes(key)) {
        return label;
      }
    }
    
    return mimeType;
  }

  private formatFileSize(bytes: number): string {
    if (bytes === 0) return '0 Bytes';
    
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  }

  private createSearchComponents(
    sessionId: string,
    state: SearchState,
    result: DriveListResult,
    similarDocuments: DriveFile[] = []
  ): ActionRowBuilder<any>[] {
    const components: ActionRowBuilder<any>[] = [];
    
    // Create file selection dropdown if there are items
    if (result.files.length > 0) {
      const selectMenu = new StringSelectMenuBuilder()
        .setCustomId(signComponentId(`drive-search-select-${sessionId}`))
        .setPlaceholder('Оберіть файл для перегляду')
        .setMaxValues(1);
      
      // Add up to 25 items to the dropdown
      const items = result.files.slice(0, 25).map(file => {
        const isFolder = file.mimeType === 'application/vnd.google-apps.folder';
        const icon = isFolder ? '📁' : this.getMimeTypeIcon(file.mimeType);
        const displayName = (file.name && file.name.length > 50) 
          ? file.name.substring(0, 47) + '...' 
          : file.name || 'Без назви';
        
        return {
          label: `${icon} ${displayName}`,
          value: `${isFolder ? 'folder' : 'file'}_${file.id}`,
          description: isFolder ? 'Папка' : this.getMimeTypeLabel(file.mimeType)
        };
      });
      
      selectMenu.addOptions(items);
      components.push(new ActionRowBuilder().addComponents(selectMenu));
    }
    
    // Create navigation buttons
    const buttonRow = new ActionRowBuilder();
    
    // Refresh button
    const refreshButton = new ButtonBuilder()
      .setCustomId(signComponentId(`drive-search-refresh-${sessionId}`))
      .setLabel('🔄 Оновити')
      .setStyle(ButtonStyle.Primary);
    
    // Highlight toggle button
    const highlightButton = new ButtonBuilder()
      .setCustomId(signComponentId(`drive-search-highlight-${sessionId}`))
      .setLabel(state.highlightTerms ? '✨ Без виділення' : '✨ Виділити')
      .setStyle(state.highlightTerms ? ButtonStyle.Secondary : ButtonStyle.Primary);
    
    // Similar documents button
    const similarButton = new ButtonBuilder()
      .setCustomId(signComponentId(`drive-search-similar-${sessionId}`))
      .setLabel('🔗 Схожі')
      .setStyle(ButtonStyle.Primary);
    
    buttonRow.addComponents(refreshButton, highlightButton, similarButton);
    
    // Add pagination buttons if needed
    if (result.nextPageToken) {
      const nextButton = new ButtonBuilder()
        .setCustomId(signComponentId(`drive-search-next-${sessionId}`))
        .setLabel('➡️ Далі')
        .setStyle(ButtonStyle.Secondary);
      
      buttonRow.addComponents(nextButton);
    }
    
    if (components.length < 5) { // Discord limit is 5 action rows
      components.push(buttonRow);
    }
    
    return components;
  }
}