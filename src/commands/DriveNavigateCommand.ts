import {
  EmbedBuilder,
  ActionRowBuilder,
  StringSelectMenuBuilder,
  ButtonBuilder,
  ButtonStyle,
  type ChatInputCommandInteraction,
  type SlashCommandBuilder,
  type SlashCommandStringOption,
  type SlashCommandIntegerOption,
} from 'discord.js';
import { BaseCommand } from './BaseCommand';
import type { BotConfig, CommandExecuteOptions } from '@/types';
import type { GoogleService } from '@/services/GoogleService';
import type { DriveFile, DriveListQuery, DriveListResult } from '@/types/drive';
import logger from '@/utils/logger';
import { signComponentId } from '@/security/componentId';
// import { t } from '@/i18n';

interface NavigationState {
  folderId: string;
  query?: string | undefined;
  pageSize: number;
  pageToken?: string | undefined;
  sortBy: 'name' | 'modifiedTime' | 'size'; // Removed 'relevance' as it's not supported in DriveListQuery
  sortDir: 'asc' | 'desc';
  mimeIncludes?: string[] | undefined;
  dateFrom?: string | undefined;
  dateTo?: string | undefined;
  sizeMin?: number | undefined;
  sizeMax?: number | undefined;
  // Added missing properties with proper optional types
  mimeFilter?: string | undefined;
  searchType?: 'folder' | 'fulltext' | undefined;
  showHidden?: boolean | undefined;
  fileTypeCategories?: string[] | undefined;
  path?: Array<{ id: string; name: string }> | undefined;
  parentId?: string | undefined;
  // Fix: Make all properties explicitly optional to match exactOptionalPropertyTypes
  [key: string]: any; // This allows any additional properties
}

export class DriveNavigateCommand extends BaseCommand {
  private readonly google: GoogleService | null;
  private static sessions = new Map<string, NavigationState>();

  constructor(config: BotConfig, google?: GoogleService) {
    super(
      'drive-navigate',
      'Інтерактивний навігатор по структурі Google Drive',
      config,
      {
        category: 'documents',
        usage: '/drive-navigate [folder] [query] [search-type]',
        examples: [
          '/drive-navigate',
          '/drive-navigate folder:1A2B3C4D5E6F7G8H9I0J',
          '/drive-navigate query:report',
          '/drive-navigate folder:1A2B3C4D5E6F7G8H9I0J query:financial',
          '/drive-navigate query:budget search-type:fulltext'
        ]
      },
      (builder: SlashCommandBuilder) => {
        builder
          .addStringOption((option: SlashCommandStringOption) =>
            option
              .setName('folder')
              .setDescription('ID папки для навігації (за замовчуванням коренева папка)')
              .setRequired(false)
              .setMaxLength(100)
          )
          .addStringOption((option: SlashCommandStringOption) =>
            option
              .setName('query')
              .setDescription('Пошуковий запит для фільтрації файлів')
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
              .setName('sort')
              .setDescription('Сортування результатів')
              .setRequired(false)
              .addChoices(
                { name: 'Назва (А-Я)', value: 'name_asc' },
                { name: 'Назва (Я-А)', value: 'name_desc' },
                { name: 'Дата зміни (новіші)', value: 'modifiedTime_desc' },
                { name: 'Дата зміни (старіші)', value: 'modifiedTime_asc' },
                { name: 'Розмір (зростання)', value: 'size_asc' },
                { name: 'Розмір (спадання)', value: 'size_desc' },
                { name: 'Релевантність', value: 'relevance' }
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
              .setName('search-type')
              .setDescription('Тип пошуку: folder (в поточній папці) або fulltext (по всіх документах)')
              .setRequired(false)
              .addChoices(
                { name: 'В поточній папці', value: 'folder' },
                { name: 'По всіх документах', value: 'fulltext' }
              )
          )
          .addBooleanOption((option) =>
            option
              .setName('show-hidden')
              .setDescription('Показувати приховані файли')
              .setRequired(false)
          )
          .addStringOption((option: SlashCommandStringOption) =>
            option
              .setName('file-category')
              .setDescription('Категорія файлів')
              .setRequired(false)
              .addChoices(
                { name: 'Документи', value: 'documents' },
                { name: 'Таблиці', value: 'spreadsheets' },
                { name: 'Презентації', value: 'presentations' },
                { name: 'Зображення', value: 'images' },
                { name: 'Відео', value: 'videos' },
                { name: 'Аудіо', value: 'audio' },
                { name: 'Архіви', value: 'archives' }
              )
          );
        return builder;
      }
    );

    this.google = google ?? null;
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
      const folderId = interaction.options.getString('folder') || this.config.drive?.folderId || 'root';
      const query = interaction.options.getString('query') || undefined;
      const mimeFilter = interaction.options.getString('mime') || undefined;
      const sortOption = interaction.options.getString('sort') || 'name_asc';
      const pageSize = interaction.options.getInteger('page_size') || 10;
      const dateFrom = interaction.options.getString('date_from') || undefined;
      const dateTo = interaction.options.getString('date_to') || undefined;
      const sizeMin = interaction.options.getInteger('size_min') || undefined;
      const sizeMax = interaction.options.getInteger('size_max') || undefined;
      const searchType = interaction.options.getString('search-type') as 'folder' | 'fulltext' || 'folder';
      const showHidden = interaction.options.getBoolean('show-hidden') || false;
      const fileCategory = interaction.options.getString('file-category') || undefined;

      // Parse sort option
      let sortBy: 'name' | 'modifiedTime' | 'size' = 'name';
      let sortDir: 'asc' | 'desc' = 'asc';
      
      if (sortOption === 'relevance') {
        sortBy = 'name'; // Fallback to name sorting since relevance is not supported
        sortDir = 'desc';
      } else if (sortOption.includes('_')) {
        const [field, direction] = sortOption.split('_');
        sortBy = field as 'name' | 'modifiedTime' | 'size';
        sortDir = direction as 'asc' | 'desc';
      }

      // Create initial navigation state with proper handling of optional properties
      const state: NavigationState = {
        folderId,
        query,
        pageSize: Math.min(25, Math.max(5, pageSize)),
        sortBy,
        sortDir,
        mimeFilter: mimeFilter !== null ? mimeFilter : undefined,
        dateFrom: dateFrom !== null ? dateFrom : undefined,
        dateTo: dateTo !== null ? dateTo : undefined,
        sizeMin: sizeMin !== undefined && sizeMin !== null ? sizeMin : undefined,
        sizeMax: sizeMax !== undefined && sizeMax !== null ? sizeMax : undefined,
        searchType: searchType !== null ? searchType : undefined,
        showHidden: showHidden !== null ? showHidden : undefined,
        fileTypeCategories: fileCategory ? [fileCategory] : undefined,
        path: [{ id: folderId, name: folderId === 'root' ? 'Коренева папка' : 'Обрана папка' }],
        parentId: folderId !== 'root' ? folderId : undefined
      };

      // Store session
      const sessionId = `${interaction.user.id}_${Date.now()}`;
      DriveNavigateCommand.sessions.set(sessionId, state);

      // Load initial content
      await this.displayFolderContent(interaction, sessionId, state);

    } catch (error) {
      logger.error('Помилка виконання команди /drive-navigate', {
        type: 'command',
        command: 'drive-navigate',
        error: error instanceof Error ? error.message : String(error),
      });
      await interaction.editReply('❌ Помилка при навігації. Спробуйте позже.');
    }
  }

  private async displayFolderContent(
    interaction: ChatInputCommandInteraction,
    sessionId: string,
    state: NavigationState
  ): Promise<void> {
    try {
      if (!this.google) {
        await interaction.editReply('GoogleService недоступен.');
        return;
      }

      // Build query for Google Drive API
      const queryParts: string[] = [];
      
      // Add folder constraint for folder search
      if (state.searchType === 'folder') {
        queryParts.push(`'${state.folderId}' in parents`);
      }
      
      // Add search query if provided
      if (state.query) {
        if (state.searchType === 'fulltext') {
          // For fulltext search, we use the query parameter in the API call
          // queryParts.push(`fullText contains '${state.query}'`);
        } else {
          queryParts.push(`name contains '${state.query}'`);
        }
      }
      
      // Add MIME filter if provided
      if (state.mimeFilter) {
        queryParts.push(`mimeType = '${state.mimeFilter}'`);
      }
      
      // Add file type categories filter
      if (state.fileTypeCategories && state.fileTypeCategories.length > 0) {
        const categoryFilters = this.getFileTypeCategoryFilters(state.fileTypeCategories);
        if (categoryFilters.length > 0) {
          queryParts.push(`(${categoryFilters.join(' or ')})`);
        }
      }
      
      // Add date filters if provided
      if (state.dateFrom) {
        // Convert YYYY-MM-DD to RFC3339 format
        const fromDate = new Date(state.dateFrom);
        if (!isNaN(fromDate.getTime())) {
          queryParts.push(`modifiedTime >= '${fromDate.toISOString()}'`);
        }
      }
      
      if (state.dateTo) {
        // Convert YYYY-MM-DD to RFC3339 format and set to end of day
        const toDate = new Date(state.dateTo);
        if (!isNaN(toDate.getTime())) {
          toDate.setHours(23, 59, 59, 999);
          queryParts.push(`modifiedTime <= '${toDate.toISOString()}'`);
        }
      }
      
      // Add size filters if provided
      if (state.sizeMin !== undefined) {
        queryParts.push(`size >= ${state.sizeMin}`);
      }
      
      if (state.sizeMax !== undefined) {
        queryParts.push(`size <= ${state.sizeMax}`);
      }
      
      // Hide hidden files unless explicitly requested
      if (!state.showHidden) {
        queryParts.push(`name not contains '.' or name contains '.doc' or name contains '.pdf' or name contains '.xls' or name contains '.ppt'`);
      }
      
      // Build final query
      const finalQuery = queryParts.length > 0 ? queryParts.join(' and ') : undefined;

      // Prepare Drive API query - removing useFullTextSearch since it doesn't exist
      const driveQuery: DriveListQuery = {
        folderId: state.folderId,
        query: finalQuery !== undefined && finalQuery !== null ? finalQuery : '',
        pageSize: state.pageSize,
        ...(state.pageToken !== undefined ? { pageToken: state.pageToken } : {}),
        sortBy: state.sortBy,
        sortDir: state.sortDir,
        ...(state.mimeFilter ? { mimeIncludes: [state.mimeFilter] } : {}),
        ...(state.dateFrom !== undefined ? { dateFrom: state.dateFrom } : {}),
        ...(state.dateTo !== undefined ? { dateTo: state.dateTo } : {}),
        ...(state.sizeMin !== undefined ? { sizeMin: state.sizeMin } : {}),
        ...(state.sizeMax !== undefined ? { sizeMax: state.sizeMax } : {})
      };

      // For fulltext search, we need to modify the query approach
      if (state.searchType === 'fulltext' && state.query) {
        // We'll handle fulltext search by setting the query directly
        driveQuery.query = state.query;
      }

      // Fetch files
      const result: DriveListResult = await this.google.listDriveFiles(driveQuery);

      // Create embed with navigation info
      const embed = new EmbedBuilder()
        .setTitle('🧭 Навігатор Google Drive')
        .setDescription(this.buildPathBreadcrumb(state.path || []))
        .setColor(0x4285f4)
        .addFields({
          name: `📁 ${state.searchType === 'fulltext' ? 'Результати пошуку' : 'Вміст папки'} (${result.files.length} елементів)`,
          value: result.files.length > 0 
            ? this.formatFileList(result.files) 
            : 'Папка порожня'
        });

      // Add query info if filtering is active
      if (state.query || state.mimeFilter || state.dateFrom || state.dateTo || 
          state.sizeMin !== undefined || state.sizeMax !== undefined ||
          state.fileTypeCategories) {
        const filters = [];
        if (state.query) filters.push(`🔍 Пошук: "${state.query}"`);
        if (state.mimeFilter) filters.push(`📎 Тип: ${this.getMimeTypeLabel(state.mimeFilter)}`);
        if (state.fileTypeCategories) filters.push(`📂 Категорія: ${state.fileTypeCategories.join(', ')}`);
        if (state.dateFrom) filters.push(`📅 Від: ${state.dateFrom}`);
        if (state.dateTo) filters.push(`📅 До: ${state.dateTo}`);
        if (state.sizeMin !== undefined) filters.push(`📊 Мін. розмір: ${this.formatFileSize(state.sizeMin)}`);
        if (state.sizeMax !== undefined) filters.push(`📊 Макс. розмір: ${this.formatFileSize(state.sizeMax)}`);
        if (state.showHidden) filters.push(`👻 Приховані файли: показано`);
        
        embed.addFields({
          name: '📊 Фільтри',
          value: filters.join('\n')
        });
      }

      // Add sorting info
      embed.addFields({
        name: '📊 Сортування',
        value: this.getSortLabel(state.sortBy, state.sortDir) + 
               (state.searchType === 'fulltext' ? ' (за релевантністю)' : '')
      });

      // Create action components
      const components = this.createNavigationComponents(sessionId, state, result);

      await interaction.editReply({
        embeds: [embed],
        components: components
      });

    } catch (error) {
      logger.error('Помилка відображення вмісту папки', {
        type: 'command',
        command: 'drive-navigate',
        sessionId,
        error: error instanceof Error ? error.message : String(error),
      });
      await interaction.editReply('❌ Помилка при завантаженні вмісту папки.');
    }
  }

  private buildPathBreadcrumb(path: Array<{ id: string; name: string }>): string {
    if (path.length === 0) return 'Коренева папка';
    
    const breadcrumb = path.map(item => item.name).join(' > ');
    return breadcrumb.length > 1024 ? breadcrumb.substring(0, 1021) + '...' : breadcrumb;
  }

  private formatFileList(files: DriveFile[]): string {
    if (files.length === 0) return 'Папка порожня';
    
    const items = files.slice(0, 15).map(file => {
      const isFolder = file.mimeType === 'application/vnd.google-apps.folder';
      const icon = isFolder ? '📁' : this.getMimeTypeIcon(file.mimeType);
      const size = file.size ? this.formatFileSize(file.size) : '';
      const modified = file.modifiedTime 
        ? `<t:${Math.floor(new Date(file.modifiedTime).getTime() / 1000)}:R>` 
        : '';
      
      // Truncate long names
      const displayName = file.name.length > 30 
        ? file.name.substring(0, 27) + '...' 
        : file.name;
      
      return `${icon} **${displayName}** ${size} ${modified}`;
    });
    
    let result = items.join('\n');
    
    // Add "and X more" if there are more items
    if (files.length > 15) {
      result += `\n\n...і ще ${files.length - 15} елементів`;
    }
    
    return result.length > 1024 ? result.substring(0, 1021) + '...' : result;
  }

  private getMimeTypeIcon(mimeType: string): string {
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

  private getSortLabel(sortBy: string, sortDir: string): string {
    const fieldLabels: Record<string, string> = {
      'name': 'Назва',
      'modifiedTime': 'Дата зміни',
      'size': 'Розмір'
    };
    
    const directionLabels: Record<string, string> = {
      'asc': '↑',
      'desc': '↓'
    };
    
    return `${fieldLabels[sortBy] || sortBy} ${directionLabels[sortDir] || sortDir}`;
  }

  private createNavigationComponents(
    sessionId: string,
    state: NavigationState,
    result: DriveListResult
  ): ActionRowBuilder<any>[] {
    const components: ActionRowBuilder<any>[] = [];
    
    // Create folder/file selection dropdown if there are items
    if (result.files.length > 0) {
      const selectMenu = new StringSelectMenuBuilder()
        .setCustomId(signComponentId({ kind: 'drive-nav-select', sid: sessionId }))
        .setPlaceholder('Оберіть папку або файл')
        .setMaxValues(1);
      
      // Add up to 25 items to the dropdown
      const items = result.files.slice(0, 25).map(file => {
        const isFolder = file.mimeType === 'application/vnd.google-apps.folder';
        const icon = isFolder ? '📁' : this.getMimeTypeIcon(file.mimeType);
        const displayName = file.name.length > 50 
          ? file.name.substring(0, 47) + '...' 
          : file.name;
        
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
      .setCustomId(signComponentId({ kind: 'drive-nav-refresh', sid: sessionId }))
      .setLabel('🔄 Оновити')
      .setStyle(ButtonStyle.Primary);
    
    // Search button for real-time search
    const searchButton = new ButtonBuilder()
      .setCustomId(signComponentId({ kind: 'drive-nav-search', sid: sessionId }))
      .setLabel('🔍 Пошук')
      .setStyle(ButtonStyle.Primary);
    
    buttonRow.addComponents(refreshButton, searchButton);
    
    // Add parent folder button if not at root
    if (state.parentId) {
      const upButton = new ButtonBuilder()
        .setCustomId(signComponentId({ kind: 'drive-nav-up', sid: sessionId }))
        .setLabel('⬆️ Вгору')
        .setStyle(ButtonStyle.Secondary);
      
      buttonRow.addComponents(upButton);
    }
    
    // Add pagination buttons if needed
    if (result.nextPageToken) {
      const nextButton = new ButtonBuilder()
        .setCustomId(signComponentId({ kind: 'drive-nav-next', sid: sessionId }))
        .setLabel('➡️ Далі')
        .setStyle(ButtonStyle.Secondary);
      
      buttonRow.addComponents(nextButton);
    }
    
    if (components.length < 5) { // Discord limit is 5 action rows
      components.push(buttonRow);
    }
    
    return components;
  }

  // New method to get file type category filters
  private getFileTypeCategoryFilters(categories: string[]): string[] {
    const filters: string[] = [];
    
    for (const category of categories) {
      switch (category) {
        case 'documents':
          filters.push(
            "mimeType = 'application/vnd.google-apps.document'",
            "mimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'",
            "mimeType = 'application/msword'",
            "mimeType = 'application/pdf'",
            "mimeType = 'text/plain'"
          );
          break;
        case 'spreadsheets':
          filters.push(
            "mimeType = 'application/vnd.google-apps.spreadsheet'",
            "mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'",
            "mimeType = 'application/vnd.ms-excel'"
          );
          break;
        case 'presentations':
          filters.push(
            "mimeType = 'application/vnd.google-apps.presentation'",
            "mimeType = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'",
            "mimeType = 'application/vnd.ms-powerpoint'"
          );
          break;
        case 'images':
          filters.push(
            "mimeType contains 'image/'"
          );
          break;
        case 'videos':
          filters.push(
            "mimeType contains 'video/'"
          );
          break;
        case 'audio':
          filters.push(
            "mimeType contains 'audio/'"
          );
          break;
        case 'archives':
          filters.push(
            "mimeType = 'application/zip'",
            "mimeType = 'application/x-rar-compressed'",
            "mimeType = 'application/x-7z-compressed'",
            "mimeType = 'application/gzip'"
          );
          break;
      }
    }
    
    return filters;
  }
}