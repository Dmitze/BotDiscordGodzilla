"use strict";
/**
 * Оптимізована команда пошуку
 * Використовує Redis кешування, Connection Pool та пагінацію
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.SearchCommand = void 0;
const discord_js_1 = require("discord.js");
const BaseCommand_1 = require("./BaseCommand");
const logger_1 = __importDefault(require("@/utils/logger"));
const security_1 = require("@/utils/security");
// Константи для конфігурації пошуку
const SEARCH_CONFIG = {
    MAX_RESULTS: 50,
    DEFAULT_LIMIT: 20,
    CACHE_TTL: 300, // 5 хвилин
    MAX_QUERY_LENGTH: 200,
    MAX_DATE_RANGE: 365, // днів
    PAGINATION_TIMEOUT: 300000, // 5 хвилин
    MAX_FILTERED_RESULTS: 1000,
    SEARCH_TIMEOUT: 30000, // 30 секунд
};
class SearchCommand extends BaseCommand_1.BaseCommand {
    constructor(config) {
        super('пошук', '🔍 Гнучкий пошук по документах ЗСУ', config, {
            category: 'search',
            cooldown: 5000, // 5 секунд
            permissions: ['ViewChannel'],
            usage: '/пошук запит:текст [опції]',
            examples: [
                '/пошук запит:особовий склад тип_документа:накази',
                '/пошук запит:техніка дата_від:01.01.2024 дата_до:31.12.2024',
                '/пошук запит:зброя підрозділ:рота пріоритет:критичний',
            ],
        }, (builder) => {
            return builder
                .addStringOption((option) => option
                .setName('запит')
                .setDescription('Що шукати? (наприклад: "особовий склад", "техніка", "зброя")')
                .setRequired(true)
                .setMaxLength(SEARCH_CONFIG.MAX_QUERY_LENGTH))
                .addStringOption((option) => option
                .setName('тип_документа')
                .setDescription('Тип документа для пошуку')
                .addChoices({ name: 'Всі документи', value: 'all' }, { name: 'Накази', value: 'orders' }, { name: 'Доповіді', value: 'reports' }, { name: 'Звіти', value: 'statistics' }, { name: 'Плани', value: 'plans' }, { name: 'Інструкції', value: 'instructions' }, { name: 'Протоколи', value: 'protocols' }, { name: 'Картки', value: 'cards' }, { name: 'Журнали', value: 'journals' }))
                .addStringOption((option) => option
                .setName('дата_від')
                .setDescription('Дата від (формат: ДД.ММ.РРРР)')
                .setMaxLength(10))
                .addStringOption((option) => option
                .setName('дата_до')
                .setDescription('Дата до (формат: ДД.ММ.РРРР)')
                .setMaxLength(10))
                .addStringOption((option) => option
                .setName('підрозділ')
                .setDescription('Підрозділ для пошуку')
                .setMaxLength(100))
                .addStringOption((option) => option
                .setName('пріоритет')
                .setDescription('Пріоритет документа')
                .addChoices({ name: 'Всі', value: 'all' }, { name: 'Критичний', value: 'critical' }, { name: 'Високий', value: 'high' }, { name: 'Середній', value: 'medium' }, { name: 'Низький', value: 'low' }))
                .addIntegerOption((option) => option
                .setName('ліміт')
                .setDescription(`Кількість результатів (макс. ${SEARCH_CONFIG.MAX_RESULTS})`)
                .setMinValue(1)
                .setMaxValue(SEARCH_CONFIG.MAX_RESULTS));
        });
        this.paginationStates = new Map();
        this.searchCache = new Map();
        this.searchStats = {
            totalSearches: 0,
            cacheHits: 0,
            cacheMisses: 0,
            averageSearchTime: 0,
            totalSearchTime: 0,
            errors: 0,
        };
    }
    /**
     * Виконання команди з детальним логуванням
     */
    async onExecute(options) {
        const { interaction } = options;
        const startTime = performance.now();
        try {
            // Валідація та отримання параметрів пошуку
            const searchParams = await this.extractAndValidateParams(interaction);
            // Відкладена відповідь
            await interaction.deferReply();
            // Логування початку пошуку
            logger_1.default.info('Початок пошуку', {
                user: interaction.user.tag,
                query: searchParams.query,
                filters: searchParams,
            });
            // Виконання пошуку
            const searchResult = await this.performSearchWithCache(searchParams, interaction.user.id);
            // Форматування результатів
            const formattedResults = this.formatResults(searchResult.rows, searchResult.headers);
            // Створення embed
            const embed = this.createSearchEmbed(searchResult, formattedResults);
            // Створення кнопок пагінації
            const components = this.createPaginationComponents(searchResult, 1);
            // Відправка відповіді
            await interaction.editReply({ embeds: [embed], components });
            // Оновлення статистики
            const duration = performance.now() - startTime;
            this.updateSearchStats(true, duration, searchResult.cacheHit);
            // Логування успішного завершення
            logger_1.default.info('Пошук успішно завершено', {
                user: interaction.user.tag,
                duration: `${duration.toFixed(2)}ms`,
                results: searchResult.filteredCount,
                cacheHit: searchResult.cacheHit,
            });
        }
        catch (error) {
            const duration = performance.now() - startTime;
            this.updateSearchStats(false, duration, false);
            logger_1.default.error('Помилка пошуку', {
                user: interaction.user.tag,
                error: error instanceof Error ? error.message : String(error),
                duration: `${duration.toFixed(2)}ms`,
            });
            await this.handleSearchError(interaction, error);
        }
    }
    /**
     * Витяг та валідація параметрів
     */
    async extractAndValidateParams(interaction) {
        const query = interaction.options.getString('запит', true);
        const documentType = interaction.options.getString('тип_документа') || 'all';
        const dateFrom = interaction.options.getString('дата_від');
        const dateTo = interaction.options.getString('дата_до');
        const unit = interaction.options.getString('підрозділ');
        const priority = interaction.options.getString('пріоритет') || 'all';
        const limit = interaction.options.getInteger('ліміт') || SEARCH_CONFIG.DEFAULT_LIMIT;
        // Валідація запиту
        const sanitizedQuery = (0, security_1.sanitizeInput)(query, 'search');
        if (!sanitizedQuery.isValid) {
            throw new Error(`Некорректний запит: ${sanitizedQuery.errors.join(', ')}`);
        }
        // Валідація дат
        if (dateFrom && !this.isValidDate(dateFrom)) {
            throw new Error('Некорректний формат дати "від" (використовуйте ДД.ММ.РРРР)');
        }
        if (dateTo && !this.isValidDate(dateTo)) {
            throw new Error('Некорректний формат дати "до" (використовуйте ДД.ММ.РРРР)');
        }
        // Перевірка діапазону дат
        if (dateFrom && dateTo) {
            const fromDate = this.parseDate(dateFrom);
            const toDate = this.parseDate(dateTo);
            if (fromDate && toDate && toDate < fromDate) {
                throw new Error('Дата "до" не може бути раніше дати "від"');
            }
        }
        // Валідація підрозділу
        if (unit) {
            const sanitizedUnit = (0, security_1.sanitizeInput)(unit, 'search');
            if (!sanitizedUnit.isValid) {
                throw new Error(`Некорректний підрозділ: ${sanitizedUnit.errors.join(', ')}`);
            }
        }
        return {
            query: sanitizedQuery.sanitizedValue || query,
            documentType,
            dateFrom,
            dateTo,
            unit: unit ? (0, security_1.sanitizeInput)(unit, 'search').sanitizedValue : undefined,
            priority,
            limit,
        };
    }
    /**
     * Виконання пошуку з кешуванням
     */
    async performSearchWithCache(searchParams, userId) {
        const cacheKey = this.generateCacheKey(searchParams);
        // Перевірка кешу
        const cached = this.searchCache.get(cacheKey);
        if (cached && Date.now() - cached.timestamp < SEARCH_CONFIG.CACHE_TTL * 1000) {
            this.searchStats.cacheHits++;
            logger_1.default.debug('Результат знайдено в кеші', { cacheKey });
            return { ...cached.result, cacheHit: true };
        }
        this.searchStats.cacheMisses++;
        // Виконання пошуку
        const searchResult = await this.performSearch(searchParams);
        // Кешування результату
        this.searchCache.set(cacheKey, {
            result: searchResult,
            timestamp: Date.now(),
        });
        // Обмеження розміру кешу
        if (this.searchCache.size > 100) {
            const oldestKey = this.searchCache.keys().next().value;
            this.searchCache.delete(oldestKey);
        }
        return { ...searchResult, cacheHit: false };
    }
    /**
     * Виконання пошуку
     */
    async performSearch(searchParams) {
        const startTime = performance.now();
        try {
            // Отримання сервісів
            const googleService = this.config?.google;
            if (!googleService) {
                throw new Error('Google сервіс не налаштовано');
            }
            // Отримання даних з Google Sheets
            const sheetData = await this.getSheetDataWithTimeout(googleService);
            if (!sheetData || !sheetData.values || sheetData.values.length === 0) {
                throw new Error('Немає даних для пошуку');
            }
            const headers = sheetData.values[0];
            const rows = sheetData.values.slice(1);
            // Фільтрація даних
            const filteredRows = this.filterData(rows, headers, searchParams);
            const searchTime = performance.now() - startTime;
            return {
                rows: filteredRows.slice(0, searchParams.limit),
                headers,
                totalCount: rows.length,
                filteredCount: filteredRows.length,
                searchTime,
                cacheHit: false,
                query: searchParams.query,
                filters: searchParams,
            };
        }
        catch (error) {
            const searchTime = performance.now() - startTime;
            logger_1.default.error('Помилка виконання пошуку', {
                error: error instanceof Error ? error.message : String(error),
                searchTime: `${searchTime.toFixed(2)}ms`,
            });
            throw error;
        }
    }
    /**
     * Отримання даних з таймаутом
     */
    async getSheetDataWithTimeout(googleService) {
        return Promise.race([
            googleService.getSheetData(this.config.google.spreadsheetId, 'A:Z', { useCache: true, cacheTTL: SEARCH_CONFIG.CACHE_TTL }),
            new Promise((_, reject) => setTimeout(() => reject(new Error('Таймаут отримання даних')), SEARCH_CONFIG.SEARCH_TIMEOUT)),
        ]);
    }
    /**
     * Фільтрація даних з оптимізацією
     */
    filterData(rows, headers, searchParams) {
        const startTime = performance.now();
        try {
            const filteredRows = rows.filter(row => {
                // Перевірка запиту
                if (!this.matchesQuery(row, headers, searchParams.query)) {
                    return false;
                }
                // Перевірка типу документа
                if (searchParams.documentType !== 'all' &&
                    !this.matchesDocumentType(row, headers, searchParams.documentType)) {
                    return false;
                }
                // Перевірка діапазону дат
                if (searchParams.dateFrom || searchParams.dateTo) {
                    if (!this.matchesDateRange(row, headers, searchParams.dateFrom, searchParams.dateTo)) {
                        return false;
                    }
                }
                // Перевірка підрозділу
                if (searchParams.unit && !this.matchesUnit(row, headers, searchParams.unit)) {
                    return false;
                }
                // Перевірка пріоритету
                if (searchParams.priority !== 'all' &&
                    !this.matchesPriority(row, headers, searchParams.priority)) {
                    return false;
                }
                return true;
            });
            const filterTime = performance.now() - startTime;
            logger_1.default.debug('Фільтрація завершена', {
                totalRows: rows.length,
                filteredRows: filteredRows.length,
                filterTime: `${filterTime.toFixed(2)}ms`,
            });
            // Обмеження кількості результатів
            if (filteredRows.length > SEARCH_CONFIG.MAX_FILTERED_RESULTS) {
                logger_1.default.warn('Кількість результатів обмежена', {
                    maxResults: SEARCH_CONFIG.MAX_FILTERED_RESULTS,
                    actualResults: filteredRows.length,
                });
                return filteredRows.slice(0, SEARCH_CONFIG.MAX_FILTERED_RESULTS);
            }
            return filteredRows;
        }
        catch (error) {
            logger_1.default.error('Помилка фільтрації даних', error);
            throw error;
        }
    }
    /**
     * Перевірка відповідності запиту з оптимізацією
     */
    matchesQuery(row, headers, query) {
        const searchTerms = query.toLowerCase().split(' ').filter(term => term.length > 0);
        if (searchTerms.length === 0)
            return true;
        return row.some((cell, index) => {
            const cellValue = cell.toLowerCase();
            return searchTerms.some(term => cellValue.includes(term));
        });
    }
    /**
     * Перевірка типу документа
     */
    matchesDocumentType(row, headers, documentType) {
        const typeIndex = headers.findIndex(h => h.toLowerCase().includes('тип'));
        if (typeIndex === -1)
            return true;
        const rowType = row[typeIndex]?.toLowerCase() || '';
        return rowType.includes(documentType.toLowerCase());
    }
    /**
     * Перевірка діапазону дат
     */
    matchesDateRange(row, headers, dateFrom, dateTo) {
        const dateIndex = headers.findIndex(h => h.toLowerCase().includes('дата'));
        if (dateIndex === -1)
            return true;
        const rowDate = this.parseDate(row[dateIndex]);
        if (!rowDate)
            return true;
        if (dateFrom) {
            const fromDate = this.parseDate(dateFrom);
            if (fromDate && rowDate < fromDate)
                return false;
        }
        if (dateTo) {
            const toDate = this.parseDate(dateTo);
            if (toDate && rowDate > toDate)
                return false;
        }
        return true;
    }
    /**
     * Перевірка підрозділу
     */
    matchesUnit(row, headers, unit) {
        const unitIndex = headers.findIndex(h => h.toLowerCase().includes('підрозділ'));
        if (unitIndex === -1)
            return true;
        const rowUnit = row[unitIndex]?.toLowerCase() || '';
        return rowUnit.includes(unit.toLowerCase());
    }
    /**
     * Перевірка пріоритету
     */
    matchesPriority(row, headers, priority) {
        const priorityIndex = headers.findIndex(h => h.toLowerCase().includes('пріоритет'));
        if (priorityIndex === -1)
            return true;
        const rowPriority = row[priorityIndex]?.toLowerCase() || '';
        return rowPriority.includes(priority.toLowerCase());
    }
    /**
     * Валідація дати
     */
    isValidDate(dateString) {
        const parsed = this.parseDate(dateString);
        return parsed !== null;
    }
    /**
     * Парсинг дати з покращеною обробкою помилок
     */
    parseDate(dateString) {
        if (!dateString || typeof dateString !== 'string')
            return null;
        try {
            // Спробувати різні формати дати
            const formats = [
                /(\d{1,2})\.(\d{1,2})\.(\d{4})/, // ДД.ММ.РРРР
                /(\d{4})-(\d{1,2})-(\d{1,2})/, // РРРР-ММ-ДД
                /(\d{1,2})\/(\d{1,2})\/(\d{4})/, // ДД/ММ/РРРР
            ];
            for (const format of formats) {
                const match = dateString.match(format);
                if (match) {
                    const [, day, month, year] = match;
                    const date = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
                    // Перевірка валідності дати
                    if (date.getFullYear() === parseInt(year) &&
                        date.getMonth() === parseInt(month) - 1 &&
                        date.getDate() === parseInt(day)) {
                        return date;
                    }
                }
            }
            return null;
        }
        catch (error) {
            logger_1.default.error('Помилка парсингу дати', { dateString, error });
            return null;
        }
    }
    /**
     * Форматування результатів з оптимізацією
     */
    formatResults(rows, headers) {
        try {
            return rows.map((row, index) => {
                const formattedRow = row.map((cell, cellIndex) => {
                    const header = headers[cellIndex] || `Колонка ${cellIndex + 1}`;
                    const cellValue = cell || 'Н/Д';
                    return `${header}: ${cellValue}`;
                });
                return `**${index + 1}.** ${formattedRow.slice(0, 3).join(' | ')}`;
            });
        }
        catch (error) {
            logger_1.default.error('Помилка форматування результатів', error);
            return ['Помилка форматування результатів'];
        }
    }
    /**
     * Створення embed для результатів пошуку
     */
    createSearchEmbed(searchResult, formattedResults) {
        const embed = new discord_js_1.EmbedBuilder()
            .setColor('#4CAF50')
            .setTitle('🔍 Результати пошуку')
            .setDescription(`**Запит:** ${searchResult.query}`)
            .addFields({
            name: '📊 Статистика',
            value: `Знайдено: **${searchResult.totalCount}**\nПісля фільтрації: **${searchResult.filteredCount}**`,
            inline: true
        }, {
            name: '📄 Тип документа',
            value: this.getDocumentTypeName(searchResult.filters.documentType),
            inline: true
        }, {
            name: '⚡ Швидкість',
            value: `${searchResult.searchTime.toFixed(2)}ms${searchResult.cacheHit ? ' (кеш)' : ''}`,
            inline: true
        })
            .setTimestamp();
        // Додавання результатів
        if (formattedResults.length > 0) {
            const resultsText = formattedResults.slice(0, 10).join('\n');
            embed.addFields({
                name: `📋 Результати (${formattedResults.length})`,
                value: resultsText.length > 1024 ? resultsText.substring(0, 1021) + '...' : resultsText
            });
        }
        else {
            embed.addFields({ name: '📋 Результати', value: 'Нічого не знайдено' });
        }
        return embed;
    }
    /**
     * Створення кнопок пагінації
     */
    createPaginationComponents(searchResult, currentPage) {
        const totalPages = Math.ceil(searchResult.filteredCount / SEARCH_CONFIG.DEFAULT_LIMIT);
        if (totalPages <= 1)
            return [];
        const row = new discord_js_1.ActionRowBuilder()
            .addComponents(new discord_js_1.ButtonBuilder()
            .setCustomId(`search_prev_${currentPage}`)
            .setLabel('◀️ Попередня')
            .setStyle(discord_js_1.ButtonStyle.Secondary)
            .setDisabled(currentPage <= 1), new discord_js_1.ButtonBuilder()
            .setCustomId(`search_next_${currentPage}`)
            .setLabel('Наступна ▶️')
            .setStyle(discord_js_1.ButtonStyle.Secondary)
            .setDisabled(currentPage >= totalPages), new discord_js_1.ButtonBuilder()
            .setCustomId(`search_close`)
            .setLabel('❌ Закрити')
            .setStyle(discord_js_1.ButtonStyle.Danger));
        return [row];
    }
    /**
     * Генерація ключа кешу
     */
    generateCacheKey(params) {
        const sortedParams = Object.keys(params)
            .sort()
            .map(key => `${key}:${params[key]}`)
            .join('|');
        return `search:${Buffer.from(sortedParams).toString('base64')}`;
    }
    /**
     * Отримання назви типу документа
     */
    getDocumentTypeName(type) {
        const typeNames = {
            'all': 'Всі документи',
            'orders': 'Накази',
            'reports': 'Доповіді',
            'statistics': 'Звіти',
            'plans': 'Плани',
            'instructions': 'Інструкції',
            'protocols': 'Протоколи',
            'cards': 'Картки',
            'journals': 'Журнали',
        };
        return typeNames[type] || type;
    }
    /**
     * Оновлення статистики пошуку
     */
    updateSearchStats(success, duration, cacheHit) {
        this.searchStats.totalSearches++;
        this.searchStats.totalSearchTime += duration;
        this.searchStats.averageSearchTime = this.searchStats.totalSearchTime / this.searchStats.totalSearches;
        if (!success) {
            this.searchStats.errors++;
        }
    }
    /**
     * Обробка помилки пошуку
     */
    async handleSearchError(interaction, error) {
        const errorMessage = error instanceof Error ? error.message : 'Невідома помилка';
        const errorEmbed = new discord_js_1.EmbedBuilder()
            .setColor('#FF6B6B')
            .setTitle('❌ Помилка пошуку')
            .setDescription(`**Помилка:** ${errorMessage}`)
            .addFields({ name: '💡 Порада', value: 'Перевірте правильність запиту та спробуйте ще раз' }, { name: '📞 Підтримка', value: 'Якщо проблема повторюється, зверніться до адміністратора' })
            .setTimestamp();
        try {
            if (interaction.deferred || interaction.replied) {
                await interaction.editReply({ embeds: [errorEmbed] });
            }
            else {
                await interaction.reply({ embeds: [errorEmbed], ephemeral: true });
            }
        }
        catch (replyError) {
            logger_1.default.error('Помилка відправки повідомлення про помилку пошуку', replyError);
        }
    }
    /**
     * Отримання статистики пошуку
     */
    getSearchStats() {
        return {
            ...this.searchStats,
            cacheSize: this.searchCache.size,
            paginationStates: this.paginationStates.size,
        };
    }
    /**
     * Очищення застарілих даних
     */
    cleanupExpiredData() {
        const now = Date.now();
        let cleanedCache = 0;
        let cleanedPagination = 0;
        // Очищення кешу
        for (const [key, cached] of this.searchCache.entries()) {
            if (now - cached.timestamp > SEARCH_CONFIG.CACHE_TTL * 1000) {
                this.searchCache.delete(key);
                cleanedCache++;
            }
        }
        // Очищення пагінації
        for (const [userId, state] of this.paginationStates.entries()) {
            if (now - state.timestamp > SEARCH_CONFIG.PAGINATION_TIMEOUT) {
                this.paginationStates.delete(userId);
                cleanedPagination++;
            }
        }
        if (cleanedCache > 0 || cleanedPagination > 0) {
            logger_1.default.debug('Очищено застарілі дані пошуку', {
                cache: cleanedCache,
                pagination: cleanedPagination,
            });
        }
    }
    /**
     * Завершення роботи
     */
    async shutdown() {
        await super.shutdown();
        this.searchCache.clear();
        this.paginationStates.clear();
        logger_1.default.info('Команда пошуку зупинена');
    }
}
exports.SearchCommand = SearchCommand;
//# sourceMappingURL=SearchCommand.js.map