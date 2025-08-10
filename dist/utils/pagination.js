"use strict";
/**
 * Утиліта для пагінації великих даних
 * Оптимізована для роботи з Discord embeds та великими наборами даних
 * TypeScript версія
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.Pagination = void 0;
const discord_js_1 = require("discord.js");
class Pagination {
    constructor(data, options = {}) {
        this.data = Array.isArray(data) ? data : [];
        this.currentPage = 0;
        this.itemsPerPage = options.itemsPerPage || 10;
        this.maxPages = options.maxPages || 50;
        this.embedColor = options.embedColor || 0x0099ff;
        this.title = options.title || 'Результати';
        this.description = options.description || '';
        this.fields = options.fields || [];
        this.footer = options.footer || '';
        this.timestamp = options.timestamp || new Date();
        this.totalItems = this.data.length;
        this.totalPages = Math.min(Math.ceil(this.totalItems / this.itemsPerPage), this.maxPages);
    }
    /**
     * Отримання поточної сторінки
     */
    getCurrentPage() {
        return this.currentPage;
    }
    /**
     * Отримання загальної кількості сторінок
     */
    getTotalPages() {
        return this.totalPages;
    }
    /**
     * Отримання загальної кількості елементів
     */
    getTotalItems() {
        return this.totalItems;
    }
    /**
     * Перевірка чи можна перейти на попередню сторінку
     */
    hasPreviousPage() {
        return this.currentPage > 0;
    }
    /**
     * Перевірка чи можна перейти на наступну сторінку
     */
    hasNextPage() {
        return this.currentPage < this.totalPages - 1;
    }
    /**
     * Перехід на попередню сторінку
     */
    previousPage() {
        if (this.hasPreviousPage()) {
            this.currentPage--;
            return true;
        }
        return false;
    }
    /**
     * Перехід на наступну сторінку
     */
    nextPage() {
        if (this.hasNextPage()) {
            this.currentPage++;
            return true;
        }
        return false;
    }
    /**
     * Перехід на конкретну сторінку
     */
    goToPage(page) {
        if (page >= 0 && page < this.totalPages) {
            this.currentPage = page;
            return true;
        }
        return false;
    }
    /**
     * Отримання даних поточної сторінки
     */
    getCurrentPageData() {
        const startIndex = this.currentPage * this.itemsPerPage;
        const endIndex = Math.min(startIndex + this.itemsPerPage, this.totalItems);
        return this.data.slice(startIndex, endIndex);
    }
    /**
     * Створення Discord embed для поточної сторінки
     */
    createEmbed() {
        const embed = new discord_js_1.EmbedBuilder()
            .setTitle(this.title)
            .setColor(this.embedColor)
            .setTimestamp(this.timestamp);
        if (this.description) {
            embed.setDescription(this.description);
        }
        // Додавання полів
        const pageData = this.getCurrentPageData();
        if (pageData.length > 0) {
            pageData.forEach((item, index) => {
                const fieldName = this.formatFieldName(item, index);
                const fieldValue = this.formatFieldValue(item, index);
                if (fieldName && fieldValue) {
                    embed.addFields({
                        name: fieldName,
                        value: fieldValue,
                        inline: true
                    });
                }
            });
        }
        else {
            embed.addFields({
                name: 'Немає даних',
                value: 'На цій сторінці немає даних для відображення',
                inline: false
            });
        }
        // Додавання footer
        const footerText = this.createFooterText();
        if (footerText) {
            embed.setFooter({ text: footerText });
        }
        return embed;
    }
    /**
     * Створення кнопок навігації
     */
    createNavigationButtons() {
        const row = new discord_js_1.ActionRowBuilder();
        // Кнопка "Перша сторінка"
        row.addComponents(new discord_js_1.ButtonBuilder()
            .setCustomId('pagination_first')
            .setLabel('⏮️')
            .setStyle(discord_js_1.ButtonStyle.Secondary)
            .setDisabled(this.currentPage === 0));
        // Кнопка "Попередня сторінка"
        row.addComponents(new discord_js_1.ButtonBuilder()
            .setCustomId('pagination_prev')
            .setLabel('◀️')
            .setStyle(discord_js_1.ButtonStyle.Primary)
            .setDisabled(!this.hasPreviousPage()));
        // Кнопка "Наступна сторінка"
        row.addComponents(new discord_js_1.ButtonBuilder()
            .setCustomId('pagination_next')
            .setLabel('▶️')
            .setStyle(discord_js_1.ButtonStyle.Primary)
            .setDisabled(!this.hasNextPage()));
        // Кнопка "Остання сторінка"
        row.addComponents(new discord_js_1.ButtonBuilder()
            .setCustomId('pagination_last')
            .setLabel('⏭️')
            .setStyle(discord_js_1.ButtonStyle.Secondary)
            .setDisabled(this.currentPage === this.totalPages - 1));
        return row;
    }
    /**
     * Обробка взаємодії з кнопками
     */
    handleButtonInteraction(customId) {
        switch (customId) {
            case 'pagination_first':
                return this.goToPage(0);
            case 'pagination_prev':
                return this.previousPage();
            case 'pagination_next':
                return this.nextPage();
            case 'pagination_last':
                return this.goToPage(this.totalPages - 1);
            default:
                return false;
        }
    }
    /**
     * Форматування назви поля
     */
    formatFieldName(item, index) {
        if (this.fields.length > 0) {
            const fieldIndex = index % this.fields.length;
            return this.fields[fieldIndex] || `Елемент ${index + 1}`;
        }
        return `Елемент ${index + 1}`;
    }
    /**
     * Форматування значення поля
     */
    formatFieldValue(item, index) {
        if (typeof item === 'string') {
            return this.truncateText(item, 100);
        }
        else if (typeof item === 'object' && item !== null) {
            return this.formatObjectValue(item);
        }
        else {
            return this.formatValue(item);
        }
    }
    /**
     * Форматування об'єкта
     */
    formatObjectValue(obj) {
        try {
            const entries = Object.entries(obj).slice(0, 3); // Беремо перші 3 поля
            const formatted = entries.map(([key, value]) => `${this.capitalizeFirst(key)}: ${this.formatValue(value)}`).join('\n');
            return this.truncateText(formatted, 100);
        }
        catch (error) {
            return 'Помилка форматування';
        }
    }
    /**
     * Форматування значення
     */
    formatValue(value) {
        if (value === null || value === undefined) {
            return '—';
        }
        else if (typeof value === 'string') {
            return value;
        }
        else if (typeof value === 'number') {
            return value.toString();
        }
        else if (typeof value === 'boolean') {
            return value ? 'Так' : 'Ні';
        }
        else if (Array.isArray(value)) {
            return value.slice(0, 3).join(', ') + (value.length > 3 ? '...' : '');
        }
        else {
            return String(value);
        }
    }
    /**
     * Створення тексту footer
     */
    createFooterText() {
        const parts = [];
        if (this.footer) {
            parts.push(this.footer);
        }
        parts.push(`Сторінка ${this.currentPage + 1} з ${this.totalPages}`);
        parts.push(`Всього елементів: ${this.totalItems}`);
        return parts.join(' • ');
    }
    /**
     * Обрізання тексту
     */
    truncateText(text, maxLength) {
        if (text.length <= maxLength)
            return text;
        return text.substring(0, maxLength - 3) + '...';
    }
    /**
     * Капіталізація першої літери
     */
    capitalizeFirst(str) {
        if (!str)
            return str;
        return str.charAt(0).toUpperCase() + str.slice(1);
    }
    /**
     * Отримання статистики пагінації
     */
    getStats() {
        return {
            totalItems: this.totalItems,
            totalPages: this.totalPages,
            currentPage: this.currentPage,
            itemsPerPage: this.itemsPerPage,
            hasNext: this.hasNextPage(),
            hasPrevious: this.hasPreviousPage(),
        };
    }
    /**
     * Створення пагінації з фільтром
     */
    static createWithFilter(data, filterFn, options = {}) {
        const filteredData = data.filter(filterFn);
        return new Pagination(filteredData, options);
    }
    /**
     * Створення пагінації з сортуванням
     */
    static createWithSort(data, sortFn, options = {}) {
        const sortedData = [...data].sort(sortFn);
        return new Pagination(sortedData, options);
    }
    /**
     * Створення пагінації з лімітом
     */
    static createWithLimit(data, limit, options = {}) {
        const limitedData = data.slice(0, limit);
        return new Pagination(limitedData, options);
    }
    /**
     * Створення пагінації для пошуку
     */
    static createForSearch(data, searchTerm, searchFields = [], options = {}) {
        if (!searchTerm) {
            return new Pagination(data, options);
        }
        const searchLower = searchTerm.toLowerCase();
        const filteredData = data.filter(item => {
            if (searchFields.length > 0) {
                return searchFields.some(field => {
                    const value = item[field];
                    return value && String(value).toLowerCase().includes(searchLower);
                });
            }
            else {
                return Object.values(item).some(value => value && String(value).toLowerCase().includes(searchLower));
            }
        });
        return new Pagination(filteredData, options);
    }
}
exports.Pagination = Pagination;
exports.default = Pagination;
//# sourceMappingURL=pagination.js.map