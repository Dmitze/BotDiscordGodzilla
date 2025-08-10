/**
 * Утиліта для пагінації великих даних
 * Оптимізована для роботи з Discord embeds та великими наборами даних
 * TypeScript версія
 */
import { EmbedBuilder, ActionRowBuilder, ButtonBuilder } from 'discord.js';
interface PaginationOptions {
    itemsPerPage?: number;
    maxPages?: number;
    embedColor?: number;
    title?: string;
    description?: string;
    fields?: string[];
    footer?: string;
    timestamp?: Date;
}
interface PaginationStats {
    totalItems: number;
    totalPages: number;
    currentPage: number;
    itemsPerPage: number;
    hasNext: boolean;
    hasPrevious: boolean;
}
declare class Pagination {
    private data;
    private currentPage;
    private itemsPerPage;
    private maxPages;
    private embedColor;
    private title;
    private description;
    private fields;
    private footer;
    private timestamp;
    private totalItems;
    private totalPages;
    constructor(data: any[], options?: PaginationOptions);
    /**
     * Отримання поточної сторінки
     */
    getCurrentPage(): number;
    /**
     * Отримання загальної кількості сторінок
     */
    getTotalPages(): number;
    /**
     * Отримання загальної кількості елементів
     */
    getTotalItems(): number;
    /**
     * Перевірка чи можна перейти на попередню сторінку
     */
    hasPreviousPage(): boolean;
    /**
     * Перевірка чи можна перейти на наступну сторінку
     */
    hasNextPage(): boolean;
    /**
     * Перехід на попередню сторінку
     */
    previousPage(): boolean;
    /**
     * Перехід на наступну сторінку
     */
    nextPage(): boolean;
    /**
     * Перехід на конкретну сторінку
     */
    goToPage(page: number): boolean;
    /**
     * Отримання даних поточної сторінки
     */
    getCurrentPageData(): any[];
    /**
     * Створення Discord embed для поточної сторінки
     */
    createEmbed(): EmbedBuilder;
    /**
     * Створення кнопок навігації
     */
    createNavigationButtons(): ActionRowBuilder<ButtonBuilder>;
    /**
     * Обробка взаємодії з кнопками
     */
    handleButtonInteraction(customId: string): boolean;
    /**
     * Форматування назви поля
     */
    private formatFieldName;
    /**
     * Форматування значення поля
     */
    private formatFieldValue;
    /**
     * Форматування об'єкта
     */
    private formatObjectValue;
    /**
     * Форматування значення
     */
    private formatValue;
    /**
     * Створення тексту footer
     */
    private createFooterText;
    /**
     * Обрізання тексту
     */
    private truncateText;
    /**
     * Капіталізація першої літери
     */
    private capitalizeFirst;
    /**
     * Отримання статистики пагінації
     */
    getStats(): PaginationStats;
    /**
     * Створення пагінації з фільтром
     */
    static createWithFilter(data: any[], filterFn: (item: any) => boolean, options?: PaginationOptions): Pagination;
    /**
     * Створення пагінації з сортуванням
     */
    static createWithSort(data: any[], sortFn: (a: any, b: any) => number, options?: PaginationOptions): Pagination;
    /**
     * Створення пагінації з лімітом
     */
    static createWithLimit(data: any[], limit: number, options?: PaginationOptions): Pagination;
    /**
     * Створення пагінації для пошуку
     */
    static createForSearch(data: any[], searchTerm: string, searchFields?: string[], options?: PaginationOptions): Pagination;
}
export default Pagination;
export { Pagination };
//# sourceMappingURL=pagination.d.ts.map