/**
 * Оптимізована команда пошуку
 * Використовує Redis кешування, Connection Pool та пагінацію
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */
import type { BotConfig, CommandExecuteOptions } from '@/types';
import { BaseCommand } from './BaseCommand';
export declare class SearchCommand extends BaseCommand {
    private paginationStates;
    private searchCache;
    private searchStats;
    constructor(config: BotConfig);
    /**
     * Виконання команди з детальним логуванням
     */
    protected onExecute(options: CommandExecuteOptions): Promise<void>;
    /**
     * Витяг та валідація параметрів
     */
    private extractAndValidateParams;
    /**
     * Виконання пошуку з кешуванням
     */
    private performSearchWithCache;
    /**
     * Виконання пошуку
     */
    private performSearch;
    /**
     * Отримання даних з таймаутом
     */
    private getSheetDataWithTimeout;
    /**
     * Фільтрація даних з оптимізацією
     */
    private filterData;
    /**
     * Перевірка відповідності запиту з оптимізацією
     */
    private matchesQuery;
    /**
     * Перевірка типу документа
     */
    private matchesDocumentType;
    /**
     * Перевірка діапазону дат
     */
    private matchesDateRange;
    /**
     * Перевірка підрозділу
     */
    private matchesUnit;
    /**
     * Перевірка пріоритету
     */
    private matchesPriority;
    /**
     * Валідація дати
     */
    private isValidDate;
    /**
     * Парсинг дати з покращеною обробкою помилок
     */
    private parseDate;
    /**
     * Форматування результатів з оптимізацією
     */
    private formatResults;
    /**
     * Створення embed для результатів пошуку
     */
    private createSearchEmbed;
    /**
     * Створення кнопок пагінації
     */
    private createPaginationComponents;
    /**
     * Генерація ключа кешу
     */
    private generateCacheKey;
    /**
     * Отримання назви типу документа
     */
    private getDocumentTypeName;
    /**
     * Оновлення статистики пошуку
     */
    private updateSearchStats;
    /**
     * Обробка помилки пошуку
     */
    private handleSearchError;
    /**
     * Отримання статистики пошуку
     */
    getSearchStats(): any;
    /**
     * Очищення застарілих даних
     */
    cleanupExpiredData(): void;
    /**
     * Завершення роботи
     */
    shutdown(): Promise<void>;
}
//# sourceMappingURL=SearchCommand.d.ts.map