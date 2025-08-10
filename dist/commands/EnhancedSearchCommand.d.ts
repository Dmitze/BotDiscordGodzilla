/**
 * 🔍 Покращений пошук з діапазонами та сортуванням
 * Розширені можливості пошуку та фільтрації даних
 */
import type { BotConfig, CommandExecuteOptions } from '@/types';
import { BaseCommand } from './BaseCommand';
export declare class EnhancedSearchCommand extends BaseCommand {
    constructor(config: BotConfig);
    /**
     * Виконання команди
     */
    protected onExecute(options: CommandExecuteOptions): Promise<void>;
    /**
     * Отримання даних з Google Sheets
     */
    private getSheetData;
    /**
     * Витягування фільтрів з interaction
     */
    private extractFilters;
    /**
     * Виконання пошуку з фільтрами
     */
    private performSearch;
    /**
     * Сортування результатів
     */
    private sortResults;
    /**
     * Створення embed з результатами
     */
    private createResultsEmbed;
    /**
     * Створення компонентів навігації
     */
    private createNavigationComponents;
    /**
     * Отримання активних фільтрів
     */
    private getActiveFilters;
    /**
     * Отримання індексу колонки
     */
    private getColumnIndex;
}
//# sourceMappingURL=EnhancedSearchCommand.d.ts.map