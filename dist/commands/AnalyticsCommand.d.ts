/**
 * 📊 Команди аналітики та звітності для ЗСУ
 * Спеціалізовані звіти та аналіз даних
 */
import type { BotConfig, CommandExecuteOptions } from '@/types';
import { BaseCommand } from './BaseCommand';
export declare class AnalyticsCommand extends BaseCommand {
    constructor(config: BotConfig);
    /**
     * Виконання команди
     */
    protected onExecute(options: CommandExecuteOptions): Promise<void>;
    /**
     * Обробка генерації звітів
     */
    private handleReport;
    /**
     * Обробка статистики
     */
    private handleStatistics;
    /**
     * Обробка прогнозування
     */
    private handleForecast;
    /**
     * Обробка порівняльного аналізу
     */
    private handleComparison;
    /**
     * Отримання назви типу звіту
     */
    private getReportTypeName;
    /**
     * Отримання назви категорії
     */
    private getCategoryName;
    /**
     * Отримання назви типу прогнозу
     */
    private getForecastTypeName;
    /**
     * Отримання назви об'єкта
     */
    private getObjectName;
    /**
     * Отримання назви метрики
     */
    private getMetricName;
}
//# sourceMappingURL=AnalyticsCommand.d.ts.map