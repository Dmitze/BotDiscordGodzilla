/**
 * Команда для моніторингу продуктивності
 * Відстеження метрик та оптимізація системи
 */
import type { BotConfig, CommandExecuteOptions } from '@/types';
import { BaseCommand } from './BaseCommand';
export declare class PerformanceCommand extends BaseCommand {
    constructor(config: BotConfig);
    /**
     * Виконання команди
     */
    protected onExecute(options: CommandExecuteOptions): Promise<void>;
    /**
     * Показ загального статусу
     */
    private showGeneralStatus;
    /**
     * Показ статистики кешу
     */
    private showCacheStats;
    /**
     * Показ статистики черг
     */
    private showQueueStats;
    /**
     * Показ статистики API
     */
    private showApiStats;
    /**
     * Показ рекомендацій по оптимізації
     */
    private showOptimizationRecommendations;
}
//# sourceMappingURL=PerformanceCommand.d.ts.map