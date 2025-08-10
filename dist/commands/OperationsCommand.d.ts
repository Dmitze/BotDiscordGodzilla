/**
 * ⚔️ Команди оперативного управління ЗСУ
 * Спеціалізовані функції для оперативної роботи
 */
import type { BotConfig, CommandExecuteOptions } from '@/types';
import { BaseCommand } from './BaseCommand';
export declare class OperationsCommand extends BaseCommand {
    constructor(config: BotConfig);
    /**
     * Виконання команди
     */
    protected onExecute(options: CommandExecuteOptions): Promise<void>;
    /**
     * Обробка оперативної ситуації
     */
    private handleSituation;
    /**
     * Обробка завдань
     */
    private handleTasks;
    /**
     * Обробка координації
     */
    private handleCoordination;
    /**
     * Обробка розвідки
     */
    private handleIntelligence;
    /**
     * Обробка зв'язку
     */
    private handleCommunications;
    /**
     * Отримання назви дії завдання
     */
    private getTaskActionName;
    /**
     * Отримання назви типу координації
     */
    private getCoordinationTypeName;
    /**
     * Отримання назви типу розвідки
     */
    private getIntelligenceTypeName;
    /**
     * Отримання назви дії зв'язку
     */
    private getCommunicationActionName;
}
//# sourceMappingURL=OperationsCommand.d.ts.map