/**
 * AI-асистент команда з природномовним інтерфейсом
 * Використовує розширений AI-модуль та систему безпеки
 */
import type { BotConfig, CommandExecuteOptions } from '@/types';
import { BaseCommand } from './BaseCommand';
export declare class AIAssistantCommand extends BaseCommand {
    constructor(config: BotConfig);
    /**
     * Виконання команди
     */
    protected onExecute(options: CommandExecuteOptions): Promise<void>;
    /**
     * Перевірка прав доступу
     */
    private checkPermission;
    /**
     * Валідація параметрів команди
     */
    private validateCommandOptions;
    /**
     * Логування події безпеки
     */
    private logSecurityEvent;
    /**
     * Обробка AI запиту
     */
    private processAIQuery;
}
//# sourceMappingURL=AIAssistantCommand.d.ts.map