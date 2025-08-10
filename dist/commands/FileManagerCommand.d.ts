/**
 * Команда для роботи з Google Drive та різними форматами файлів
 * Включає пошук, читання та аналіз файлів
 */
import type { BotConfig, CommandExecuteOptions } from '@/types';
import { BaseCommand } from './BaseCommand';
export declare class FileManagerCommand extends BaseCommand {
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
     * Витяг параметрів з interaction
     */
    private extractOptions;
    /**
     * Валідація параметрів
     */
    private validateOptions;
    /**
     * Обробка пошуку файлів
     */
    private handleSearch;
    /**
     * Обробка читання файлу
     */
    private handleRead;
    /**
     * Обробка аналізу файлу
     */
    private handleAnalyze;
    /**
     * Обробка створення звіту
     */
    private handleReport;
    /**
     * Відправка результату
     */
    private sendResult;
    /**
     * Отримання назви типу файлу
     */
    private getFileTypeName;
    /**
     * Отримання назви типу аналізу
     */
    private getAnalysisTypeName;
    /**
     * Отримання заголовку підкоманди
     */
    private getSubcommandTitle;
    /**
     * Логування події безпеки
     */
    private logSecurityEvent;
}
//# sourceMappingURL=FileManagerCommand.d.ts.map