/**
 * 📄 Команди для роботи з військовими документами ЗСУ
 * Спеціалізовані функції для різних типів документів
 */
import type { BotConfig, CommandExecuteOptions } from '@/types';
import { BaseCommand } from './BaseCommand';
export declare class DocumentsCommand extends BaseCommand {
    constructor(config: BotConfig);
    /**
     * Виконання команди
     */
    protected onExecute(options: CommandExecuteOptions): Promise<void>;
    /**
     * Обробка особового складу
     */
    private handlePersonnel;
    /**
     * Обробка техніки
     */
    private handleEquipment;
    /**
     * Обробка матеріалів
     */
    private handleMaterials;
    /**
     * Обробка операцій
     */
    private handleOperations;
    /**
     * Обробка наказів
     */
    private handleOrders;
    /**
     * Отримання назви дії
     */
    private getActionName;
}
//# sourceMappingURL=DocumentsCommand.d.ts.map