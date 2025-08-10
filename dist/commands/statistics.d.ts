/**
 * Команда для роботи зі статистикою та складними формулами Google Sheets
 * Підтримує підрахунок по парних/непарних стовпцях, агрегацію по аркушах
 * TypeScript версія 3.0.0
 */
import { SlashCommandBuilder, CommandInteraction } from 'discord.js';
import type { BaseCommand } from '@/types';
declare class StatisticsCommand implements BaseCommand {
    readonly name = "statistics";
    readonly description = "\u041E\u0442\u0440\u0438\u043C\u0430\u043D\u043D\u044F \u0441\u0442\u0430\u0442\u0438\u0441\u0442\u0438\u043A\u0438 \u0437 Google Sheets \u0437 \u043F\u0456\u0434\u0442\u0440\u0438\u043C\u043A\u043E\u044E \u0441\u043A\u043B\u0430\u0434\u043D\u0438\u0445 \u0444\u043E\u0440\u043C\u0443\u043B";
    readonly usage = "/statistics <\u043E\u043F\u0435\u0440\u0430\u0446\u0456\u044F> <\u0430\u0440\u043A\u0443\u0448\u0456> [\u043E\u043F\u0446\u0456\u0457]";
    private readonly googleService;
    private readonly aiService;
    constructor();
    /**
     * Створення команди
     */
    getCommandData(): SlashCommandBuilder;
    /**
     * Виконання команди
     */
    execute(interaction: CommandInteraction): Promise<void>;
    /**
     * Витягування опцій з interaction
     */
    private extractOptions;
    /**
     * Схема валідації
     */
    private getValidationSchema;
    /**
     * Отримання статистики
     */
    private getStatistics;
    /**
     * Розрахунок статистики по парних/непарних стовпцях
     */
    private calculateColumnStatistics;
    /**
     * Виконання складних формул
     */
    private executeComplexFormula;
    /**
     * Розрахунок базової статистики
     */
    private calculateBasicStatistics;
    /**
     * Отримання індексу стовпця
     */
    private getColumnIndex;
    /**
     * Генерація підсумку
     */
    private generateSummary;
    /**
     * Створення embed для відповіді
     */
    private createStatisticsEmbed;
    /**
     * Створення кнопок дій
     */
    private createActionButtons;
    /**
     * Обробка помилок
     */
    private handleError;
    /**
     * Отримання назви команди
     */
    getName(): string;
    /**
     * Отримання опису команди
     */
    getDescription(): string;
}
export default StatisticsCommand;
//# sourceMappingURL=statistics.d.ts.map