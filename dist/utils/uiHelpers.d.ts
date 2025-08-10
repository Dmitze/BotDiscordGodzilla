/**
 * Модуль для покращеного UI/UX
 * Включає красиві embed повідомлення, інтерактивні кнопки та прогрес-бари
 * TypeScript версія
 */
import { EmbedBuilder, ActionRowBuilder, ButtonBuilder, ButtonStyle } from 'discord.js';
declare const COLORS: {
    readonly SUCCESS: 65280;
    readonly ERROR: 16711680;
    readonly WARNING: 16753920;
    readonly INFO: 39423;
    readonly AI: 10181046;
    readonly SEARCH: 3447003;
    readonly FILES: 15105570;
    readonly EXPORT: 2600544;
};
declare const EMOJIS: {
    readonly SUCCESS: "✅";
    readonly ERROR: "❌";
    readonly WARNING: "⚠️";
    readonly INFO: "ℹ️";
    readonly AI: "🤖";
    readonly SEARCH: "🔍";
    readonly FILES: "📁";
    readonly EXPORT: "📤";
    readonly LOADING: "⏳";
    readonly DONE: "🎉";
    readonly HELP: "❓";
    readonly SETTINGS: "⚙️";
    readonly STATS: "📊";
    readonly SECURITY: "🔒";
};
interface ActionButton {
    id: string;
    label: string;
    style?: ButtonStyle;
    emoji?: string;
    disabled?: boolean;
}
interface MenuOption {
    id: string;
    label: string;
    description?: string;
}
interface BotStats {
    totalCommands?: number;
    uniqueUsers?: number;
    activeConversations?: number;
    commandStats?: Record<string, number>;
    aiStats?: {
        requests?: number;
        provider?: string;
        avgResponseTime?: number;
    };
}
/**
 * Клас для створення покращених UI елементів
 */
declare class UIHelper {
    /**
     * Створення базового embed
     */
    static createBaseEmbed(title: string, description: string, color?: number): EmbedBuilder;
    /**
     * Створення embed для результатів пошуку
     */
    static createSearchResultsEmbed(results: any[], query: string, page?: number, totalPages?: number): EmbedBuilder;
    /**
     * Форматування результату пошуку
     */
    static formatSearchResult(result: any): string;
    /**
     * Створення embed для AI відповіді
     */
    static createAIResponseEmbed(query: string, response: string, confidence?: number): EmbedBuilder;
    /**
     * Створення embed для роботи з файлами
     */
    static createFileEmbed(action: string, fileName: string, content?: string | null, metadata?: Record<string, any> | null): EmbedBuilder;
    /**
     * Створення embed для експорту
     */
    static createExportEmbed(format: string, recordCount: number, fileName: string): EmbedBuilder;
    /**
     * Створення embed для помилок
     */
    static createErrorEmbed(error: Error | string, context?: string): EmbedBuilder;
    /**
     * Створення embed для успіху
     */
    static createSuccessEmbed(message: string, details?: string | null): EmbedBuilder;
    /**
     * Створення кнопок для навігації
     */
    static createNavigationButtons(currentPage: number, totalPages: number, customIds?: Record<string, string>): ActionRowBuilder<ButtonBuilder>;
    /**
     * Створення кнопок для дій
     */
    static createActionButtons(actions: ActionButton[]): ActionRowBuilder<ButtonBuilder>;
    /**
     * Створення прогрес-бару
     */
    static createProgressBar(current: number, total: number, width?: number): string;
    /**
     * Створення embed з прогрес-баром
     */
    static createProgressEmbed(title: string, current: number, total: number, status?: string): EmbedBuilder;
    /**
     * Створення embed для довідки
     */
    static createHelpEmbed(category?: string): EmbedBuilder;
    /**
     * Створення embed для статистики
     */
    static createStatsEmbed(stats: BotStats): EmbedBuilder;
    /**
     * Створення embed для безпеки
     */
    static createSecurityEmbed(event: string, details: Record<string, any>): EmbedBuilder;
    /**
     * Створення інтерактивного меню
     */
    static createInteractiveMenu(title: string, options: MenuOption[], description?: string): {
        embed: EmbedBuilder;
        row: ActionRowBuilder<ButtonBuilder>;
    };
    /**
     * Обробка інтерактивних компонентів
     */
    static handleInteraction(interaction: any, timeout?: number): Promise<any>;
}
export { UIHelper, COLORS, EMOJIS };
//# sourceMappingURL=uiHelpers.d.ts.map