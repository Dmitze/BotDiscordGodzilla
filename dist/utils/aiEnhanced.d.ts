/**
 * Розширений AI-модуль для Discord Bot
 * Включає природномовний інтерфейс, контекстну пам'ять та аналіз даних
 * TypeScript версія
 */
interface Message {
    role: 'user' | 'assistant';
    content: string;
    timestamp: number;
}
declare class AIEnhanced {
    private providers;
    private currentProvider;
    private stats;
    constructor();
    private createOpenAIProvider;
    private createOllamaProvider;
    getConversationContext(userId: string): Message[];
    saveToContext(userId: string, role: 'user' | 'assistant', content: string): void;
    analyzeNaturalLanguage(userInput: string): Promise<any>;
    generateResponse(prompt: string, options?: any): Promise<string>;
    analyzeData(data: any[], analysisType?: string): Promise<string>;
    generateReport(data: any[], options?: any): Promise<string>;
    processNaturalLanguageQuery(userId: string, userInput: string, sheetData?: any[] | null): Promise<string>;
    getHelpMessage(): string;
    clearContext(userId: string): void;
    getStats(): any;
    private updateStats;
}
export declare const aiEnhanced: AIEnhanced;
export default aiEnhanced;
//# sourceMappingURL=aiEnhanced.d.ts.map