/**
 * Розширена система обробки файлів для Discord AI Assistant Bot
 * Безпечна робота з файлами та документами
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */
export interface FileInfo {
    name: string;
    path: string;
    size: number;
    extension: string;
    mimeType: string;
    lastModified: Date;
    isReadable: boolean;
    isWritable: boolean;
    isValid: boolean;
    errors: string[];
    warnings: string[];
}
export interface FileOperationResult {
    success: boolean;
    fileInfo?: FileInfo;
    content?: string | Buffer;
    error?: string;
    warnings: string[];
    duration: number;
    bytesProcessed: number;
}
export interface FileProcessorStats {
    totalOperations: number;
    successfulOperations: number;
    failedOperations: number;
    bytesProcessed: number;
    averageOperationTime: number;
    totalOperationTime: number;
    filesProcessed: number;
    cleanupOperations: number;
    lastOperation?: {
        type: string;
        filename: string;
        duration: number;
        success: boolean;
    };
}
export declare class FileProcessor {
    private static instance;
    private stats;
    private activeOperations;
    private cleanupInterval;
    private _isInitialized;
    constructor();
    /**
     * Ініціалізація обробника файлів
     */
    private initialize;
    /**
     * Створення необхідних директорій
     */
    private ensureDirectories;
    /**
     * Запуск періодичного очищення
     */
    private startCleanupInterval;
    /**
     * Безпечне читання файлу
     */
    readFile(filePath: string): Promise<FileOperationResult>;
    /**
     * Безпечне записування файлу
     */
    writeFile(filePath: string, content: string | Buffer, options?: {
        backup?: boolean;
        validate?: boolean;
    }): Promise<FileOperationResult>;
    /**
     * Валідація файлу
     */
    private validateFile;
    /**
     * Створення інформації про файл
     */
    private createFileInfo;
    /**
     * Читання вмісту файлу
     */
    private readFileContent;
    /**
     * Читання файлу по частинах
     */
    private readFileInChunks;
    /**
     * Запис вмісту файлу
     */
    private writeFileContent;
    /**
     * Створення резервної копії
     */
    private createBackup;
    /**
     * Визначення MIME типу
     */
    private getMimeType;
    /**
     * Генерація ID операції
     */
    private generateOperationId;
    /**
     * Очищення тимчасових файлів
     */
    private cleanupTempFiles;
    /**
     * Оновлення статистики
     */
    private updateStats;
    /**
     * Отримання статистики
     */
    getStats(): FileProcessorStats;
    /**
     * Очищення ресурсів
     */
    cleanup(): void;
    /**
     * Перевірка стану ініціалізації
     */
    isInitialized(): boolean;
}
export declare const fileProcessor: FileProcessor;
export declare const readFile: (filePath: string) => Promise<FileOperationResult>;
export declare const writeFile: (filePath: string, content: string | Buffer, options?: {
    backup?: boolean;
    validate?: boolean;
}) => Promise<FileOperationResult>;
export declare const getFileProcessorStats: () => FileProcessorStats;
export declare const cleanupFileProcessor: () => void;
export default fileProcessor;
//# sourceMappingURL=fileProcessor.d.ts.map