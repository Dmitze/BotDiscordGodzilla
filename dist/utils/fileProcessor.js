"use strict";
/**
 * Розширена система обробки файлів для Discord AI Assistant Bot
 * Безпечна робота з файлами та документами
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.cleanupFileProcessor = exports.getFileProcessorStats = exports.writeFile = exports.readFile = exports.fileProcessor = exports.FileProcessor = void 0;
const fs_1 = require("fs");
const promises_1 = require("fs/promises");
const path_1 = require("path");
const errorHandler_1 = require("./errorHandler");
const logger_1 = __importDefault(require("./logger"));
const security_1 = require("./security");
// Константи для обробки файлів
const FILE_PROCESSOR_CONSTANTS = {
    MAX_FILE_SIZE: 50 * 1024 * 1024, // 50MB
    MAX_FILENAME_LENGTH: 255,
    ALLOWED_EXTENSIONS: ['.txt', '.md', '.json', '.csv', '.xlsx', '.xls', '.pdf', '.doc', '.docx'],
    ALLOWED_MIME_TYPES: [
        'text/plain',
        'text/markdown',
        'application/json',
        'text/csv',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel',
        'application/pdf',
        'application/msword',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    ],
    TEMP_DIR: 'data/tmp',
    BACKUP_DIR: 'data/backup',
    CLEANUP_INTERVAL: 24 * 60 * 60 * 1000, // 24 години
    MAX_TEMP_AGE: 7 * 24 * 60 * 60 * 1000, // 7 днів
    CHUNK_SIZE: 1024 * 1024, // 1MB
    MAX_CONCURRENT_OPERATIONS: 5,
};
class FileProcessor {
    constructor() {
        this.activeOperations = new Set();
        this.cleanupInterval = null;
        this._isInitialized = false;
        if (FileProcessor.instance) {
            return FileProcessor.instance;
        }
        FileProcessor.instance = this;
        this.stats = {
            totalOperations: 0,
            successfulOperations: 0,
            failedOperations: 0,
            bytesProcessed: 0,
            averageOperationTime: 0,
            totalOperationTime: 0,
            filesProcessed: 0,
            cleanupOperations: 0,
        };
        this.initialize();
    }
    /**
     * Ініціалізація обробника файлів
     */
    initialize() {
        try {
            logger_1.default.info('📁 Ініціалізація FileProcessor...');
            // Створення необхідних директорій
            this.ensureDirectories();
            // Запуск періодичного очищення
            this.startCleanupInterval();
            this._isInitialized = true;
            logger_1.default.info('✅ FileProcessor успішно ініціалізовано');
        }
        catch (error) {
            (0, errorHandler_1.handleError)(error, {
                serviceName: 'FileProcessor',
                additionalContext: { operation: 'initialize' },
            });
            throw new Error('Помилка ініціалізації FileProcessor');
        }
    }
    /**
     * Створення необхідних директорій
     */
    ensureDirectories() {
        try {
            const directories = [
                FILE_PROCESSOR_CONSTANTS.TEMP_DIR,
                FILE_PROCESSOR_CONSTANTS.BACKUP_DIR,
            ];
            for (const dir of directories) {
                if (!(0, fs_1.existsSync)(dir)) {
                    (0, fs_1.mkdirSync)(dir, { recursive: true });
                    logger_1.default.debug(`📁 Створено директорію: ${dir}`);
                }
            }
        }
        catch (error) {
            (0, errorHandler_1.handleError)(error, {
                serviceName: 'FileProcessor',
                additionalContext: { operation: 'ensureDirectories' },
            });
        }
    }
    /**
     * Запуск періодичного очищення
     */
    startCleanupInterval() {
        this.cleanupInterval = setInterval(() => {
            this.cleanupTempFiles();
        }, FILE_PROCESSOR_CONSTANTS.CLEANUP_INTERVAL);
        logger_1.default.info('⏰ Періодичне очищення файлів запущено');
    }
    /**
     * Безпечне читання файлу
     */
    async readFile(filePath) {
        const operationId = this.generateOperationId('read', filePath);
        const startTime = performance.now();
        try {
            // Перевірка обмежень
            if (this.activeOperations.size >= FILE_PROCESSOR_CONSTANTS.MAX_CONCURRENT_OPERATIONS) {
                throw new Error('Досягнуто ліміт одночасних операцій');
            }
            this.activeOperations.add(operationId);
            logger_1.default.debug('📖 Початок читання файлу...', {
                filePath,
                operationId,
            });
            // Валідація файлу
            const fileInfo = await this.validateFile(filePath);
            if (!fileInfo.isValid) {
                throw new Error(`Файл не валідний: ${fileInfo.errors.join(', ')}`);
            }
            // Читання файлу
            const content = await this.readFileContent(filePath, fileInfo.size);
            const duration = performance.now() - startTime;
            const result = {
                success: true,
                fileInfo,
                content,
                warnings: fileInfo.warnings,
                duration,
                bytesProcessed: fileInfo.size,
            };
            this.updateStats(true, duration, fileInfo.size);
            this.stats.filesProcessed++;
            logger_1.default.info('✅ Файл успішно прочитано', {
                filePath,
                size: fileInfo.size,
                duration: `${duration.toFixed(2)}ms`,
                operationId,
            });
            return result;
        }
        catch (error) {
            const duration = performance.now() - startTime;
            this.updateStats(false, duration, 0);
            const errorMessage = error instanceof Error ? error.message : 'Невідома помилка';
            logger_1.default.error('❌ Помилка читання файлу', {
                filePath,
                error: errorMessage,
                duration: `${duration.toFixed(2)}ms`,
                operationId,
            });
            return {
                success: false,
                error: errorMessage,
                warnings: [],
                duration,
                bytesProcessed: 0,
            };
        }
        finally {
            this.activeOperations.delete(operationId);
        }
    }
    /**
     * Безпечне записування файлу
     */
    async writeFile(filePath, content, options = {}) {
        const operationId = this.generateOperationId('write', filePath);
        const startTime = performance.now();
        try {
            if (this.activeOperations.size >= FILE_PROCESSOR_CONSTANTS.MAX_CONCURRENT_OPERATIONS) {
                throw new Error('Досягнуто ліміт одночасних операцій');
            }
            this.activeOperations.add(operationId);
            logger_1.default.debug('📝 Початок запису файлу...', {
                filePath,
                contentSize: content.length,
                operationId,
            });
            // Валідація вмісту
            if (options.validate) {
                const validation = (0, security_1.validateInput)(content.toString(), { inputType: 'file' });
                if (!validation.isValid) {
                    throw new Error(`Невалідний вміст: ${validation.errors.join(', ')}`);
                }
            }
            // Створення резервної копії
            if (options.backup && (0, fs_1.existsSync)(filePath)) {
                await this.createBackup(filePath);
            }
            // Запис файлу
            await this.writeFileContent(filePath, content);
            // Валідація записаного файлу
            const fileInfo = await this.validateFile(filePath);
            const duration = performance.now() - startTime;
            const result = {
                success: true,
                fileInfo,
                warnings: fileInfo.warnings,
                duration,
                bytesProcessed: content.length,
            };
            this.updateStats(true, duration, content.length);
            this.stats.filesProcessed++;
            logger_1.default.info('✅ Файл успішно записано', {
                filePath,
                size: content.length,
                duration: `${duration.toFixed(2)}ms`,
                operationId,
            });
            return result;
        }
        catch (error) {
            const duration = performance.now() - startTime;
            this.updateStats(false, duration, 0);
            const errorMessage = error instanceof Error ? error.message : 'Невідома помилка';
            logger_1.default.error('❌ Помилка запису файлу', {
                filePath,
                error: errorMessage,
                duration: `${duration.toFixed(2)}ms`,
                operationId,
            });
            return {
                success: false,
                error: errorMessage,
                warnings: [],
                duration,
                bytesProcessed: 0,
            };
        }
        finally {
            this.activeOperations.delete(operationId);
        }
    }
    /**
     * Валідація файлу
     */
    async validateFile(filePath) {
        const errors = [];
        const warnings = [];
        try {
            // Перевірка існування
            if (!(0, fs_1.existsSync)(filePath)) {
                errors.push('Файл не існує');
                return this.createFileInfo(filePath, errors, warnings);
            }
            // Отримання статистики файлу
            const stats = (0, fs_1.statSync)(filePath);
            const extension = (0, path_1.extname)(filePath).toLowerCase();
            const name = (0, path_1.basename)(filePath);
            // Перевірка розміру
            if (stats.size > FILE_PROCESSOR_CONSTANTS.MAX_FILE_SIZE) {
                errors.push(`Файл занадто великий (${stats.size} байт, максимум ${FILE_PROCESSOR_CONSTANTS.MAX_FILE_SIZE})`);
            }
            // Перевірка імені файлу
            if (name.length > FILE_PROCESSOR_CONSTANTS.MAX_FILENAME_LENGTH) {
                errors.push(`Ім'я файлу занадто довге (${name.length} символів, максимум ${FILE_PROCESSOR_CONSTANTS.MAX_FILENAME_LENGTH})`);
            }
            // Перевірка розширення
            const allowedExts = FILE_PROCESSOR_CONSTANTS.ALLOWED_EXTENSIONS;
            if (!allowedExts.includes(extension)) {
                warnings.push(`Недозволене розширення файлу: ${extension}`);
            }
            // Перевірка прав доступу
            try {
                await (0, promises_1.access)(filePath, promises_1.constants.R_OK);
            }
            catch {
                errors.push('Файл недоступний для читання');
            }
            try {
                await (0, promises_1.access)(filePath, promises_1.constants.W_OK);
            }
            catch {
                warnings.push('Файл недоступний для запису');
            }
            // Визначення MIME типу
            const mimeType = this.getMimeType(extension);
            return {
                name,
                path: filePath,
                size: stats.size,
                extension,
                mimeType,
                lastModified: stats.mtime,
                isReadable: errors.length === 0,
                isWritable: !warnings.some(w => w.includes('запису')),
                isValid: errors.length === 0,
                errors,
                warnings,
            };
        }
        catch (error) {
            errors.push(`Помилка валідації: ${error instanceof Error ? error.message : 'Невідома помилка'}`);
            return this.createFileInfo(filePath, errors, warnings);
        }
    }
    /**
     * Створення інформації про файл
     */
    createFileInfo(filePath, errors, warnings) {
        return {
            name: (0, path_1.basename)(filePath),
            path: filePath,
            size: 0,
            extension: (0, path_1.extname)(filePath).toLowerCase(),
            mimeType: 'unknown',
            lastModified: new Date(),
            isReadable: false,
            isWritable: false,
            isValid: errors.length === 0,
            errors,
            warnings,
        };
    }
    /**
     * Читання вмісту файлу
     */
    async readFileContent(filePath, fileSize) {
        if (fileSize > FILE_PROCESSOR_CONSTANTS.CHUNK_SIZE) {
            // Читання великих файлів по частинах
            return this.readFileInChunks(filePath);
        }
        else {
            // Читання малих файлів повністю
            return await (0, promises_1.readFile)(filePath, 'utf8');
        }
    }
    /**
     * Читання файлу по частинах
     */
    async readFileInChunks(filePath) {
        const chunks = [];
        const fileHandle = await Promise.resolve().then(() => __importStar(require('fs/promises'))).then(fs => fs.open(filePath, 'r'));
        try {
            const buffer = Buffer.alloc(FILE_PROCESSOR_CONSTANTS.CHUNK_SIZE);
            let bytesRead;
            while ((bytesRead = (await fileHandle.read(buffer, 0, buffer.length)).bytesRead) > 0) {
                chunks.push(buffer.toString('utf8', 0, bytesRead));
            }
            return chunks.join('');
        }
        finally {
            await fileHandle.close();
        }
    }
    /**
     * Запис вмісту файлу
     */
    async writeFileContent(filePath, content) {
        // Створення директорії якщо не існує
        const dir = (0, path_1.dirname)(filePath);
        if (!(0, fs_1.existsSync)(dir)) {
            (0, fs_1.mkdirSync)(dir, { recursive: true });
        }
        await (0, promises_1.writeFile)(filePath, content);
    }
    /**
     * Створення резервної копії
     */
    async createBackup(filePath) {
        try {
            const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
            const backupName = `${(0, path_1.basename)(filePath)}.backup.${timestamp}`;
            const backupPath = (0, path_1.join)(FILE_PROCESSOR_CONSTANTS.BACKUP_DIR, backupName);
            const content = await (0, promises_1.readFile)(filePath);
            await (0, promises_1.writeFile)(backupPath, content);
            logger_1.default.debug('💾 Створено резервну копію', {
                original: filePath,
                backup: backupPath,
            });
        }
        catch (error) {
            (0, errorHandler_1.handleError)(error, {
                serviceName: 'FileProcessor',
                additionalContext: { operation: 'createBackup', filePath },
            });
        }
    }
    /**
     * Визначення MIME типу
     */
    getMimeType(extension) {
        const mimeTypes = {
            '.txt': 'text/plain',
            '.md': 'text/markdown',
            '.json': 'application/json',
            '.csv': 'text/csv',
            '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            '.xls': 'application/vnd.ms-excel',
            '.pdf': 'application/pdf',
            '.doc': 'application/msword',
            '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        };
        return mimeTypes[extension] || 'application/octet-stream';
    }
    /**
     * Генерація ID операції
     */
    generateOperationId(type, filePath) {
        const timestamp = Date.now();
        const hash = require('crypto').createHash('md5').update(`${type}:${filePath}:${timestamp}`).digest('hex');
        return `${type}_${hash.substring(0, 8)}`;
    }
    /**
     * Очищення тимчасових файлів
     */
    async cleanupTempFiles() {
        try {
            const tempDir = FILE_PROCESSOR_CONSTANTS.TEMP_DIR;
            if (!(0, fs_1.existsSync)(tempDir))
                return;
            const fs = require('fs/promises');
            const files = await fs.readdir(tempDir);
            const now = Date.now();
            let cleanedCount = 0;
            for (const file of files) {
                const filePath = (0, path_1.join)(tempDir, file);
                const stats = (0, fs_1.statSync)(filePath);
                const age = now - stats.mtime.getTime();
                if (age > FILE_PROCESSOR_CONSTANTS.MAX_TEMP_AGE) {
                    try {
                        await (0, promises_1.unlink)(filePath);
                        cleanedCount++;
                    }
                    catch (error) {
                        logger_1.default.warn('⚠️ Не вдалося видалити тимчасовий файл', {
                            filePath,
                            error: error instanceof Error ? error.message : 'Невідома помилка',
                        });
                    }
                }
            }
            if (cleanedCount > 0) {
                this.stats.cleanupOperations++;
                logger_1.default.info(`🧹 Очищено ${cleanedCount} тимчасових файлів`);
            }
        }
        catch (error) {
            (0, errorHandler_1.handleError)(error, {
                serviceName: 'FileProcessor',
                additionalContext: { operation: 'cleanupTempFiles' },
            });
        }
    }
    /**
     * Оновлення статистики
     */
    updateStats(success, duration, bytesProcessed) {
        try {
            this.stats.totalOperations++;
            this.stats.totalOperationTime += duration;
            this.stats.averageOperationTime = this.stats.totalOperationTime / this.stats.totalOperations;
            this.stats.bytesProcessed += bytesProcessed;
            if (success) {
                this.stats.successfulOperations++;
            }
            else {
                this.stats.failedOperations++;
            }
        }
        catch (error) {
            (0, errorHandler_1.handleError)(error, {
                serviceName: 'FileProcessor',
                additionalContext: { operation: 'updateStats' },
            });
        }
    }
    /**
     * Отримання статистики
     */
    getStats() {
        return { ...this.stats };
    }
    /**
     * Очищення ресурсів
     */
    cleanup() {
        try {
            if (this.cleanupInterval) {
                clearInterval(this.cleanupInterval);
                this.cleanupInterval = null;
            }
            this.activeOperations.clear();
            logger_1.default.info('🧹 Ресурси FileProcessor очищено');
        }
        catch (error) {
            (0, errorHandler_1.handleError)(error, {
                serviceName: 'FileProcessor',
                additionalContext: { operation: 'cleanup' },
            });
        }
    }
    /**
     * Перевірка стану ініціалізації
     */
    isInitialized() {
        return this._isInitialized;
    }
}
exports.FileProcessor = FileProcessor;
FileProcessor.instance = null;
// Експорт єдиного екземпляра
exports.fileProcessor = new FileProcessor();
// Експорт функцій для зручності
const readFile = (filePath) => exports.fileProcessor.readFile(filePath);
exports.readFile = readFile;
const writeFile = (filePath, content, options) => exports.fileProcessor.writeFile(filePath, content, options);
exports.writeFile = writeFile;
const getFileProcessorStats = () => exports.fileProcessor.getStats();
exports.getFileProcessorStats = getFileProcessorStats;
const cleanupFileProcessor = () => exports.fileProcessor.cleanup();
exports.cleanupFileProcessor = cleanupFileProcessor;
exports.default = exports.fileProcessor;
//# sourceMappingURL=fileProcessor.js.map