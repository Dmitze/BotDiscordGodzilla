/**
 * Розширена система обробки файлів для Discord AI Assistant Bot
 * Безпечна робота з файлами та документами
 * Версія 3.0.0 - Повністю рефакторовано з детальним логуванням
 */

import type { LogMeta } from '@/types';
import { existsSync, mkdirSync, statSync } from 'fs';
import {
  access as fsAccess,
  constants as fsConstants,
  readFile as fsReadFile,
  unlink as fsUnlink,
  writeFile as fsWriteFile,
} from 'fs/promises';

import { basename, dirname, extname, join } from 'path';
import { handleError } from './errorHandler';
import logger from './logger';
import { validateInput, sanitizeInput } from './security';

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
} as const;

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

export class FileProcessor {
  private static instance: FileProcessor | null = null;
  private stats!: FileProcessorStats;
  private activeOperations = new Set<string>();
  private cleanupInterval: NodeJS.Timeout | null = null;
  private _isInitialized = false;

  constructor() {
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
  private initialize(): void {
    try {
      logger.info('📁 Ініціалізація FileProcessor...');

      // Створення необхідних директорій
      this.ensureDirectories();

      // Запуск періодичного очищення
      this.startCleanupInterval();

      this._isInitialized = true;
      logger.info('✅ FileProcessor успішно ініціалізовано');
    } catch (error) {
      handleError(error, {
        serviceName: 'FileProcessor',
        additionalContext: { operation: 'initialize' },
      });
      throw new Error('Помилка ініціалізації FileProcessor');
    }
  }

  /**
   * Створення необхідних директорій
   */
  private ensureDirectories(): void {
    try {
      const directories = [FILE_PROCESSOR_CONSTANTS.TEMP_DIR, FILE_PROCESSOR_CONSTANTS.BACKUP_DIR];

      for (const dir of directories) {
        if (!existsSync(dir)) {
          mkdirSync(dir, { recursive: true });
          logger.debug(`📁 Створено директорію: ${dir}`);
        }
      }
    } catch (error) {
      handleError(error, {
        serviceName: 'FileProcessor',
        additionalContext: { operation: 'ensureDirectories' },
      });
    }
  }

  /**
   * Запуск періодичного очищення
   */
  private startCleanupInterval(): void {
    this.cleanupInterval = setInterval(() => {
      this.cleanupTempFiles();
    }, FILE_PROCESSOR_CONSTANTS.CLEANUP_INTERVAL);

    logger.info('⏰ Періодичне очищення файлів запущено');
  }

  /**
   * Безпечне читання файлу
   */
  public async readFile(filePath: string): Promise<FileOperationResult> {
    const operationId = this.generateOperationId('read', filePath);
    const startTime = performance.now();

    try {
      // Перевірка обмежень
      if (this.activeOperations.size >= FILE_PROCESSOR_CONSTANTS.MAX_CONCURRENT_OPERATIONS) {
        throw new Error('Досягнуто ліміт одночасних операцій');
      }

      this.activeOperations.add(operationId);

      logger.debug('📖 Початок читання файлу...', {
        filePath,
        operationId,
      } as LogMeta);

      // Валідація файлу
      const fileInfo = await this.validateFile(filePath);
      if (!fileInfo.isValid) {
        throw new Error(`Файл не валідний: ${fileInfo.errors.join(', ')}`);
      }

      // Читання файлу
      const content = await this.readFileContent(filePath, fileInfo.size);

      const duration = performance.now() - startTime;
      const result: FileOperationResult = {
        success: true,
        fileInfo,
        content,
        warnings: fileInfo.warnings,
        duration,
        bytesProcessed: fileInfo.size,
      };

      this.updateStats(true, duration, fileInfo.size);
      this.stats.filesProcessed++;

      logger.info('✅ Файл успішно прочитано', {
        filePath,
        size: fileInfo.size,
        duration: `${duration.toFixed(2)}ms`,
        operationId,
      } as LogMeta);

      return result;
    } catch (error) {
      const duration = performance.now() - startTime;
      this.updateStats(false, duration, 0);

      const errorMessage = error instanceof Error ? error.message : 'Невідома помилка';

      logger.error('❌ Помилка читання файлу', {
        filePath,
        error: errorMessage,
        duration: `${duration.toFixed(2)}ms`,
        operationId,
      } as LogMeta);

      return {
        success: false,
        error: errorMessage,
        warnings: [],
        duration,
        bytesProcessed: 0,
      };
    } finally {
      this.activeOperations.delete(operationId);
    }
  }

  /**
   * Безпечне записування файлу
   */
  public async writeFile(
    filePath: string,
    content: string | Buffer,
    options: { backup?: boolean; validate?: boolean } = {}
  ): Promise<FileOperationResult> {
    const operationId = this.generateOperationId('write', filePath);
    const startTime = performance.now();

    try {
      if (this.activeOperations.size >= FILE_PROCESSOR_CONSTANTS.MAX_CONCURRENT_OPERATIONS) {
        throw new Error('Досягнуто ліміт одночасних операцій');
      }

      this.activeOperations.add(operationId);

      logger.debug('📝 Початок запису файлу...', {
        filePath,
        contentSize: content.length,
        operationId,
      } as LogMeta);

      // Валідація вмісту
      if (options.validate) {
        const validation = validateInput(content.toString(), { inputType: 'file' });
        if (!validation.isValid) {
          throw new Error(`Невалідний вміст: ${validation.errors.join(', ')}`);
        }
      }

      // Створення резервної копії
      if (options.backup && existsSync(filePath)) {
        await this.createBackup(filePath);
      }

      // Запис файлу
      await this.writeFileContent(filePath, content);

      // Валідація записаного файлу
      const fileInfo = await this.validateFile(filePath);

      const duration = performance.now() - startTime;
      const result: FileOperationResult = {
        success: true,
        fileInfo,
        warnings: fileInfo.warnings,
        duration,
        bytesProcessed: content.length,
      };

      this.updateStats(true, duration, content.length);
      this.stats.filesProcessed++;

      logger.info('✅ Файл успішно записано', {
        filePath,
        size: content.length,
        duration: `${duration.toFixed(2)}ms`,
        operationId,
      } as LogMeta);

      return result;
    } catch (error) {
      const duration = performance.now() - startTime;
      this.updateStats(false, duration, 0);

      const errorMessage = error instanceof Error ? error.message : 'Невідома помилка';

      logger.error('❌ Помилка запису файлу', {
        filePath,
        error: errorMessage,
        duration: `${duration.toFixed(2)}ms`,
        operationId,
      } as LogMeta);

      return {
        success: false,
        error: errorMessage,
        warnings: [],
        duration,
        bytesProcessed: 0,
      };
    } finally {
      this.activeOperations.delete(operationId);
    }
  }

  /**
   * Валідація файлу
   */
  private async validateFile(filePath: string): Promise<FileInfo> {
    const errors: string[] = [];
    const warnings: string[] = [];

    try {
      // Перевірка існування
      if (!existsSync(filePath)) {
        errors.push('Файл не існує');
        return this.createFileInfo(filePath, errors, warnings);
      }

      // Отримання статистики файлу
      const stats = statSync(filePath);
      const extension = extname(filePath).toLowerCase();
      const name = basename(filePath);

      // Перевірка розміру
      if (stats.size > FILE_PROCESSOR_CONSTANTS.MAX_FILE_SIZE) {
        errors.push(
          `Файл занадто великий (${stats.size} байт, максимум ${FILE_PROCESSOR_CONSTANTS.MAX_FILE_SIZE})`
        );
      }

      // Перевірка імені файлу
      if (name.length > FILE_PROCESSOR_CONSTANTS.MAX_FILENAME_LENGTH) {
        errors.push(
          `Ім'я файлу занадто довге (${name.length} символів, максимум ${FILE_PROCESSOR_CONSTANTS.MAX_FILENAME_LENGTH})`
        );
      }

      // Перевірка розширення
      const allowedExts = FILE_PROCESSOR_CONSTANTS.ALLOWED_EXTENSIONS as readonly string[];
      if (!allowedExts.includes(extension)) {
        warnings.push(`Недозволене розширення файлу: ${extension}`);
      }

      // Перевірка прав доступу
      try {
        await fsAccess(filePath, fsConstants.R_OK);
      } catch {
        errors.push('Файл недоступний для читання');
      }

      try {
        await fsAccess(filePath, fsConstants.W_OK);
      } catch {
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
    } catch (error) {
      errors.push(
        `Помилка валідації: ${error instanceof Error ? error.message : 'Невідома помилка'}`
      );
      return this.createFileInfo(filePath, errors, warnings);
    }
  }

  /**
   * Створення інформації про файл
   */
  private createFileInfo(filePath: string, errors: string[], warnings: string[]): FileInfo {
    return {
      name: basename(filePath),
      path: filePath,
      size: 0,
      extension: extname(filePath).toLowerCase(),
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
  private async readFileContent(filePath: string, fileSize: number): Promise<string | Buffer> {
    if (fileSize > FILE_PROCESSOR_CONSTANTS.CHUNK_SIZE) {
      // Читання великих файлів по частинах
      return this.readFileInChunks(filePath);
    } else {
      // Читання малих файлів повністю
      return await fsReadFile(filePath, 'utf8');
    }
  }

  /**
   * Читання файлу по частинах
   */
  private async readFileInChunks(filePath: string): Promise<string> {
    const chunks: string[] = [];
    const fileHandle = await import('fs/promises').then(fs => fs.open(filePath, 'r'));

    try {
      const buffer = Buffer.alloc(FILE_PROCESSOR_CONSTANTS.CHUNK_SIZE);
      let bytesRead: number;

      while ((bytesRead = (await fileHandle.read(buffer, 0, buffer.length)).bytesRead) > 0) {
        chunks.push(buffer.toString('utf8', 0, bytesRead));
      }

      return chunks.join('');
    } finally {
      await fileHandle.close();
    }
  }

  /**
   * Запис вмісту файлу
   */
  private async writeFileContent(filePath: string, content: string | Buffer): Promise<void> {
    // Створення директорії якщо не існує
    const dir = dirname(filePath);
    if (!existsSync(dir)) {
      mkdirSync(dir, { recursive: true });
    }

    await fsWriteFile(filePath, content);
  }

  /**
   * Створення резервної копії
   */
  private async createBackup(filePath: string): Promise<void> {
    try {
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      const backupName = `${basename(filePath)}.backup.${timestamp}`;
      const backupPath = join(FILE_PROCESSOR_CONSTANTS.BACKUP_DIR, backupName);

      const content = await fsReadFile(filePath);
      await fsWriteFile(backupPath, content);

      logger.debug('💾 Створено резервну копію', {
        original: filePath,
        backup: backupPath,
      } as LogMeta);
    } catch (error) {
      handleError(error, {
        serviceName: 'FileProcessor',
        additionalContext: { operation: 'createBackup', filePath },
      });
    }
  }

  /**
   * Визначення MIME типу
   */
  private getMimeType(extension: string): string {
    const mimeTypes: Record<string, string> = {
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
  private generateOperationId(type: string, filePath: string): string {
    const timestamp = Date.now();
    const hash = require('crypto')
      .createHash('md5')
      .update(`${type}:${filePath}:${timestamp}`)
      .digest('hex');
    return `${type}_${hash.substring(0, 8)}`;
  }

  /**
   * Очищення тимчасових файлів
   */
  private async cleanupTempFiles(): Promise<void> {
    try {
      const tempDir = FILE_PROCESSOR_CONSTANTS.TEMP_DIR;
      if (!existsSync(tempDir)) return;

      const fs = require('fs/promises');
      const files = await fs.readdir(tempDir);
      const now = Date.now();
      let cleanedCount = 0;

      for (const file of files) {
        const filePath = join(tempDir, file);
        const stats = statSync(filePath);
        const age = now - stats.mtime.getTime();

        if (age > FILE_PROCESSOR_CONSTANTS.MAX_TEMP_AGE) {
          try {
            await fsUnlink(filePath);
            cleanedCount++;
          } catch (error) {
            logger.warn('⚠️ Не вдалося видалити тимчасовий файл', {
              filePath,
              error: error instanceof Error ? error.message : 'Невідома помилка',
            } as LogMeta);
          }
        }
      }

      if (cleanedCount > 0) {
        this.stats.cleanupOperations++;
        logger.info(`🧹 Очищено ${cleanedCount} тимчасових файлів`);
      }
    } catch (error) {
      handleError(error, {
        serviceName: 'FileProcessor',
        additionalContext: { operation: 'cleanupTempFiles' },
      });
    }
  }

  /**
   * Оновлення статистики
   */
  private updateStats(success: boolean, duration: number, bytesProcessed: number): void {
    try {
      this.stats.totalOperations++;
      this.stats.totalOperationTime += duration;
      this.stats.averageOperationTime = this.stats.totalOperationTime / this.stats.totalOperations;
      this.stats.bytesProcessed += bytesProcessed;

      if (success) {
        this.stats.successfulOperations++;
      } else {
        this.stats.failedOperations++;
      }
    } catch (error) {
      handleError(error, {
        serviceName: 'FileProcessor',
        additionalContext: { operation: 'updateStats' },
      });
    }
  }

  /**
   * Отримання статистики
   */
  public getStats(): FileProcessorStats {
    return { ...this.stats };
  }

  /**
   * Очищення ресурсів
   */
  public cleanup(): void {
    try {
      if (this.cleanupInterval) {
        clearInterval(this.cleanupInterval);
        this.cleanupInterval = null;
      }

      this.activeOperations.clear();

      logger.info('🧹 Ресурси FileProcessor очищено');
    } catch (error) {
      handleError(error, {
        serviceName: 'FileProcessor',
        additionalContext: { operation: 'cleanup' },
      });
    }
  }

  /**
   * Перевірка стану ініціалізації
   */
  public isInitialized(): boolean {
    return this._isInitialized;
  }
}

// Експорт єдиного екземпляра
export const fileProcessor = new FileProcessor();

// Експорт функцій для зручності
export const readFile = (filePath: string) => fileProcessor.readFile(filePath);
export const writeFile = (
  filePath: string,
  content: string | Buffer,
  options?: { backup?: boolean; validate?: boolean }
) => fileProcessor.writeFile(filePath, content, options);
export const getFileProcessorStats = () => fileProcessor.getStats();
export const cleanupFileProcessor = () => fileProcessor.cleanup();
export default fileProcessor;

/**
 * Нормализация текста: переносы в \n, табы → 2 пробела, схлопывание пробелов, удаление невидимых символов
 */
export function normalizeText(input: string): string {
  const s = String(input ?? '');
  // Приводим переносы к \n
  let out = s.replace(/\r\n?|\u2028|\u2029/g, '\n');
  // Табуляции → два пробела
  out = out.replace(/\t/g, '  ');
  // Удаляем Zero-Width и прочие невидимые
  out = out.replace(/[\u200B-\u200D\uFEFF]/g, '');
  // Схлопываем более 2 пустых строк подряд
  out = out.replace(/\n{3,}/g, '\n\n');
  // Ограничиваем повторяющиеся пробелы
  out = out.replace(/ {3,}/g, '  ');
  return out;
}

/**
 * Санитизация безопасного вывода в чат Discord с ограничением длины
 * По умолчанию ~1800 символов, чтобы не упираться в лимиты и оставить место под служебный текст
 */
export function sanitizeTextForChat(input: string, maxLen = 1800): string {
  const cleaned = sanitizeInput(normalizeText(String(input ?? '')));
  if (cleaned.length <= maxLen) return cleaned.trim();
  // Обрезаем по границе абзаца/предложения, если возможно
  const slice = cleaned.slice(0, maxLen);
  const lastBreak = Math.max(slice.lastIndexOf('\n\n'), slice.lastIndexOf('\n'), slice.lastIndexOf('. '));
  const base = lastBreak > 300 ? slice.slice(0, lastBreak + 1) : slice;
  return base.trimEnd() + '\n…';
}
