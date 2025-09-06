import { BaseService } from '@/core/BaseService';
import type { BotConfig } from '@/types';
import type { GoogleService } from '@/services/GoogleService';
import type { DriveFile } from '@/types/drive';
import type { Message, TextChannel } from 'discord.js';
import logger from '@/utils/logger';

export interface DocumentMention {
  fileId: string;
  fileName: string;
  mimeType: string;
  match: string;
  position: number;
}

// New interfaces for enhanced functionality
export interface DocumentAnnotation {
  id: string;
  fileId: string;
  userId: string;
  userName: string;
  content: string;
  timestamp: Date;
  position?: { start: number; end: number }; // For text documents
  page?: number; // For PDFs
  coordinates?: { x: number; y: number }; // For images
}

export interface CollaborationSession {
  id: string;
  fileId: string;
  fileName: string;
  participants: { userId: string; userName: string; joinedAt: Date }[];
  createdAt: Date;
  lastActivity: Date;
  isActive: boolean;
}

export interface RealTimeEdit {
  userId: string;
  userName: string;
  action: 'insert' | 'delete' | 'replace';
  content: string;
  position: number;
  timestamp: Date;
}

export class DocumentMentionHandler extends BaseService {
  private google: GoogleService | null = null;
  private documentCache: Map<string, DriveFile> = new Map();
  // New properties for enhanced functionality
  private annotations: Map<string, DocumentAnnotation[]> = new Map();
  private collaborationSessions: Map<string, CollaborationSession> = new Map();
  private realTimeEdits: Map<string, RealTimeEdit[]> = new Map();
  private activeCollaborations: Map<string, string> = new Map(); // userId -> sessionId
  private readonly CACHE_TTL = 5 * 60 * 1000; // 5 minutes
  private readonly MAX_ANNOTATIONS_PER_FILE = 100;
  private readonly MAX_EDITS_PER_FILE = 1000;

  constructor(config: BotConfig) {
    super('DocumentMentionHandler', config);
  }

  /**
   * Ініціалізує сервіс з необхідними залежностями
   */
  initializeServices(google: GoogleService): void {
    this.google = google;
  }

  /**
   * Обробляє згадки документів у повідомленні
   */
  async handleDocumentMentions(message: Message): Promise<void> {
    try {
      // Перевіряємо чи повідомлення не від бота
      if (message.author.bot) return;

      // Знаходимо згадки документів у повідомленні
      const mentions = this.findDocumentMentions(message.content);
      
      if (mentions.length === 0) return;

      // Обробляємо кожну згадку
      for (const mention of mentions) {
        await this.processDocumentMention(message, mention);
      }
    } catch (error) {
      logger.error('Помилка обробки згадок документів', {
        component: 'DocumentMentionHandler',
        messageId: message.id,
        error: error instanceof Error ? error.message : String(error)
      });
    }
  }

  /**
   * Знаходить згадки документів у тексті
   */
  private findDocumentMentions(text: string): DocumentMention[] {
    const mentions: DocumentMention[] = [];
    
    // Патерни для пошуку згадок документів
    const patterns = [
      // Назва файлу в лапках
      /"([^"]+?\.(?:docx?|xlsx?|pdf|txt|pptx?))"/gi,
      // Назва файлу з крапкою
      /(\S+\.(?:docx?|xlsx?|pdf|txt|pptx?))/gi,
      // Згадка з префіксом doc:
      /doc:([^\s]+)/gi
    ];
    
    for (const pattern of patterns) {
      let match;
      while ((match = pattern.exec(text)) !== null) {
        const fullMatch = match[0];
        const fileName = match[1];
        const position = match.index;
        
        mentions.push({
          fileId: '', // Буде заповнено пізніше
          fileName: fileName || '',
          mimeType: this.guessMimeType(fileName || ''),
          match: fullMatch,
          position
        });
      }
    }
    
    // Видаляємо дублікати
    return this.deduplicateMentions(mentions);
  }

  /**
   * Видаляє дублікати згадок
   */
  private deduplicateMentions(mentions: DocumentMention[]): DocumentMention[] {
    const uniqueMentions = new Map<string, DocumentMention>();
    
    for (const mention of mentions) {
      // Використовуємо назву файлу як ключ
      const key = mention.fileName.toLowerCase();
      
      // Якщо згадка ще не існує або нова згадка більш конкретна
      if (!uniqueMentions.has(key) || 
          mention.match.length > (uniqueMentions.get(key)?.match.length || 0)) {
        uniqueMentions.set(key, mention);
      }
    }
    
    return Array.from(uniqueMentions.values());
  }

  /**
   * Визначає MIME-тип за назвою файлу
   */
  private guessMimeType(fileName: string): string {
    const ext = fileName.toLowerCase().split('.').pop() || '';
    
    const mimeTypes: Record<string, string> = {
      'doc': 'application/msword',
      'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'pdf': 'application/pdf',
      'txt': 'text/plain',
      'xls': 'application/vnd.ms-excel',
      'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'ppt': 'application/vnd.ms-powerpoint',
      'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    };
    
    return mimeTypes[ext] || 'application/octet-stream';
  }

  /**
   * Обробляє окрему згадку документа
   */
  private async processDocumentMention(message: Message, mention: DocumentMention): Promise<void> {
    try {
      // Шукаємо файл у Google Drive
      const file = await this.findDocumentInDrive(mention.fileName);
      
      if (!file) {
        logger.debug('Документ не знайдено', {
          component: 'DocumentMentionHandler',
          fileName: mention.fileName
        });
        return;
      }
      
      // Оновлюємо інформацію про згадку
      mention.fileId = file.id;
      
      // Надсилаємо інформацію про документ
      await this.sendDocumentInfo(message, file, mention);
    } catch (error) {
      logger.error('Помилка обробки згадки документа', {
        component: 'DocumentMentionHandler',
        fileName: mention.fileName,
        error: error instanceof Error ? error.message : String(error)
      });
    }
  }

  /**
   * Шукає документ у Google Drive
   */
  private async findDocumentInDrive(fileName: string): Promise<DriveFile | null> {
    try {
      if (!this.google) {
        throw new Error('GoogleService не ініціалізовано');
      }

      // Перевіряємо кеш
      const cachedFile = this.getCachedFile(fileName);
      if (cachedFile) {
        return cachedFile;
      }

      // Шукаємо файл у Google Drive
      const folderId = this.config.drive?.folderId || 'root';
      
      const result = await this.google.listDriveFiles({
        folderId,
        query: `name contains '${fileName}'`,
        pageSize: 1
      });
      
      if (result.files.length > 0) {
        const file = result.files[0];
        
        // Зберігаємо у кеш
        this.cacheFile(fileName, file);
        
        return file;
      }
      
      return null;
    } catch (error) {
      logger.error('Помилка пошуку документа у Drive', {
        component: 'DocumentMentionHandler',
        fileName,
        error: error instanceof Error ? error.message : String(error)
      });
      return null;
    }
  }

  /**
   * Отримує файл з кешу
   */
  private getCachedFile(fileName: string): DriveFile | null {
    const cached = this.documentCache.get(fileName.toLowerCase());
    
    if (cached) {
      // Перевіряємо термін дії кешу
      const now = Date.now();
      const fileAge = now - (cached as any).__cachedAt;
      
      if (fileAge < this.CACHE_TTL) {
        return cached;
      } else {
        // Видаляємо прострочений кеш
        this.documentCache.delete(fileName.toLowerCase());
      }
    }
    
    return null;
  }

  /**
   * Зберігає файл у кеш
   */
  private cacheFile(fileName: string, file: DriveFile | null): void {
    if (!file) return;
    
    // Додаємо час кешування
    const fileWithTimestamp = {
      ...file,
      __cachedAt: Date.now()
    };
    
    this.documentCache.set(fileName.toLowerCase(), fileWithTimestamp as DriveFile);
    
    // Обмежуємо розмір кешу
    if (this.documentCache.size > 100) {
      const firstKey = this.documentCache.keys().next().value;
      if (firstKey) {
        this.documentCache.delete(firstKey);
      }
    }
  }

  /**
   * Надсилає інформацію про документ
   */
  private async sendDocumentInfo(message: Message, file: DriveFile, mention: DocumentMention): Promise<void> {
    try {
      // Створюємо посилання на документ
      const link = file.webViewLink || `https://drive.google.com/file/d/${file.id}/view`;
      
      // Отримуємо анотації для документа
      const fileAnnotations = this.annotations.get(file.id) || [];
      
      // Створюємо повідомлення з інформацією про документ
      let fileInfo = `📁 **${file.name}**\n` +
                    `📎 Тип: ${this.getMimeTypeLabel(file.mimeType || '')}\n` +
                    `🔗 [Відкрити в Google Drive](${link})`;
      
      // Додаємо інформацію про анотації
      if (fileAnnotations.length > 0) {
        fileInfo += `\n📝 Анотацій: ${fileAnnotations.length}`;
      }
      
      // Додаємо кнопки для додаткових дій
      fileInfo += `\n\nДії: \`/doc annotate ${file.id}\`, \`/doc collaborate ${file.id}\``;
      
      // Надсилаємо повідомлення
      await message.reply({
        content: fileInfo,
        allowedMentions: { repliedUser: false }
      });
      
      logger.debug('Надіслано інформацію про документ', {
        component: 'DocumentMentionHandler',
        fileId: file.id,
        fileName: file.name
      });
    } catch (error) {
      logger.error('Помилка надсилання інформації про документ', {
        component: 'DocumentMentionHandler',
        fileId: file.id,
        fileName: file.name,
        error: error instanceof Error ? error.message : String(error)
      });
    }
  }

  /**
   * Додає анотацію до документа
   */
  addAnnotation(
    fileId: string,
    userId: string,
    userName: string,
    content: string,
    position?: { start: number; end: number }
  ): DocumentAnnotation {
    try {
      const annotation: DocumentAnnotation = {
        id: this.generateId(),
        fileId,
        userId,
        userName,
        content,
        timestamp: new Date(),
        ...(position !== undefined && { position })
      };
      
      // Отримуємо існуючі анотації для файлу
      let fileAnnotations = this.annotations.get(fileId) || [];
      
      // Додаємо нову анотацію
      fileAnnotations.push(annotation);
      
      // Обмежуємо кількість анотацій
      if (fileAnnotations.length > this.MAX_ANNOTATIONS_PER_FILE) {
        fileAnnotations = fileAnnotations.slice(-this.MAX_ANNOTATIONS_PER_FILE);
      }
      
      // Зберігаємо оновлені анотації
      this.annotations.set(fileId, fileAnnotations);
      
      logger.info('Додано анотацію до документа', {
        component: 'DocumentMentionHandler',
        fileId,
        userId,
        annotationId: annotation.id
      });
      
      return annotation;
    } catch (error) {
      logger.error('Помилка додавання анотації до документа', {
        component: 'DocumentMentionHandler',
        fileId,
        userId,
        error: error instanceof Error ? error.message : String(error)
      });
      
      throw error;
    }
  }

  /**
   * Отримує анотації для документа
   */
  getDocumentAnnotations(fileId: string): DocumentAnnotation[] {
    return this.annotations.get(fileId) || [];
  }

  /**
   * Видаляє анотацію
   */
  removeAnnotation(fileId: string, annotationId: string, userId: string): boolean {
    try {
      const fileAnnotations = this.annotations.get(fileId);
      
      if (!fileAnnotations) {
        return false;
      }
      
      const initialLength = fileAnnotations.length;
      const filteredAnnotations = fileAnnotations.filter(a => 
        a.id !== annotationId || a.userId === userId
      );
      
      if (filteredAnnotations.length < initialLength) {
        this.annotations.set(fileId, filteredAnnotations);
        
        logger.info('Видалено анотацію з документа', {
          component: 'DocumentMentionHandler',
          fileId,
          annotationId,
          userId
        });
        
        return true;
      }
      
      return false;
    } catch (error) {
      logger.error('Помилка видалення анотації з документа', {
        component: 'DocumentMentionHandler',
        fileId,
        annotationId,
        userId,
        error: error instanceof Error ? error.message : String(error)
      });
      
      return false;
    }
  }

  /**
   * Створює сесію співпраці
   */
  createCollaborationSession(
    fileId: string,
    fileName: string,
    creatorId: string,
    creatorName: string
  ): CollaborationSession {
    try {
      const sessionId = this.generateId();
      
      const session: CollaborationSession = {
        id: sessionId,
        fileId,
        fileName,
        participants: [{
          userId: creatorId,
          userName: creatorName,
          joinedAt: new Date()
        }],
        createdAt: new Date(),
        lastActivity: new Date(),
        isActive: true
      };
      
      this.collaborationSessions.set(sessionId, session);
      this.activeCollaborations.set(creatorId, sessionId);
      
      logger.info('Створено сесію співпраці', {
        component: 'DocumentMentionHandler',
        sessionId,
        fileId,
        creatorId
      });
      
      return session;
    } catch (error) {
      logger.error('Помилка створення сесії співпраці', {
        component: 'DocumentMentionHandler',
        fileId,
        creatorId,
        error: error instanceof Error ? error.message : String(error)
      });
      
      throw error;
    }
  }

  /**
   * Приєднує користувача до сесії співпраці
   */
  joinCollaborationSession(sessionId: string, userId: string, userName: string): boolean {
    try {
      const session = this.collaborationSessions.get(sessionId);
      
      if (!session || !session.isActive) {
        return false;
      }
      
      // Перевіряємо чи користувач вже в сесії
      const isAlreadyParticipant = session.participants.some(p => p.userId === userId);
      
      if (!isAlreadyParticipant) {
        session.participants.push({
          userId,
          userName,
          joinedAt: new Date()
        });
      }
      
      session.lastActivity = new Date();
      this.activeCollaborations.set(userId, sessionId);
      
      logger.info('Користувач приєднався до сесії співпраці', {
        component: 'DocumentMentionHandler',
        sessionId,
        userId,
        userName
      });
      
      return true;
    } catch (error) {
      logger.error('Помилка приєднання до сесії співпраці', {
        component: 'DocumentMentionHandler',
        sessionId,
        userId,
        error: error instanceof Error ? error.message : String(error)
      });
      
      return false;
    }
  }

  /**
   * Виходить з сесії співпраці
   */
  leaveCollaborationSession(sessionId: string, userId: string): boolean {
    try {
      const session = this.collaborationSessions.get(sessionId);
      
      if (!session) {
        return false;
      }
      
      // Видаляємо користувача з учасників
      session.participants = session.participants.filter(p => p.userId !== userId);
      session.lastActivity = new Date();
      
      // Видаляємо з активних співпраць
      this.activeCollaborations.delete(userId);
      
      logger.info('Користувач вийшов з сесії співпраці', {
        component: 'DocumentMentionHandler',
        sessionId,
        userId
      });
      
      return true;
    } catch (error) {
      logger.error('Помилка виходу з сесії співпраці', {
        component: 'DocumentMentionHandler',
        sessionId,
        userId,
        error: error instanceof Error ? error.message : String(error)
      });
      
      return false;
    }
  }

  /**
   * Отримує активні сесії співпраці для користувача
   */
  getUserActiveSessions(userId: string): CollaborationSession[] {
    const userSessionId = this.activeCollaborations.get(userId);
    
    if (!userSessionId) {
      return [];
    }
    
    const session = this.collaborationSessions.get(userSessionId);
    
    if (session && session.isActive) {
      return [session];
    }
    
    return [];
  }

  /**
   * Додає реальний час редагування
   */
  addRealTimeEdit(
    fileId: string,
    userId: string,
    userName: string,
    action: 'insert' | 'delete' | 'replace',
    content: string,
    position: number
  ): void {
    try {
      const edit: RealTimeEdit = {
        userId,
        userName,
        action,
        content,
        position,
        timestamp: new Date()
      };
      
      // Отримуємо існуючі редагування для файлу
      let fileEdits = this.realTimeEdits.get(fileId) || [];
      
      // Додаємо нове редагування
      fileEdits.push(edit);
      
      // Обмежуємо кількість редагувань
      if (fileEdits.length > this.MAX_EDITS_PER_FILE) {
        fileEdits = fileEdits.slice(-this.MAX_EDITS_PER_FILE);
      }
      
      // Зберігаємо оновлені редагування
      this.realTimeEdits.set(fileId, fileEdits);
      
      logger.debug('Додано реальний час редагування', {
        component: 'DocumentMentionHandler',
        fileId,
        userId,
        action
      });
    } catch (error) {
      logger.error('Помилка додавання реального часу редагування', {
        component: 'DocumentMentionHandler',
        fileId,
        userId,
        error: error instanceof Error ? error.message : String(error)
      });
    }
  }

  /**
   * Отримує реальні редагування для документа
   */
  getRealTimeEdits(fileId: string): RealTimeEdit[] {
    return this.realTimeEdits.get(fileId) || [];
  }

  /**
   * Отримує назву типу файлу
   */
  private getMimeTypeLabel(mimeType: string): string {
    const labelMap: Record<string, string> = {
      'application/pdf': 'PDF документ',
      'application/vnd.google-apps.document': 'Google Docs',
      'application/vnd.google-apps.spreadsheet': 'Google Sheets',
      'application/vnd.google-apps.presentation': 'Google Slides',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'Word документ',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'Excel таблиця',
      'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'PowerPoint презентація',
      'text/plain': 'Текстовий файл',
      'image/': 'Зображення'
    };
    
    for (const [key, label] of Object.entries(labelMap)) {
      if (mimeType.startsWith(key) || mimeType.includes(key)) {
        return label;
      }
    }
    
    return 'Файл';
  }

  /**
   * Прикріплює документ до повідомлення
   */
  async attachDocument(message: Message, fileName: string): Promise<void> {
    try {
      if (!this.google) {
        throw new Error('GoogleService не ініціалізовано');
      }

      // Шукаємо документ
      const file = await this.findDocumentInDrive(fileName);
      
      if (!file) {
        await message.reply(`❌ Документ "${fileName}" не знайдено.`);
        return;
      }

      // Для демонстрації просто надсилаємо посилання
      // У реальній реалізації можна завантажити файл та прикріпити його
      const link = file.webViewLink || `https://drive.google.com/file/d/${file.id}/view`;
      
      await message.reply({
        content: `📎 **Прикріплений документ:**\n[${file.name}](${link})`,
        allowedMentions: { repliedUser: false }
      });
    } catch (error) {
      logger.error('Помилка прикріплення документа', {
        component: 'DocumentMentionHandler',
        fileName,
        error: error instanceof Error ? error.message : String(error)
      });
      
      await message.reply(`❌ Помилка прикріплення документа "${fileName}".`);
    }
  }

  /**
   * Швидко прикріплює останній згаданий документ
   */
  async attachLastMentionedDocument(message: Message): Promise<void> {
    try {
      // У спрощеній реалізації просто шукаємо останній документ
      // У реальній реалізації потрібно зберігати історію згадок
      
      await message.reply('Для демонстрації функції прикріплення останнього документа, будь ласка, вкажіть назву файлу.');
    } catch (error) {
      logger.error('Помилка прикріплення останнього документа', {
        component: 'DocumentMentionHandler',
        error: error instanceof Error ? error.message : String(error)
      });
      
      await message.reply('❌ Помилка прикріплення останнього документа.');
    }
  }

  /**
   * Генерує ID для нових об'єктів
   */
  private generateId(): string {
    return Date.now().toString(36) + Math.random().toString(36).substr(2);
  }

  // === BaseService required methods ===
  
  protected async onInitialize(): Promise<void> {
    logger.info('DocumentMentionHandler ініціалізовано', {
      component: 'DocumentMentionHandler'
    });
  }

  protected async onShutdown(): Promise<void> {
    logger.info('DocumentMentionHandler зупинено', {
      component: 'DocumentMentionHandler'
    });
  }

  protected async onHealthCheck(): Promise<any> {
    return {
      healthy: true,
      service: this.name,
      cacheSize: this.documentCache.size,
      annotationsCount: Array.from(this.annotations.values()).reduce((sum, arr) => sum + arr.length, 0),
      collaborationSessionsCount: this.collaborationSessions.size,
      activeCollaborationsCount: this.activeCollaborations.size
    };
  }

  protected onGetStats(): any {
    return {
      cacheSize: this.documentCache.size,
      annotationsCount: Array.from(this.annotations.values()).reduce((sum, arr) => sum + arr.length, 0),
      collaborationSessionsCount: this.collaborationSessions.size,
      activeCollaborationsCount: this.activeCollaborations.size
    };
  }
}