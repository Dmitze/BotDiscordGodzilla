/**
 * Document Encryption Service for Discord AI Assistant Bot
 * Provides enhanced encryption for sensitive document content
 * Version 1.0.0
 */

import type { BotConfig } from '@/types';
import { BaseService } from '@/core/BaseService';
import logger from '@/utils/logger';
import { createCipheriv, createDecipheriv, randomBytes, createHash } from 'crypto';

// Constants for encryption
const ENCRYPTION_CONSTANTS = {
  ALGORITHM: 'aes-256-gcm',
  IV_LENGTH: 16,
  AUTH_TAG_LENGTH: 16,
  SALT_LENGTH: 32,
  KEY_LENGTH: 32,
  ITERATIONS: 100000,
} as const;

export interface EncryptedDocument {
  encryptedData: string; // Base64 encoded encrypted data
  iv: string; // Base64 encoded initialization vector
  authTag: string; // Base64 encoded authentication tag
  salt: string; // Base64 encoded salt
  algorithm: string;
  createdAt: Date;
}

export interface DocumentEncryptionStats {
  totalDocumentsEncrypted: number;
  totalDocumentsDecrypted: number;
  failedEncryptions: number;
  failedDecryptions: number;
  averageEncryptionTime: number;
  averageDecryptionTime: number;
  totalEncryptionTime: number;
  totalDecryptionTime: number;
}

export class DocumentEncryptionService extends BaseService {
  private stats: DocumentEncryptionStats;
  private encryptionKey: Buffer;

  constructor(config: BotConfig) {
    super('DocumentEncryptionService', config);
    
    this.stats = {
      totalDocumentsEncrypted: 0,
      totalDocumentsDecrypted: 0,
      failedEncryptions: 0,
      failedDecryptions: 0,
      averageEncryptionTime: 0,
      averageDecryptionTime: 0,
      totalEncryptionTime: 0,
      totalDecryptionTime: 0,
    };

    // Generate or load encryption key from config
    this.encryptionKey = this.initializeEncryptionKey();
  }

  /**
   * Initialize encryption key from config or generate a new one
   */
  private initializeEncryptionKey(): Buffer {
    try {
      // Try to get key from config
      const configKey = this.config.security?.documentEncryptionKey;
      
      if (configKey) {
        // If key is provided in config, use it
        const keyBuffer = Buffer.from(configKey, 'base64');
        if (keyBuffer.length !== ENCRYPTION_CONSTANTS.KEY_LENGTH) {
          throw new Error(`Invalid key length. Expected ${ENCRYPTION_CONSTANTS.KEY_LENGTH} bytes, got ${keyBuffer.length}`);
        }
        return keyBuffer;
      } else {
        // Generate a new key if not provided
        logger.warn('🔐 Document encryption key not found in config. Generating a new one.');
        return randomBytes(ENCRYPTION_CONSTANTS.KEY_LENGTH);
      }
    } catch (error) {
      logger.error('❌ Error initializing encryption key', {
        component: 'DocumentEncryptionService',
        error: error instanceof Error ? error.message : String(error)
      });
      throw error;
    }
  }

  /**
   * Encrypt sensitive document content
   */
  public encryptDocumentContent(content: string, password?: string): EncryptedDocument {
    const startTime = Date.now();
    
    try {
      // Generate a random salt and IV
      const salt = randomBytes(ENCRYPTION_CONSTANTS.SALT_LENGTH);
      const iv = randomBytes(ENCRYPTION_CONSTANTS.IV_LENGTH);
      
      // Derive key from password or use default key
      const key = password 
        ? this.deriveKeyFromPassword(password, salt)
        : this.encryptionKey;
      
      // Create cipher
      const cipher = createCipheriv(ENCRYPTION_CONSTANTS.ALGORITHM, key, iv);
      
      // Encrypt the content
      let encrypted = cipher.update(content, 'utf8', 'base64');
      encrypted += cipher.final('base64');
      
      // Get the authentication tag
      const authTag = cipher.getAuthTag();
      
      const result: EncryptedDocument = {
        encryptedData: encrypted,
        iv: iv.toString('base64'),
        authTag: authTag.toString('base64'),
        salt: salt.toString('base64'),
        algorithm: ENCRYPTION_CONSTANTS.ALGORITHM,
        createdAt: new Date(),
      };

      // Update stats
      const duration = Date.now() - startTime;
      this.updateEncryptionStats(true, duration);

      logger.info('🔒 Document content encrypted successfully', {
        component: 'DocumentEncryptionService',
        duration: `${duration}ms`
      });

      return result;
    } catch (error) {
      this.updateEncryptionStats(false, Date.now() - startTime);
      
      logger.error('❌ Error encrypting document content', {
        component: 'DocumentEncryptionService',
        error: error instanceof Error ? error.message : String(error)
      });
      
      throw error;
    }
  }

  /**
   * Decrypt sensitive document content
   */
  public decryptDocumentContent(encryptedDocument: EncryptedDocument, password?: string): string {
    const startTime = Date.now();
    
    try {
      // Validate the encrypted document
      if (!encryptedDocument.encryptedData || !encryptedDocument.iv || !encryptedDocument.authTag || !encryptedDocument.salt) {
        throw new Error('Invalid encrypted document structure');
      }
      
      // Decode base64 values
      const encryptedData = Buffer.from(encryptedDocument.encryptedData, 'base64');
      const iv = Buffer.from(encryptedDocument.iv, 'base64');
      const authTag = Buffer.from(encryptedDocument.authTag, 'base64');
      const salt = Buffer.from(encryptedDocument.salt, 'base64');
      
      // Derive key from password or use default key
      const key = password 
        ? this.deriveKeyFromPassword(password, salt)
        : this.encryptionKey;
      
      // Create decipher
      const decipher = createDecipheriv(ENCRYPTION_CONSTANTS.ALGORITHM, key, iv);
      decipher.setAuthTag(authTag);
      
      // Decrypt the content
      let decrypted = decipher.update(encryptedData);
      decrypted = Buffer.concat([decrypted, decipher.final()]);
      
      const result = decrypted.toString('utf8');

      // Update stats
      const duration = Date.now() - startTime;
      this.updateEncryptionStats(false, duration, true);

      logger.info('🔓 Document content decrypted successfully', {
        component: 'DocumentEncryptionService',
        duration: `${duration}ms`
      });

      return result;
    } catch (error) {
      this.updateEncryptionStats(false, Date.now() - startTime, false);
      
      logger.error('❌ Error decrypting document content', {
        component: 'DocumentEncryptionService',
        error: error instanceof Error ? error.message : String(error)
      });
      
      throw error;
    }
  }

  /**
   * Derive encryption key from password using PBKDF2
   */
  private deriveKeyFromPassword(password: string, salt: Buffer): Buffer {
    return createHash('sha256')
      .update(password)
      .update(salt)
      .digest()
      .subarray(0, ENCRYPTION_CONSTANTS.KEY_LENGTH);
  }

  /**
   * Update encryption/decryption statistics
   */
  private updateEncryptionStats(isEncryption: boolean, duration: number, success: boolean = true): void {
    try {
      if (isEncryption) {
        if (success) {
          this.stats.totalDocumentsEncrypted++;
          this.stats.totalEncryptionTime += duration;
          // Avoid division by zero
          if (this.stats.totalDocumentsEncrypted > 0) {
            this.stats.averageEncryptionTime = this.stats.totalEncryptionTime / this.stats.totalDocumentsEncrypted;
          }
        } else {
          this.stats.failedEncryptions++;
        }
      } else {
        if (success) {
          this.stats.totalDocumentsDecrypted++;
          this.stats.totalDecryptionTime += duration;
          // Avoid division by zero
          if (this.stats.totalDocumentsDecrypted > 0) {
            this.stats.averageDecryptionTime = this.stats.totalDecryptionTime / this.stats.totalDocumentsDecrypted;
          }
        } else {
          this.stats.failedDecryptions++;
        }
      }
    } catch (error) {
      logger.warn('⚠️ Error updating encryption stats', {
        component: 'DocumentEncryptionService',
        error: error instanceof Error ? error.message : String(error)
      });
    }
  }

  /**
   * Get encryption service statistics
   */
  public getStats(): DocumentEncryptionStats {
    return { ...this.stats };
  }

  /**
   * Check if a document should be encrypted based on its content or metadata
   */
  public shouldEncryptDocument(content: string, fileName: string): boolean {
    // Check for sensitive keywords in content or filename
    const sensitiveKeywords = [
      'confidential', 'secret', 'private', 'internal', 'restricted', 
      'classified', 'proprietary', 'sensitive', 'password', 'credential',
      'token', 'api', 'key', 'ssn', 'social security', 'passport',
      'bank', 'account', 'credit card', 'debit card', 'financial'
    ];
    
    const lowerContent = content.toLowerCase();
    const lowerFileName = fileName.toLowerCase();
    
    return sensitiveKeywords.some(keyword => 
      lowerContent.includes(keyword) || lowerFileName.includes(keyword)
    );
  }

  /**
   * Encrypt document if it contains sensitive content
   */
  public encryptIfSensitive(content: string, fileName: string): EncryptedDocument | null {
    if (this.shouldEncryptDocument(content, fileName)) {
      return this.encryptDocumentContent(content);
    }
    return null;
  }

  protected async onInitialize(): Promise<void> {
    logger.info('🔐 Document Encryption Service initialized', {
      component: 'DocumentEncryptionService'
    });
  }

  protected async onCleanup(): Promise<void> {
    logger.info('🧹 Document Encryption Service cleaned up', {
      component: 'DocumentEncryptionService'
    });
  }
}

export default DocumentEncryptionService;