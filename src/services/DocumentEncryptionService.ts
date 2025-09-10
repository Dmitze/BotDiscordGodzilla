/**
 * Document Encryption Service for Discord AI Assistant Bot
 * Provides enhanced encryption for sensitive document content
 * Version 1.0.0
 */

import type { BotConfig, HealthStatus, ServiceStats } from '@/types';
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
  private stats: DocumentEncryptionStats = {
    totalDocumentsEncrypted: 0,
    totalDocumentsDecrypted: 0,
    failedEncryptions: 0,
    failedDecryptions: 0,
    averageEncryptionTime: 0,
    averageDecryptionTime: 0,
    totalEncryptionTime: 0,
    totalDecryptionTime: 0
  };
  private encryptionKey: string;
  
  constructor(config: BotConfig) {
    super('DocumentEncryptionService', config);
    
    // Use the document encryption key from security config, or generate a default one
    this.encryptionKey = config.security?.documentEncryptionKey || 
      this.generateDefaultKey();
  }

  /**
   * Generate a default encryption key
   */
  private generateDefaultKey(): string {
    return randomBytes(32).toString('base64');
  }

  /**
   * Initialize service
   */
  protected async onInitialize(): Promise<void> {
    // Implementation for initialization if needed
    logger.info('DocumentEncryptionService initialized', {
      component: 'DocumentEncryptionService'
    });
  }

  /**
   * Shutdown service
   */
  protected async onShutdown(): Promise<void> {
    // Implementation for shutdown if needed
    logger.info('DocumentEncryptionService shutdown', {
      component: 'DocumentEncryptionService'
    });
  }

  /**
   * Health check
   */
  protected async onHealthCheck(): Promise<HealthStatus> {
    return {
      healthy: true,
      service: 'DocumentEncryptionService'
    };
  }

  /**
   * Get service stats
   */
  protected onGetStats(): Partial<ServiceStats> {
    return {
      totalDocumentsEncrypted: this.stats.totalDocumentsEncrypted,
      totalDocumentsDecrypted: this.stats.totalDocumentsDecrypted,
      failedEncryptions: this.stats.failedEncryptions,
      failedDecryptions: this.stats.failedDecryptions
    };
  }

  /**
   * Get service statistics
   */
  public override getStats(): ServiceStats {
    // Get base stats from parent class
    const baseStats = super.getStats();
    
    return {
      ...baseStats,
      totalDocumentsEncrypted: this.stats.totalDocumentsEncrypted,
      totalDocumentsDecrypted: this.stats.totalDocumentsDecrypted,
      failedEncryptions: this.stats.failedEncryptions,
      failedDecryptions: this.stats.failedDecryptions,
      averageEncryptionTime: this.stats.averageEncryptionTime,
      averageDecryptionTime: this.stats.averageDecryptionTime,
      totalEncryptionTime: this.stats.totalEncryptionTime,
      totalDecryptionTime: this.stats.totalDecryptionTime
    };
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
      
      // Get encryption key (32 bytes for AES-256-GCM)
      const key = this.getEncryptionKey(password, salt);
      
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
      
      // Get encryption key (32 bytes for AES-256-GCM)
      const key = this.getEncryptionKey(password, salt);
      
      // Create decipher
      const decipher = createDecipheriv(ENCRYPTION_CONSTANTS.ALGORITHM, key, iv);
      decipher.setAuthTag(authTag);
      
      // Decrypt the content
      let decrypted = decipher.update(encryptedData);
      decrypted = Buffer.concat([decrypted, decipher.final()]);
      
      // Update stats (isEncryption = false for decryption)
      const duration = Date.now() - startTime;
      this.updateEncryptionStats(false, duration, true);

      logger.info('🔓 Document content decrypted successfully', {
        component: 'DocumentEncryptionService',
        duration: `${duration}ms`
      });

      return decrypted.toString('utf8');
    } catch (error) {
      // Update stats for failure (isEncryption = false for decryption)
      this.updateEncryptionStats(false, Date.now() - startTime, false);
      
      logger.error('❌ Error decrypting document content', {
        component: 'DocumentEncryptionService',
        error: error instanceof Error ? error.message : String(error)
      });
      
      throw error;
    }
  }

  /**
   * Derive a key from a password and salt
   */
  private deriveKeyFromPassword(password: string): Buffer {
    return createHash('sha256').update(password).digest();
  }

  /**
   * Get or derive encryption key
   */
  private getEncryptionKey(password?: string, salt?: Buffer): Buffer {
    if (password && salt) {
      return this.deriveKeyFromPassword(password);
    }
    
    // Ensure the key is exactly 32 bytes for AES-256-GCM
    const keyBuffer = Buffer.from(this.encryptionKey, 'base64');
    if (keyBuffer.length === 32) {
      return keyBuffer;
    }
    
    // If the key is not 32 bytes, derive a 32-byte key from it
    return createHash('sha256').update(keyBuffer).digest();
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
}

export default DocumentEncryptionService;